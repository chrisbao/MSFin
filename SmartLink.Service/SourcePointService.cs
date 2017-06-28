/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using AutoMapper;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json;
using SmartLink.Common;
using SmartLink.Entity;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public class SourceService : ISourceService
    {
        protected readonly SmartlinkDbContext _dbContext;
        protected readonly IMapper _mapper;
        protected readonly IAzureStorageService _azureStorageService;
        protected readonly ILogService _logService;
        protected readonly IUserProfileService _userProfileService;

        public SourceService(SmartlinkDbContext dbContext, IMapper mapper, IAzureStorageService azureStorageService, ILogService logService, IUserProfileService userProfileService)
        {
            _dbContext = dbContext;
            _mapper = mapper;
            _azureStorageService = azureStorageService;
            _logService = logService;
            _userProfileService = userProfileService;
        }

        /// <summary>
        /// Add source point in the Azure DB.
        /// The file name is the absolute path of the file.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sourcePoint"></param>
        /// <returns></returns>
        public async Task<SourcePoint> AddSourcePointAsync(string fileName, SourcePoint sourcePoint)
        {
            try
            {
                var sourceCatalog = _dbContext.SourceCatalogs.FirstOrDefault(o => o.Name == fileName);
                bool addSourceCatalog = (sourceCatalog == null);
                if (addSourceCatalog)
                {
                    try
                    {
                        sourceCatalog = new SourceCatalog() { Name = fileName };
                        _dbContext.SourceCatalogs.Add(sourceCatalog);
                    }
                    catch (Exception ex)
                    {
                        var entity = new LogEntity()
                        {
                            LogId = "30006",
                            Action = Constant.ACTIONTYPE_ADD,
                            ActionType = ActionTypeEnum.ErrorLog,
                            PointType = Constant.POINTTYPE_SOURCECATALOG,
                            Message = ".Net Error",
                        };
                        entity.Subject = $"{entity.LogId} - {entity.Action} - {entity.PointType} - Error";
                        await _logService.WriteLogAsync(entity);

                        throw new ApplicationException("Add Source Catalog failed", ex);
                    }
                }

                sourcePoint.Created = DateTime.Now.ToUniversalTime().ToPSTDateTime();
                sourcePoint.Creator = _userProfileService.GetCurrentUser().Username;

                sourceCatalog.SourcePoints.Add(sourcePoint);
                foreach (var groupId in sourcePoint.Groups)
                {
                    _dbContext.SourcePointGroups.Attach(groupId);
                }
                var history = _mapper.Map<PublishedHistory>(sourcePoint);
                sourcePoint.PublishedHistories.Add(history);

                await _dbContext.SaveChangesAsync();

                if (addSourceCatalog)
                {
                    await _logService.WriteLogAsync(new LogEntity()
                    {
                        LogId = "30003",
                        Action = Constant.ACTIONTYPE_ADD,
                        PointType = Constant.POINTTYPE_SOURCECATALOG,
                        ActionType = ActionTypeEnum.AuditLog,
                        Message = $"Add Source Catalog named {sourceCatalog.Name}."
                    });
                }
                await _logService.WriteLogAsync(new LogEntity()
                {
                    LogId = "10001",
                    Action = Constant.ACTIONTYPE_ADD,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    ActionType = ActionTypeEnum.AuditLog,
                    Message = $"Create source point named {sourcePoint.Name} in the location: {sourcePoint.Position}, value: {sourcePoint.Value} in the excel file named: {sourceCatalog.FileName} by {sourcePoint.Creator}"
                });

            }
            catch (ApplicationException ex)
            {
                throw ex.InnerException;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "10005",
                    Action = Constant.ACTIONTYPE_ADD,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLogAsync(logEntity);
                throw ex;
            }
            return sourcePoint;
        }

        /// <summary>
        /// Update the source point.
        /// </summary>
        /// <param name="groupIds"></param>
        /// <param name="sourcePoint"></param>
        /// <returns></returns>
        public async Task<SourcePoint> EditSourcePointAsync(int[] groupIds, SourcePoint sourcePoint)
        {
            try
            {
                var previousSourcePoint = await _dbContext.SourcePoints.Include(o => o.Catalog).Include(o => o.Groups).Include(o => o.PublishedHistories).FirstOrDefaultAsync(o => o.Id == sourcePoint.Id);

                var newGroup = _dbContext.SourcePointGroups.Where(o => groupIds.Contains(o.Id));

                var previousValue = new SourcePoint() { Name = sourcePoint.Name, Position = sourcePoint.Position, RangeId = sourcePoint.RangeId, Value = sourcePoint.Value };

                if (previousSourcePoint != null)
                {
                    previousSourcePoint.Name = sourcePoint.Name;
                    previousSourcePoint.Position = sourcePoint.Position;
                    previousSourcePoint.RangeId = sourcePoint.RangeId;
                    previousSourcePoint.Value = sourcePoint.Value;

                    //var newGroups = groupIds.Select(o => new SourcePointGroup() { Id = o });
                    var newGroups = _dbContext.SourcePointGroups.Where(o => groupIds.Contains(o.Id));
                    var deletingGroups = previousSourcePoint.Groups.Except(newGroups, new Comparer<SourcePointGroup>((x, y) => x.Id == y.Id)).ToList();
                    var addingCourses = newGroups.AsEnumerable().Except(previousSourcePoint.Groups, new Comparer<SourcePointGroup>((x, y) => x.Id == y.Id));

                    foreach (var group in deletingGroups)
                    {
                        previousSourcePoint.Groups.Remove(group);
                    }

                    foreach (var group in addingCourses)
                    {
                        if (_dbContext.Entry(group).State == EntityState.Detached)
                            _dbContext.SourcePointGroups.Attach(group);
                        previousSourcePoint.Groups.Add(group);
                    }
                }

                await _dbContext.SaveChangesAsync();
                await _logService.WriteLogAsync(new LogEntity()
                {
                    LogId = "10002",
                    Action = Constant.ACTIONTYPE_EDIT,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    ActionType = ActionTypeEnum.AuditLog,
                    Message = $"Edit source point by {_userProfileService.GetCurrentUser().Username} Previous value: source point named: {previousSourcePoint.Name} in the location: {previousSourcePoint.Position} value: {previousSourcePoint.Value} in the excel file named: {previousSourcePoint.Catalog.FileName} " +
                              $"Current value: source point named: {sourcePoint.Name} in the location {sourcePoint.Position} value: {sourcePoint.Value} in the excel file name: {previousSourcePoint.Catalog.FileName}"
                });

                return previousSourcePoint;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "10008",
                    Action = Constant.ACTIONTYPE_EDIT,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLogAsync(logEntity);
                throw ex;
            }
        }

        /// <summary>
        /// Delete source point by source point guid.
        /// </summary>
        /// <param name="sourcePointId"></param>
        /// <returns></returns>
        public async Task<int> DeleteSourcePointAsync(Guid sourcePointId)
        {
            try
            {
                var sourcePoint = _dbContext.SourcePoints.Include(o => o.Catalog).FirstOrDefault(o => o.Id == sourcePointId);
                if (sourcePoint == null)
                {
                    throw new NullReferenceException(string.Format("Sourcepoint: {0} is not existed", sourcePointId));
                }
                if (sourcePoint.Status == SourcePointStatus.Deleted)
                {
                    return await Task.FromResult<int>(0);
                }
                else
                {
                    sourcePoint.Status = SourcePointStatus.Deleted;
                }
                var task = await _dbContext.SaveChangesAsync();
                await _logService.WriteLogAsync(new LogEntity()
                {
                    LogId = "10003",
                    Action = Constant.ACTIONTYPE_DELETE,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    ActionType = ActionTypeEnum.AuditLog,
                    Message = $"Delete source point named: {sourcePoint.Name} in the excel file named: {sourcePoint.Catalog.FileName}"
                });
                return task;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "10006",
                    Action = Constant.ACTIONTYPE_DELETE,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLogAsync(logEntity);
                throw ex;
            }

        }

        /// <summary>
        /// Delete a bunch of source points by source points guid.
        /// </summary>
        /// <param name="selectedSourcePointIds"></param>
        /// <returns></returns>
        public async Task DeleteSelectedSourcePointAsync(IEnumerable<Guid> selectedSourcePointIds)
        {
            try
            {
                foreach (var sourcePointId in selectedSourcePointIds)
                {
                    var sourcePoint = _dbContext.SourcePoints.Include(o => o.Catalog).FirstOrDefault(o => o.Id == sourcePointId);
                    if (sourcePoint == null)
                    {
                        throw new NullReferenceException(string.Format("Sourcepoint: {0} is not existed", sourcePointId));
                    }
                    if (sourcePoint.Status != SourcePointStatus.Deleted)
                    {
                        sourcePoint.Status = SourcePointStatus.Deleted;
                    }
                    var task = await _dbContext.SaveChangesAsync();
                    await _logService.WriteLogAsync(new LogEntity()
                    {
                        LogId = "10003",
                        Action = Constant.ACTIONTYPE_DELETE,
                        PointType = Constant.POINTTYPE_SOURCEPOINT,
                        ActionType = ActionTypeEnum.AuditLog,
                        Message = $"Delete source point named: {sourcePoint.Name} in the excel file named: {sourcePoint.Catalog.FileName}"
                    });
                }
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "10006",
                    Action = Constant.ACTIONTYPE_DELETE,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLogAsync(logEntity);
                throw ex;
            }

        }

        /// <summary>
        /// get the source catalog by file name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public async Task<SourceCatalog> GetSourceCatalogAsync(string fileName)
        {
            try
            {
                var sourceCatalog = await _dbContext.SourceCatalogs.Where(o => o.Name == fileName).FirstOrDefaultAsync();
                if (sourceCatalog != null)
                {
                    var sourcePoint = await _dbContext.SourcePoints.Where(o => o.Status == SourcePointStatus.Created && o.CatalogId == sourceCatalog.Id)
                        .Include(o => o.DestinationPoints)
                        .Include(o => o.Groups)
                        .Include(o => o.PublishedHistories).OrderByDescending(o => o.Name).ToArrayAsync();
                    foreach (var item in sourcePoint)
                    {
                        item.PublishedHistories = item.PublishedHistories.OrderByDescending(p => p.PublishedDate).ToArray();
                    }
                    sourceCatalog.SourcePoints = sourcePoint;

                    await _logService.WriteLogAsync(new LogEntity()
                    {
                        LogId = "30001",
                        Action = Constant.ACTIONTYPE_GET,
                        PointType = Constant.POINTTYPE_SOURCECATALOG,
                        ActionType = ActionTypeEnum.AuditLog,
                        Message = $"Get source catalog named {sourceCatalog.Name}"
                    });
                }
                return sourceCatalog;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "30004",
                    Action = Constant.ACTIONTYPE_GET,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCECATALOG,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLogAsync(logEntity);
                throw ex;
            }
        }

        /// <summary>
        /// Get all source catalogs
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<SourceCatalog>> GetAllSourceCatalogAsync()
        {
            var sourceCatalog = await _dbContext.SourceCatalogs.Include(o => o.SourcePoints.Select(m => m.Groups)).ToListAsync();
            Parallel.ForEach(sourceCatalog, o => o.SourcePoints = o.SourcePoints.Where(m => m.Status == SourcePointStatus.Created).ToArray());
            return sourceCatalog;
        }

        /// <summary>
        /// Publish source points 
        /// Download the word file and update the destination point value with source point value in word file then upload the word file to overwrite existed one.
        /// 
        /// </summary>
        /// <param name="publishSourcePointForms"></param>
        /// <returns></returns>
        public async Task<PublishSourcePointResult> PublishSourcePointListAsync(IEnumerable<PublishSourcePointForm> publishSourcePointForms)
        {
            try
            {
                var sourcePointIdList = publishSourcePointForms.Select(o => o.SourcePointId).ToArray();
                var sourcePointList = _dbContext.SourcePoints.Include(o => o.Catalog).Include(o => o.PublishedHistories).Where(o => sourcePointIdList.Contains(o.Id)).ToList();
                var currentUser = _userProfileService.GetCurrentUser();

                //Update database
                IList<PublishedHistory> histories = new List<PublishedHistory>();
                foreach (var sourcePoint in sourcePointList)
                {
                    sourcePoint.Value = publishSourcePointForms.First(o => o.SourcePointId == sourcePoint.Id).CurrentValue;
                    sourcePoint.Position = publishSourcePointForms.First(o => o.SourcePointId == sourcePoint.Id).Position;
                    var history = _mapper.Map<PublishedHistory>(sourcePoint);
                    history.PublishedDate = DateTime.Now.ToUniversalTime().ToPSTDateTime();
                    history.PublishedUser = currentUser.Username;
                    sourcePoint.PublishedHistories.Add(history);
                    histories.Add(history);
                }
                await _dbContext.SaveChangesAsync();

                //Update Table Storage
                var table = _azureStorageService.GetTable(Constant.PUBLISH_TABLE_NAME);

                var batchId = Guid.NewGuid();

                //A single batch operation can include up to 100 entities, so seperate histories into several batch when histories 
                //https://docs.microsoft.com/en-us/azure/storage/storage-dotnet-how-to-use-tables
                var batchCount = Math.Ceiling((float)histories.Count() / Constant.AZURETABLE_BATCH_COUNT);
                var batchTasks = new List<Task<IList<TableResult>>>();
                for (int i = 0; i < batchCount; i++)
                {
                    var historiesPerBatch = histories.Skip(Constant.AZURETABLE_BATCH_COUNT * i).Take(Constant.AZURETABLE_BATCH_COUNT);
                    var batchOpt = new TableBatchOperation();
                    foreach (var item in historiesPerBatch)
                    {
                        batchOpt.Insert(new PublishStatusEntity(batchId.ToString(), item.SourcePointId.ToString(), item.Id.ToString()));
                    }
                    batchTasks.Add(table.ExecuteBatchAsync(batchOpt));
                }

                Task.WaitAll(batchTasks.ToArray());
                IList<SourcePoint> errorSourcePoints = new List<SourcePoint>();
                var batchResults = batchTasks.SelectMany(o => o.Result).ToArray();
                for (int i = 0; i < batchResults.Count(); i++)
                {
                    if (!(batchResults[i].HttpStatusCode == 200 || batchResults[i].HttpStatusCode == 204))
                    {
                        errorSourcePoints.Add(sourcePointList[i]);
                    }
                }
                if (errorSourcePoints.Count() > 0)
                {
                    throw new SourcePointException(sourcePointList, string.Concat(batchResults.Where(o => !(o.HttpStatusCode == 200 || o.HttpStatusCode == 204)).Select(o => o.Result)), null);
                }

                //Push message to queue
                IList<Task> writeQueueTask = new List<Task>();
                foreach (var history in histories)
                {
                    var message = new PublishedMessage() { PublishHistoryId = history.Id, SourcePointId = history.SourcePointId, PublishBatchId = batchId };
                    writeQueueTask.Add(Task.Run(async () =>
                       {
                           await _azureStorageService.WriteMessageToQueueAsync(JsonConvert.SerializeObject(message), Constant.PUBLISH_QUEUE_NAME);
                           await _logService.WriteLogAsync(new LogEntity()
                           {
                               LogId = "10004",
                               Action = Constant.ACTIONTYPE_PUBLISH,
                               ActionType = ActionTypeEnum.AuditLog,
                               PointType = Constant.POINTTYPE_SOURCEPOINT,
                               Message = $"Publish source point named: {history.Name} in the location {history.Position}, value: {history.Value} in the excel file named:{history.SourcePoint.Catalog.Name} by {currentUser.Username}"
                           });
                       }));
                }
                await Task.WhenAll(writeQueueTask);

                foreach (var item in sourcePointList)
                {
                    item.PublishedHistories = item.PublishedHistories.OrderByDescending(p => p.PublishedDate).ToArray();
                }

                return new PublishSourcePointResult() { BatchId = batchId, SourcePoints = sourcePointList };
            }
            catch (SourcePointException ex)
            {
                foreach (var sourcePoint in ex.ErrorSourcePoints)
                {
                    var logEntity = new LogEntity()
                    {
                        LogId = "10007",
                        Action = Constant.ACTIONTYPE_PUBLISH,
                        ActionType = ActionTypeEnum.ErrorLog,
                        PointType = Constant.POINTTYPE_SOURCEPOINT,
                        Message = ".Net Error",
                        Detail = $"{sourcePoint.Id}-{sourcePoint.Name}-{ex.ToString()}"
                    };
                    logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                    await _logService.WriteLogAsync(logEntity);
                }

                throw ex;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "10007",
                    Action = Constant.ACTIONTYPE_PUBLISH,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_SOURCEPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLogAsync(logEntity);
                throw ex;
            }
        }

        /// <summary>
        /// Get all sourcec point groups.
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<SourcePointGroup>> GetAllSourcePointGroupAsync()
        {
            return await _dbContext.SourcePointGroups.ToListAsync();
        }

        /// <summary>
        /// Get the status of all publish hisotries by batchID. 
        /// </summary>
        /// <param name="batchId"></param>
        /// <returns></returns>
        public IEnumerable<PublishStatusEntity> GetPublishStatus(string batchId)
        {
            var table = _azureStorageService.GetTable(Constant.PUBLISH_TABLE_NAME);
            return table.ExecuteQuery(new TableQuery<PublishStatusEntity>()
            {
                FilterString = $"PartitionKey eq '{batchId}'"
            });

        }

        /// <summary>
        /// Get first publish history by publish history ID.
        /// </summary>
        /// <param name="publishHistoryId"></param>
        /// <returns></returns>
        public async Task<PublishedHistory> GetPublishHistoryByIdAsync(Guid publishHistoryId)
        {
            var publishHistory = await _dbContext.PublishedHistories.Include(o => o.SourcePoint.Catalog).FirstOrDefaultAsync(o => o.Id == publishHistoryId);
            if (publishHistory != null)
            {
                publishHistory.SourcePoint.SerializeCatalog = true;
                publishHistory.SourcePoint.Catalog.SerializeSourcePoints = false;
            }
            return publishHistory;
        }
    }

    class SourcePointException : Exception
    {
        public IList<SourcePoint> ErrorSourcePoints { get; protected set; }

        public SourcePointException(IList<SourcePoint> errorSourcePoints, string message, Exception innerException) : base(message, innerException)
        {
            ErrorSourcePoints = errorSourcePoints;
        }
    }
}