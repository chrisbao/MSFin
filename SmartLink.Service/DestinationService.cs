/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using AutoMapper;
using SmartLink.Common;
using SmartLink.Entity;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public class DestinationService : IDestinationService
    {
        protected readonly SmartlinkDbContext _dbContext;
        protected readonly IMapper _mapper;
        protected readonly IAzureStorageService _azureStorageService;
        protected readonly ILogService _logService;
        protected readonly IUserProfileService _userProfileService;

        public DestinationService(SmartlinkDbContext dbContext, IMapper mapper, IAzureStorageService azureStorageService, ILogService logService, IUserProfileService userProfileService)
        {
            _dbContext = dbContext;
            _mapper = mapper;
            _azureStorageService = azureStorageService;
            _logService = logService;
            _userProfileService = userProfileService;
        }
        /// <summary>
        /// Add destination point to the Azure DB.
        /// file name is the absolute path of the file.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="destinationPoint"></param>
        /// <returns></returns>
        public async Task<DestinationPoint> AddDestinationPoint(string fileName, DestinationPoint destinationPoint)
        {
            try
            {

                ///Get destination catalog by file server absolute path.
                var destinationCatalog = _dbContext.DestinationCatalogs.FirstOrDefault(o => o.Name == fileName);
                bool addDestinationCatalog = (destinationCatalog == null);
                if (addDestinationCatalog)
                {
                    try
                    {
                        destinationCatalog = new DestinationCatalog() { Name = fileName };
                        _dbContext.DestinationCatalogs.Add(destinationCatalog);
                    }
                    catch (Exception ex)
                    {
                        var entity = new LogEntity()
                        {
                            LogId = "40002",
                            Action = Constant.ACTIONTYPE_ADD,
                            ActionType = ActionTypeEnum.ErrorLog,
                            PointType = Constant.POINTYTPE_DESTINATIONCATALOG,
                            Message = ".Net Error",
                        };
                        entity.Subject = $"{entity.LogId} - {entity.Action} - {entity.PointType} - Error";
                        await _logService.WriteLog(entity);

                        throw new ApplicationException("Add Source Catalog failed", ex);
                    }
                }

                destinationPoint.Created = DateTime.Now.ToUniversalTime().ToPSTDateTime();
                destinationPoint.Creator = _userProfileService.GetCurrentUser().Username;

                destinationCatalog.DestinationPoints.Add(destinationPoint);

                _dbContext.SourcePoints.Attach(destinationPoint.ReferencedSourcePoint);
                _dbContext.DestinationPoints.Add(destinationPoint);

                foreach (var formatId in destinationPoint.CustomFormats)
                {
                    _dbContext.CustomFormats.Attach(formatId);
                }

                await _dbContext.SaveChangesAsync();
                await _dbContext.Entry(destinationPoint.ReferencedSourcePoint).ReloadAsync();
                foreach (var customFormatItem in destinationPoint.CustomFormats)
                {
                    await _dbContext.Entry(customFormatItem).ReloadAsync();
                }

                if (addDestinationCatalog)
                {
                    await _logService.WriteLog(new LogEntity()
                    {
                        LogId = "40001",
                        Action = Constant.ACTIONTYPE_ADD,
                        PointType = Constant.POINTYTPE_DESTINATIONCATALOG,
                        ActionType = ActionTypeEnum.AuditLog,
                        Message = $"Add destination catalog {destinationCatalog.Name}."
                    });
                }
                await _logService.WriteLog(new LogEntity()
                {
                    LogId = "20001",
                    Action = Constant.ACTIONTYPE_ADD,
                    PointType = Constant.POINTTYPE_DESTINATIONPOINT,
                    ActionType = ActionTypeEnum.AuditLog,
                    Message = $"Create destination point value: {destinationPoint.ReferencedSourcePoint.Value} in the word file named:{destinationCatalog.FileName} by {_userProfileService.GetCurrentUser().Username}"
                });

                destinationPoint = await _dbContext.DestinationPoints
                    .Include(o => o.ReferencedSourcePoint)
                    .Include(o => o.ReferencedSourcePoint.Catalog)
                    .Include(o => o.ReferencedSourcePoint.PublishedHistories)
                    .Include(o => o.CustomFormats).FirstAsync(o => o.Id == destinationPoint.Id);
                destinationPoint.ReferencedSourcePoint.PublishedHistories = destinationPoint.ReferencedSourcePoint.PublishedHistories.OrderByDescending(p => p.PublishedDate).ToArray();
                destinationPoint.ReferencedSourcePoint.SerializeCatalog = true;
                destinationPoint.ReferencedSourcePoint.Catalog.SerializeSourcePoints = false;

                await _logService.WriteLog(new LogEntity()
                {
                    LogId = "30002",
                    Action = Constant.ACTIONTYPE_GET,
                    PointType = Constant.POINTTYPE_SOURCECATALOGLIST,
                    ActionType = ActionTypeEnum.AuditLog,
                    Message = $"Get source catalogs."
                });
                return destinationPoint;
            }
            catch (ApplicationException ex)
            {
                throw ex.InnerException;
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "20003",
                    Action = Constant.ACTIONTYPE_ADD,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_DESTINATIONPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw ex;
            }
        }
        /// <summary>
        /// Delete destination point by destination point guid.
        /// </summary>
        /// <param name="destinationPointId"></param>
        /// <returns></returns>
        public async Task DeleteDestinationPoint(Guid destinationPointId)
        {
            try
            {
                var destinationPoint = await _dbContext.DestinationPoints
                    .Include(o => o.ReferencedSourcePoint)
                    .Include(o => o.Catalog).FirstOrDefaultAsync(o => o.Id == destinationPointId);

                if (destinationPoint != null)
                {
                    var logResult = new
                    {
                        location = destinationPoint.RangeId,
                        value = destinationPoint.ReferencedSourcePoint.Value,
                        fileName = destinationPoint.Catalog.FileName,
                        user = _userProfileService.GetCurrentUser().Username
                    };

                    _dbContext.DestinationPoints.Remove(destinationPoint);
                    await _dbContext.SaveChangesAsync();
                    await _logService.WriteLog(new LogEntity()
                    {
                        LogId = "20002",
                        Action = Constant.ACTIONTYPE_DELETE,
                        PointType = Constant.POINTTYPE_DESTINATIONPOINT,
                        ActionType = ActionTypeEnum.AuditLog,
                        Message = $"Delete destination point in the location {logResult.location}, value: {logResult.value} in the word file named:{logResult.fileName} by {logResult.user}"
                    });
                }
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "20003",
                    Action = Constant.ACTIONTYPE_DELETE,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_DESTINATIONPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw;
            }
        }
        /// <summary>
        /// Delete a bunch of destination points by guids.
        /// </summary>
        /// <param name="seletedDestinationPointIds"></param>
        /// <returns></returns>
        public async Task DeleteSelectedDestinationPoint(IEnumerable<Guid> seletedDestinationPointIds)
        {
            try
            {
                foreach (var destinationPointId in seletedDestinationPointIds)
                {
                    var destinationPoint = await _dbContext.DestinationPoints
                        .Include(o => o.ReferencedSourcePoint)
                        .Include(o => o.Catalog).FirstOrDefaultAsync(o => o.Id == destinationPointId);

                    if (destinationPoint != null)
                    {
                        var logResult = new
                        {
                            location = destinationPoint.RangeId,
                            value = destinationPoint.ReferencedSourcePoint.Value,
                            fileName = destinationPoint.Catalog.FileName,
                            user = _userProfileService.GetCurrentUser().Username
                        };

                        _dbContext.DestinationPoints.Remove(destinationPoint);
                        await _dbContext.SaveChangesAsync();
                        await _logService.WriteLog(new LogEntity()
                        {
                            LogId = "20002",
                            Action = Constant.ACTIONTYPE_DELETE,
                            PointType = Constant.POINTTYPE_DESTINATIONPOINT,
                            ActionType = ActionTypeEnum.AuditLog,
                            Message = $"Delete destination point in the location {logResult.location}, value: {logResult.value} in the word file named:{logResult.fileName} by {logResult.user}"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "20003",
                    Action = Constant.ACTIONTYPE_DELETE,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_DESTINATIONPOINT,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw;
            }
        }
        /// <summary>
        /// Get destination catalog by file name
        /// File name is the absolute path of the file.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public async Task<DestinationCatalog> GetDestinationCatalog(string fileName)
        {
            try
            {
                var destinationCatalog = await _dbContext.DestinationCatalogs
                    .Include(o => o.DestinationPoints.Select(m => m.ReferencedSourcePoint.PublishedHistories))
                    .Include(o => o.DestinationPoints.Select(m => m.ReferencedSourcePoint.Groups))
                    .Include(o => o.DestinationPoints.Select(m => m.ReferencedSourcePoint.Catalog))
                    .Include(o => o.DestinationPoints.Select(m => m.CustomFormats))
                    .FirstOrDefaultAsync(o => o.Name == fileName);
                if (destinationCatalog != null)
                {
                    foreach (var sourcePoint in destinationCatalog.DestinationPoints.Select(o => o.ReferencedSourcePoint))
                    {
                        sourcePoint.PublishedHistories = sourcePoint.PublishedHistories.OrderByDescending(p => p.PublishedDate).ToArray();
                        sourcePoint.SerializeCatalog = true;
                        sourcePoint.Catalog.SerializeSourcePoints = false;
                    }
                }

                await _logService.WriteLog(new LogEntity()
                {
                    LogId = "20005",
                    Action = Constant.ACTIONTYPE_GET,
                    PointType = Constant.POINTTYPE_DESTINATIONPOINT,
                    ActionType = ActionTypeEnum.AuditLog,
                    Message = $"Get destination points list by {fileName}"
                });
                return destinationCatalog;

            }
            catch (Exception ex)
            {
                var logEntity = new LogEntity()
                {
                    LogId = "20006",
                    Action = Constant.ACTIONTYPE_GET,
                    ActionType = ActionTypeEnum.ErrorLog,
                    PointType = Constant.POINTTYPE_DESTINATIONLIST,
                    Message = ".Net Error",
                    Detail = ex.ToString()
                };
                logEntity.Subject = $"{logEntity.LogId} - {logEntity.Action} - {logEntity.PointType} - Error";
                await _logService.WriteLog(logEntity);
                throw ex;
            }
        }
        /// <summary>
        /// Get the destination point references by source point guid.
        /// </summary>
        /// <param name="sourcePointId"></param>
        /// <returns></returns>
        public async Task<IEnumerable<DestinationPoint>> GetDestinationPointBySourcePoint(Guid sourcePointId)
        {
            var destinationPoints = await _dbContext.DestinationPoints
                .Include(o => o.Catalog)
                .Include(o => o.CustomFormats)
                .Where(o => o.SourcePointId == sourcePointId).ToArrayAsync();
            foreach (var destinationPoint in destinationPoints)
            {
                destinationPoint.SerializeCatalog = true;
                destinationPoint.Catalog.SerializeDestinationPoints = false;
            }
            return destinationPoints;
        }
        /// <summary>
        /// Get all custom formats
        /// </summary>
        /// <returns></returns>
        public async Task<IEnumerable<CustomFormat>> GetCustomFormats()
        {
            var customFormats = await _dbContext.CustomFormats.ToArrayAsync();
            return customFormats;
        }
    }
}
