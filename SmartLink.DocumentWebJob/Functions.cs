/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Microsoft.Azure.WebJobs;
using Microsoft.WindowsAzure.Storage.Table;
using SmartLink.Service;
using SmartLink.Common;
using SmartLink.Entity;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SmartLink.DocumentWebJob
{
    public class Functions
    {
        private readonly ILogService _logService;
        private readonly IAzureStorageService _azureStorageService;
        private readonly ISourceService _sourceService;
        private readonly IDestinationService _destinationService;
        private readonly IDocumentService _documentService;

        public Functions(ILogService logService, ISourceService sourceService, IDestinationService destinationService, IAzureStorageService azureStorageService, IDocumentService documentService)
        {
            _logService = logService;
            _azureStorageService = azureStorageService;
            _sourceService = sourceService;
            _destinationService = destinationService;
            _documentService = documentService;
        }

        /// <summary>
        /// Update the document url by document id.
        /// </summary>
        /// <param name="log"></param>
        /// <returns></returns>
        [NoAutomaticTrigger]
        public async Task UpdateDocumentUrlByIdAsync(TextWriter log)
        {
            try
            {
                var table = _azureStorageService.GetTable(Constant.CHECK_TABLE_NAME);
                var retValue = new CheckDocumentEntity();
                List<Task<DocumentCheckResult>> tasks = new List<Task<DocumentCheckResult>>();
                var catalogs = await _sourceService.GetAllCatalogsAsync();
                log.WriteLine("Start search file url.");
                foreach (var catalog in catalogs)
                {
                    try
                    {
                        tasks.Add(_documentService.GetDocumentUrlByIdAsync(catalog));
                    }
                    catch (Exception ex)
                    {
                        log.Write($"Get the file url by ID'{catalog.DocumentId}' failed due to {ex.ToString()}");
                    }
                }
                Task.WaitAll(tasks.ToArray());

                var documents = tasks.Where(o => o.Result.IsSuccess).Select(o => o.Result);
                log.WriteLine($"End search file url, file size:{documents.Count()}");
                log.WriteLine(documents.Count() > 0 ? String.Join("\n , ", documents.Select(o => o.DocumentUrl)) : "No file url");
                log.WriteLine("Start check and update file url to DB.");
                var updateResult = (await _sourceService.UpdateDocumentUrlByIdAsync(documents)).Where(o => o.IsUpdated);
                if (updateResult.Count() > 0)
                {
                    log.WriteLine($"{updateResult.Count()} files updated.");
                }
                string comments = updateResult.Count() > 0 ? String.Join("\n , ", updateResult.Select(o => o.Message)) : "No files need to be updated.";
                retValue.Comments = comments;
                log.WriteLine($"{comments}");
                table.Execute(TableOperation.InsertOrReplace(retValue));
            }
            catch (Exception ex)
            {
                log.WriteLine($"Update the document url failed due to {ex.ToString()}");
            }
        }
    }
}
