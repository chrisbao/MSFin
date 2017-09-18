/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using AutoMapper;
using ContosoO365DocSync.Entity;
using ContosoO365DocSync.Service;
using ContosoO365DocSync.Web.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace ContosoO365DocSync.Web.Controllers
{
    [APIAuthorize]
    public class SourcePointController : ApiController
    {
        protected readonly ISourceService _sourceService;
        protected readonly IMapper _mapper;

        public SourcePointController(ISourceService sourceService, IMapper mapper)
        {
            _sourceService = sourceService;
            _mapper = mapper;
        }

        /// <summary>
        /// Get all source point group
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("api/SourcePointGroup")]
        public async Task<IEnumerable<SourcePointGroup>> GetAllSourceGroupAsync()
        {
            return await _sourceService.GetAllSourcePointGroupAsync();
        }

        /// <summary>
        /// Add source point
        /// </summary>
        /// <param name="sourcePointAdded"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("api/SourcePoint")]

        public async Task<IHttpActionResult> PostAsync([FromBody]SourcePointForm sourcePointAdded)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                var sourcePoint = _mapper.Map<SourcePoint>(sourcePointAdded);
                var catalogName = HttpUtility.UrlDecode(sourcePointAdded.CatalogName);
                var documentId = HttpUtility.UrlDecode(sourcePointAdded.DocumentId);
                return Ok(await _sourceService.AddSourcePointAsync(catalogName, documentId, sourcePoint));
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        /// <summary>
        /// Get source point.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="documentId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("api/SourcePointCatalog")]
        public async Task<IHttpActionResult> GetSourcePointCatalog(string name, string documentId)
        {
            var retValue = await _sourceService.GetSourceCatalogAsync(HttpUtility.UrlDecode(name), HttpUtility.UrlDecode(documentId));
            return Ok(retValue);
        }

        /// <summary>
        /// get source catalog by document id.
        /// </summary>
        /// <param name="documentId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("api/SourcePointCatalog")]
        public async Task<IHttpActionResult> GetSourcePointCatalogAsync(string documentId)
        {
            var retValue = await _sourceService.GetSourceCatalogAsync(HttpUtility.UrlDecode(documentId));
            return Ok(retValue);
        }

        /// <summary>
        /// get all source catalogs
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("api/SourcePointCatalog")]
        public async Task<IHttpActionResult> GetAllSourcePointCatalogAsync()
        {
            var retValue = await _sourceService.GetAllSourceCatalogAsync();
            return Ok(retValue);
        }

        /// <summary>
        /// get publish status by batchId
        /// If there is any error in the batch process, then return the error.
        /// If there is any item still processing in the batch process, then return InProgress.
        /// Otherwise return completed.
        /// </summary>
        /// <param name="batchId"></param>
        /// <returns></returns>
        [HttpGet]
        [Route("api/PublishStatus")]
        public IHttpActionResult GetPublishStatus(string batchId)
        {
            string publishStatusCode = string.Empty;
            var publishStatus = _sourceService.GetPublishStatus(batchId).ToArray();
            if (publishStatus.Any(o => o.Status == PublishStatus.Error))
                publishStatusCode = "Error";
            else if (publishStatus.Any(o => o.Status == PublishStatus.InProgess))
                publishStatusCode = "InProgess";
            else
                publishStatusCode = "Completed";

            var retValue = new PublishStatusViewModel()
            {
                Status = publishStatusCode,
                SourcePoints = publishStatus.Select(o => new PublishItemViewModel() { Id = o.SourcePointId, Status = o.Status.ToString(), Message = o.ErrorSummary }).ToArray()
            };
            return Ok(retValue);
        }

        /// <summary>
        /// publish a list of source points.
        /// </summary>
        /// <param name="sourcePointPublishForm"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("api/PublishSourcePoints")]
        public async Task<IHttpActionResult> PublishSourcePointsAsync([FromBody]IEnumerable<PublishSourcePointForm> sourcePointPublishForm)
        {
            if (!ModelState.IsValid || sourcePointPublishForm.Count() == 0)
            {
                return BadRequest("Invalid posted data.");
            }

            var retValue = await _sourceService.PublishSourcePointListAsync(sourcePointPublishForm);

            return Ok(retValue);
        }

        /// <summary>
        /// Edit the source point.
        /// </summary>
        /// <param name="sourcePointAdded"></param>
        /// <returns></returns>
        [HttpPut]
        [Route("api/SourcePoint")]
        public async Task<IHttpActionResult> EditSourcePointAsync([FromBody]SourcePointForm sourcePointAdded)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                var sourcePoint = _mapper.Map<SourcePoint>(sourcePointAdded);
                return Ok(await _sourceService.EditSourcePointAsync(sourcePointAdded.GroupIds != null ? sourcePointAdded.GroupIds : new int[] { }, sourcePoint));
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        /// <summary>
        /// Delete source point by source point GUID.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [HttpDelete]
        [Route("api/SourcePoint")]
        public async Task<IHttpActionResult> DeleteSourcePointAsync(string id)
        {
            var retValue = await _sourceService.DeleteSourcePointAsync(new Guid(id));
            return Ok();
        }

        /// <summary>
        /// Delete selected source points by source point guids.
        /// </summary>
        /// <param name="seletedIds"></param>
        /// <returns></returns>
        [HttpPost]
        [Route("api/DeleteSelectedSourcePoint")]
        public async Task<IHttpActionResult> DeleteSelectedSourcePointAsync([FromBody]IEnumerable<Guid> seletedIds)
        {
            await _sourceService.DeleteSelectedSourcePointAsync(seletedIds);
            return Ok();
        }
    }
}