/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using AutoMapper;
using SmartLink.Entity;
using SmartLink.Service;
using SmartLink.Web.Common;
using SmartLink.Web.ViewModel;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace SmartLink.Web.Controllers
{
    [APIAuthorize]
    public class DestinationPointController : ApiController
    {
        protected readonly IDestinationService _destinationService;
        protected readonly IMapper _mapper;
        public DestinationPointController(IDestinationService destinationService, IMapper mapper)
        {
            _destinationService = destinationService;
            _mapper = mapper;
        }

        [HttpPost]
        [Route("api/DestinationPoint")]
        public async Task<IHttpActionResult> Post([FromBody]DestinationPointForm destinationPointAdded)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                var destinationPoint = _mapper.Map<DestinationPoint>(destinationPointAdded);
                var catalogName = HttpUtility.UrlDecode(destinationPointAdded.CatalogName);
                var documentId = HttpUtility.UrlDecode(destinationPointAdded.DocumentId);
                return Ok(await _destinationService.AddDestinationPointAsync(catalogName, documentId, destinationPoint));
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        [HttpDelete]
        [Route("api/DestinationPoint")]
        public async Task<IHttpActionResult> DeleteSourcePoint(string id)
        {
            await _destinationService.DeleteDestinationPointAsync(Guid.Parse(id));
            return Ok();
        }

        [HttpPost]
        [Route("api/DeleteSelectedDestinationPoint")]
        public async Task<IHttpActionResult> DeleteSelectedDestinationPoint([FromBody]IEnumerable<Guid> seletedIds)
        {
            await _destinationService.DeleteSelectedDestinationPointAsync(seletedIds);
            return Ok();
        }

        [HttpGet]
        [Route("api/DestinationPointCatalog")]
        public async Task<IHttpActionResult> GetDestinationPointCatalog(string name, string documentId)
        {
            var retValue = await _destinationService.GetDestinationCatalogAsync(HttpUtility.UrlDecode(name), documentId);
            return Ok(retValue);
        }

        [HttpGet]
        [Route("api/DestinationPoint")]
        public async Task<IHttpActionResult> GetDestinationPointBySourcePoint(string sourcePointId)
        {
            var retValue = await _destinationService.GetDestinationPointBySourcePointAsync(Guid.Parse(sourcePointId));
            return Ok(retValue);
        }

        [HttpGet]
        [Route("api/GraphAccessToken")]
        public async Task<IHttpActionResult> GetGraphAccessToken()
        {
            var retValue = await AuthenticationHelper.AcquireTokenAsync();
            return Ok(retValue);
        }

        [HttpGet]
        [Route("api/CustomFormats")]
        public async Task<IHttpActionResult> GetCustomFormats()
        {
            var retValue = await _destinationService.GetCustomFormatsAsync();
            return Ok(retValue);
        }

        [HttpPut]
        [Route("api/UpdateDestinationPointCustomFormat")]
        public async Task<IHttpActionResult> UpdateDestinationPointCustomFormat([FromBody]DestinationPointForm destinationPointAdded)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest("Invalid posted data.");
            }

            try
            {
                var destinationPoint = _mapper.Map<DestinationPoint>(destinationPointAdded);
                return Ok(await _destinationService.UpdateDestinationPointCustomFormatAsync(destinationPoint));
            }
            catch (Exception ex)
            {
                return BadRequest(ex.ToString());
            }
        }

        [HttpGet]
        [Route("api/SharePointAccessToken")]
        public async Task<IHttpActionResult> GetSharePointAccessToken()
        {
            var retValue = await AuthenticationHelper.AcquireSharePointTokenAsync();
            return Ok(retValue);
        }
    }
}