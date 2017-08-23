/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System.Web.Mvc;

namespace SmartLink.Web.Controllers
{
    [Authorize]
    public class PowerPointController : Controller
    {
        // GET: PowerPoint
        public ActionResult Point()
        {
            return View();
        }
    }
}