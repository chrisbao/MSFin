/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System.Web.Mvc;

namespace ContosoO365DocSync.Web.Controllers
{
    [Authorize]
    public class ExcelController : Controller
    {
        // GET: Excel
        public ActionResult Point()
        {
            return View();
        }
    }
}