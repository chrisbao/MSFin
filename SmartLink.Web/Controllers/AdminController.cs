/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using SmartLink.Web.Common;
using SmartLink.Web.ViewModel;
using System.Web;
using System.Web.Mvc;

namespace SmartLink.Web.Controllers
{
    public class AdminController : Controller
    {
        [Authorize]
        public ActionResult Consent()
        {
            if (AuthenticationHelper.consentUrl == null)
            {
                string callbackUrl = Url.Action("Consent", "Admin", routeValues: null, protocol: Request.Url.Scheme);
                HttpContext.GetOwinContext().Authentication.SignOut(new AuthenticationProperties { RedirectUri = callbackUrl },
                    OpenIdConnectAuthenticationDefaults.AuthenticationType, CookieAuthenticationDefaults.AuthenticationType);
            }

            return View(new AdminConsentViewModel()
            {
                State = "0",
                Url = AuthenticationHelper.consentUrl
            });
        }

        public ActionResult Result()
        {
            return View("Consent", new AdminConsentViewModel()
            {
                State = Request.Url.ToString().ToUpper().IndexOf("CODE=") > -1 ? "1" : "2"
            });
        }
    }
}