using Microsoft.Azure;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using Owin;
using SmartLink.Web.Common;
using System;
using System.Threading.Tasks;
using System.Web;

namespace SmartLink.Web
{
    public partial class Startup
    {
        private static string clientId = CloudConfigurationManager.GetSetting("ida:ClientId");
        private static string aadInstance = CloudConfigurationManager.GetSetting("ida:AADInstance");
        private static string tenantId = CloudConfigurationManager.GetSetting("ida:TenantId");
        private static string postLogoutRedirectUri = CloudConfigurationManager.GetSetting("ida:PostLogoutRedirectUri");
        private static string appKey = CloudConfigurationManager.GetSetting("ida:ClientSecret");
        private static string resourceId = CloudConfigurationManager.GetSetting("ResourceId");
        private static string consentResource = CloudConfigurationManager.GetSetting("ConsentResource");
        private static string authority = aadInstance + tenantId;
        /// <summary>
        /// Implement the OPENID authentication and get the access token to access SP site.
        /// </summary>
        /// <param name="app"></param>
        public void ConfigureAuth(IAppBuilder app)
        {
            app.SetDefaultSignInAsAuthenticationType(CookieAuthenticationDefaults.AuthenticationType);

            app.UseCookieAuthentication(new CookieAuthenticationOptions());

            app.UseOpenIdConnectAuthentication(
                new OpenIdConnectAuthenticationOptions
                {
                    ClientId = clientId,
                    Authority = authority,
                    PostLogoutRedirectUri = postLogoutRedirectUri
                    ,
                    Notifications = new OpenIdConnectAuthenticationNotifications()
                    {
                        AuthorizationCodeReceived = (context) =>
                        {
                            var code = context.Code;
                            var redirectUrl = new Uri(HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Path));
                            var consentRedirectUrl = new Uri(HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Path) + "Admin/Result");
                            ClientCredential credential = new ClientCredential(clientId, appKey);
                            AuthenticationContext authContext = new AuthenticationContext(authority);

                            string userObjectID = context.AuthenticationTicket.Identity.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;
                            UserIdentifier userIdentifier = new UserIdentifier(userObjectID, UserIdentifierType.UniqueId);
                            AuthenticationHelper.consentUrl = authContext.GetAuthorizationRequestURL(consentResource, clientId, consentRedirectUrl, userIdentifier, "prompt=admin_consent").ToString();

                            AuthenticationResult result = authContext.AcquireTokenByAuthorizationCode(code, redirectUrl, credential, resourceId);
                            AuthenticationHelper.token = result.AccessToken;

                            return Task.FromResult(0);
                        }
                    }
                });


        }
    }
}