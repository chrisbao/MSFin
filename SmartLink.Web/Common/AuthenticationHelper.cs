using System.Threading.Tasks;

namespace SmartLink.Web.Common
{
    public class AuthenticationHelper
    {
        public static string token;

        public static string consentUrl;

        public static async Task<string> AcquireTokenAsync()
        {
             return token;
        }
    }
}