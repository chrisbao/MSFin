/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System.Threading.Tasks;

namespace ContosoO365DocSync.Web.Common
{
    public class AuthenticationHelper
    {
        public static string token;

        public static string consentUrl;

        public static string sharePointToken;

        public static Task<string> AcquireTokenAsync()
        {
             return Task.FromResult(token);
        }

        public static Task<string> AcquireSharePointTokenAsync()
        {
            return Task.FromResult(sharePointToken);
        }
    }
}