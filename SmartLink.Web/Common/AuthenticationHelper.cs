/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

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