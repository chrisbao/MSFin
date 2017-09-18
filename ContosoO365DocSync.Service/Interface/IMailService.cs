/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System.Collections.Generic;
using System.Threading.Tasks;

namespace ContosoO365DocSync.Service
{
    public interface IMailService
    {
        Task SendPlainTextMailAsync(string fromAddress, string fromDisplayName, IEnumerable<string> toAddresses, string subject, string content);
    }
}