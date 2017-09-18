/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using ContosoO365DocSync.Entity;
using System.Threading.Tasks;

namespace ContosoO365DocSync.Service
{
    public interface ILogService
    {
        Task WriteLogAsync(LogEntity entity);

        void Flush();
    }
}