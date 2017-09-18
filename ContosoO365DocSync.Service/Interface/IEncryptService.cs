/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

namespace ContosoO365DocSync.Service
{
    public interface IEncryptService
    {
        string EncryptString(string planText);
        string DecryptString(string cipherText);
    }
}
