/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

namespace ContosoO365DocSync.Service
{
    public class EncryptionService : IEncryptService
    {
        private AES _decrypt;
        public EncryptionService()
        {
            _decrypt = new AES();
        }

        /// <summary>
        /// Decrypt string.
        /// </summary>
        /// <param name="cipherText"></param>
        /// <returns></returns>
        public string DecryptString(string cipherText)
        {
            return _decrypt.Decrypt(cipherText);
        }

        /// <summary>
        /// Encrypt string.
        /// </summary>
        /// <param name="plainText"></param>
        /// <returns></returns>
        public string EncryptString(string plainText)
        {
            return _decrypt.Encrypt(plainText);
        }
    }
}
