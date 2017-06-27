/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using SmartLink.Entity;

namespace SmartLink.Service
{
    public interface IUserProfileService
    {
        UserProfile GetCurrentUser();
    }
}
