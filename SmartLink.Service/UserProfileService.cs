﻿/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using SmartLink.Entity;
using System.Linq;
using System.Security.Claims;
using System.Threading;

namespace SmartLink.Service
{
    public class UserProfileService : IUserProfileService
    {
        /// <summary>
        /// Get current user profile information.
        /// </summary>
        /// <returns></returns>
        public UserProfile GetCurrentUser()
        {
            var currentUserProfile = new UserProfile();
            var icp = Thread.CurrentPrincipal as ClaimsPrincipal;
            if (icp != null)
            {
                currentUserProfile.Email = icp.Identity.Name;
                var nameClaim = icp.Claims.FirstOrDefault(o => o.Type == "name");
                if (nameClaim != null)
                {
                    currentUserProfile.Username = nameClaim.Value;
                }
            }
            return currentUserProfile;
        }
    }
}