/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System;

namespace SmartLink.Service
{
    public static class DateTimeService
    {
        /// <summary>
        /// Convert UTC to PST timezone.
        /// </summary>
        /// <param name="utcDateTime"></param>
        /// <returns></returns>
        public static DateTime ToPSTDateTime(this DateTime utcDateTime)
        {
            return TimeZoneInfo.ConvertTimeFromUtc(utcDateTime, TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time"));
        }
    }
}
