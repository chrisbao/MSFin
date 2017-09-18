﻿/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

namespace ContosoO365DocSync.Entity
{
    public enum ActionTypeEnum
    {
        AuditLog = 0,
        ErrorLog = 1
    }

    public class LogEntity
    {
        public string LogId { get; set; }

        public string Action { get; set; }

        public string PointType { get; set; }

        public ActionTypeEnum ActionType { get; set; }

        public string Subject { get; set; }

        public string Message { get; set; }

        public string Detail { get; set; }
    }
}