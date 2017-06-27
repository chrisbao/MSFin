/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

namespace SmartLink.Web.ViewModel
{
    public class PublishStatusViewModel
    {
        public string Status { get; set; }
        public PublishItemViewModel[] SourcePoints { get; set; }
    }

    public class PublishItemViewModel
    {
        public string Id { get; set; }
        public string Status { get; set; }
        public string Message { get; set; }
    }
}