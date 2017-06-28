/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

namespace SmartLink.Web.ViewModel
{
    public class DestinationPointForm
    {
        public string CatalogName { get; set; }

        public string RangeId { get; set; }

        public string SourcePointId { get; set; }

        public int[] CustomFormatIds { get; set; }
    }
}