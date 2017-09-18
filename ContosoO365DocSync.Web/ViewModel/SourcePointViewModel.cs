/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using System;

namespace ContosoO365DocSync.Web.ViewModel
{
    public class SourcePointForm
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string CatalogName { get; set; }

        public string RangeId { get; set; }

        public string Position { get; set; }

        public string Value { get; set; }

        public string Creator { get; set; }

        public DateTime Created { get; set; }

        public int[] GroupIds { get; set; }

        public string NameRangeId { get; set; }

        public string NamePosition { get; set; }

        public string DocumentId { get; set; }
    }
}