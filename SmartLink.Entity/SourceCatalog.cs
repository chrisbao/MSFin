/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Newtonsoft.Json;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace SmartLink.Entity
{
    public class SourceCatalog : BaseEntity
    {
        [NotMapped]
        [JsonIgnore]
        public string FileName
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(Name))
                    return Name.Substring(Name.LastIndexOf('/') + 1);
                else
                    return Name;
            }
        }

        public string Name { get; set; }

        public ICollection<SourcePoint> SourcePoints { get; set; }

        public SourceCatalog()
        {
            SourcePoints = new List<SourcePoint>();
        }

        [NotMapped]
        [JsonIgnore]
        public bool SerializeSourcePoints { get; set; } = true;

        public bool ShouldSerializeSourcePoints()
        {
            return SerializeSourcePoints;
        }
    }
}