/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Newtonsoft.Json;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace ContosoO365DocSync.Entity
{
    public class DestinationCatalog : BaseEntity
    {
        [NotMapped]
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

        public string DocumentId { get; set; }

        public ICollection<DestinationPoint> DestinationPoints { get; set; }

        public DestinationCatalog()
        {
            DestinationPoints = new List<DestinationPoint>();
        }

        [NotMapped]
        [JsonIgnore]
        public bool SerializeDestinationPoints { get; set; } = true;

        public bool ShouldSerializeDestinationPoints()
        {
            return SerializeDestinationPoints;
        }
    }
}