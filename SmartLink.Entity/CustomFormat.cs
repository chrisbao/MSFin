﻿/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Newtonsoft.Json;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SmartLink.Entity
{
    public class CustomFormat
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public int Id { get; set; }

        public string Name { get; set; }

        public string DisplayName { get; set; }

        public string Description { get; set; }

        public int OrderBy { get; set; }

        public bool IsDeleted { get; set; }

        public string GroupName { get; set; }

        public int GroupOrderBy { get; set; }

        [JsonIgnore]
        public virtual ICollection<DestinationPoint> DestinationPoints { get; set; }
    }
}