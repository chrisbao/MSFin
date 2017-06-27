﻿/*   
 *   * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.  
 *   * See LICENSE in the project root for license information.  
 */

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SmartLink.Entity
{
    public enum SourcePointStatus
    {
        Created = 0,
        Deleted = 1
    }
    public class SourcePoint : BaseEntity
    {
        [StringLength(255)]
        public string Name { get; set; }
        [StringLength(255)]
        public string RangeId { get; set; }

        public string Position { get; set; }
        public string Value { get; set; }
        [StringLength(255)]
        public string Creator { get; set; }
        public DateTime Created { get; set; }
        public SourcePointStatus Status { get; set; }

        public Guid CatalogId { get; set; }
        [ForeignKey("CatalogId")]
        public virtual SourceCatalog Catalog { get; set; }
        public virtual ICollection<SourcePointGroup> Groups { get; set; }
        public virtual ICollection<PublishedHistory> PublishedHistories { get; set; }
        [JsonIgnore]
        public virtual ICollection<DestinationPoint> DestinationPoints { get; set; }

        public SourcePoint()
        {
            DestinationPoints = new List<DestinationPoint>();
            PublishedHistories = new List<PublishedHistory>();
            Groups = new List<SourcePointGroup>();
        }

        [NotMapped]
        [JsonIgnore]
        public bool SerializeCatalog { get; set; } = false;
        public bool ShouldSerializeCatalog()
        {
            return SerializeCatalog;
        }
    }
}
