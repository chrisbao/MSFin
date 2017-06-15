using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace SmartLink.Entity
{
    public class DestinationPoint : BaseEntity
    {
        [ForeignKey("SourcePointId")]
        public SourcePoint ReferencedSourcePoint { get; set; }
        public Guid SourcePointId { get; set; }
        [ForeignKey("CatalogId")]
        public DestinationCatalog Catalog { get; set; }
        [StringLength(255)]
        public string RangeId { get; set; }
        public Guid CatalogId { get; set; }
        [StringLength(255)]
        public string Creator { get; set; }
        public DateTime Created { get; set; }

        [NotMapped]
        [JsonIgnore]
        public bool SerializeCatalog { get; set; } = false;
        public bool ShouldSerializeCatalog()
        {
            return SerializeCatalog;
        }

        public virtual ICollection<CustomFormat> CustomFormats { get; set; }

        public DestinationPoint()
        {
            CustomFormats = new List<CustomFormat>();
        }
    }
}
