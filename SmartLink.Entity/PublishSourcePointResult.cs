using System;
using System.Collections.Generic;

namespace SmartLink.Entity
{
    public class PublishSourcePointResult
    {
        public Guid BatchId { get; set; }
        public ICollection<SourcePoint> SourcePoints { get; set; }
    }
}
