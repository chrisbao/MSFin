using System;

namespace SmartLink.Entity
{
    public class PublishedMessage
    {
        public Guid PublishBatchId { get; set; }
        public Guid PublishHistoryId { get; set; }
        public Guid SourcePointId { get; set; }
    }
}
