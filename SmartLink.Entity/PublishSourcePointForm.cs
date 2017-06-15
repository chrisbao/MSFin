using System;

namespace SmartLink.Entity
{
    public class PublishSourcePointForm
    {
        public Guid SourcePointId { get; set; }
        public string CurrentValue { get; set; }
        public string Position { get; set; }
    }
}