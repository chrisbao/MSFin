using System;

namespace SmartLink.Web.ViewModel
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
    }

}