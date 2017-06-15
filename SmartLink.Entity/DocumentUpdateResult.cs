using System.Collections.Generic;

namespace SmartLink.Entity
{
    public class DocumentUpdateResult
    {
        public bool IsSuccess { get; set; }
        public List<string> Message { get; set; }
        public DocumentUpdateResult()
        {
            Message = new List<string>();
        }
    }
}
