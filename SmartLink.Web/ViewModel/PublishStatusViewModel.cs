namespace SmartLink.Web.ViewModel
{
    public class PublishStatusViewModel
    {
        public string Status { get; set; }
        public PublishItemViewModel[] SourcePoints { get; set; }
    }

    public class PublishItemViewModel
    {
        public string Id { get; set; }
        public string Status { get; set; }
        public string Message { get; set; }
    }
}