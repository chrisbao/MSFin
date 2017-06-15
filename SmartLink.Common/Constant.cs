namespace SmartLink.Common
{
    static public class Constant
    {
        public const int AZURETABLE_BATCH_COUNT = 100;
        public const string PUBLISH_QUEUE_NAME = "publishqueue";
        public const string PUBLISH_TABLE_NAME = "publishtable";

        static public readonly string POINTTYPE_SOURCEPOINT = "Source Point";
        static public readonly string POINTTYPE_SOURCEPOINTHISTORY = "Source Point history";
        static public readonly string POINTTYPE_DESTINATIONPOINT = "Destination Point";
        static public readonly string POINTYTPE_DESTINATIONCATALOG = "Destination Catalog";
        static public readonly string POINTTYPE_DESTINATIONLIST = "Destination Points list";
        static public readonly string POINTTYPE_SOURCECATALOG = "Source Catalog";
        static public readonly string POINTTYPE_SOURCECATALOGLIST = "Source Catalog list";

        static public readonly string ACTIONTYPE_GET = "Get";
        static public readonly string ACTIONTYPE_ADD = "Add";
        static public readonly string ACTIONTYPE_EDIT = "Edit";
        static public readonly string ACTIONTYPE_DELETE = "Delete";
        static public readonly string ACTIONTYPE_PUBLISH = "Publish";
    }
}
