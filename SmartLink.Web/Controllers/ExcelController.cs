using System.Web.Mvc;

namespace SmartLink.Web.Controllers
{
    [Authorize]
    public class ExcelController : Controller
    {
        // GET: Excel
        public ActionResult Point()
        {
            return View();
        }
    }
}