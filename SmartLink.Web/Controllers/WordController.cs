using System.Web.Mvc;

namespace SmartLink.Web.Controllers
{
    [Authorize]
    public class WordController : Controller
    {
        // GET: Word
        public ActionResult Point()
        {
            return View();
        }
    }
}