using System.Web.Mvc;
using SharePoint2010Migration.Web.Models;
using SharePoint2010Migration.Web.Services;

namespace SharePoint2010Migration.Web.Controllers
{
    public class AdminController : Controller
    {
        private readonly AdminSettingsService _settingsService;

        public AdminController()
        {
            _settingsService = new AdminSettingsService();
        }

        [HttpGet]
        public ActionResult Index()
        {
            var model = _settingsService.Load();
            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Save(AdminSettingsModel model)
        {
            if (!ModelState.IsValid)
            {
                AdminSettingsService.EnrichPathChecks(model);
                return View("Index", model);
            }

            _settingsService.Save(model);
            AdminSettingsService.EnrichPathChecks(model);
            ViewBag.Message = "Settings saved successfully.";

            return View("Index", model);
        }
    }
}
