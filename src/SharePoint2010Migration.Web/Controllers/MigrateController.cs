using System;
using System.Web.Mvc;
using SharePoint2010Migration.Web.Models;
using SharePoint2010Migration.Web.Services;

namespace SharePoint2010Migration.Web.Controllers
{
    public class MigrateController : Controller
    {
        private readonly AdminSettingsService _settingsService;

        public MigrateController()
        {
            _settingsService = new AdminSettingsService();
        }

        [HttpGet]
        public ActionResult Index()
        {
            var settings = _settingsService.Load();
            var model = new MigrationRunRequest
            {
                SiteUrl = settings.SiteUrl,
                LibraryTitle = settings.LibraryTitle,
                Domain = settings.Domain,
                UserName = settings.UserName,
                Password = settings.Password,
                ModifiedSinceUtc = DateTime.UtcNow.AddDays(-30)
            };

            ViewBag.AdminSettings = settings;
            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Run(MigrationRunRequest request)
        {
            var settings = _settingsService.Load();
            PopulateMissingFromAdminSettings(request, settings);

            if (!ModelState.IsValid)
            {
                ViewBag.AdminSettings = settings;
                return View("Index", request);
            }

            var cfg = MigrationConfiguration.LoadFromConfig();
            var repository = new DocumentRepository(cfg.SqlConnectionString);
            var migrator = new SharePointDocumentMigrator(repository);

            var result = migrator.Run(request);
            ViewBag.Result = result;
            ViewBag.AdminSettings = settings;

            return View("Index", request);
        }

        private static void PopulateMissingFromAdminSettings(MigrationRunRequest request, AdminSettingsModel settings)
        {
            if (string.IsNullOrWhiteSpace(request.SiteUrl))
            {
                request.SiteUrl = settings.SiteUrl;
            }

            if (string.IsNullOrWhiteSpace(request.LibraryTitle))
            {
                request.LibraryTitle = settings.LibraryTitle;
            }

            if (string.IsNullOrWhiteSpace(request.Domain))
            {
                request.Domain = settings.Domain;
            }

            if (string.IsNullOrWhiteSpace(request.UserName))
            {
                request.UserName = settings.UserName;
            }

            if (string.IsNullOrWhiteSpace(request.Password))
            {
                request.Password = settings.Password;
            }
        }
    }
}
