using System;
using System.Configuration;
using System.Web.Mvc;
using SharePoint2010Migration.Web.Models;
using SharePoint2010Migration.Web.Services;

namespace SharePoint2010Migration.Web.Controllers
{
    public class MigrateController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {
            var model = new MigrationRunRequest
            {
                SiteUrl = ConfigurationManager.AppSettings["SharePointSiteUrl"],
                LibraryTitle = ConfigurationManager.AppSettings["SharePointLibraryTitle"],
                Domain = ConfigurationManager.AppSettings["SharePointDomain"],
                UserName = ConfigurationManager.AppSettings["SharePointUser"],
                ModifiedSinceUtc = DateTime.UtcNow.AddDays(-30)
            };

            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Run(MigrationRunRequest request)
        {
            if (!ModelState.IsValid)
            {
                return View("Index", request);
            }

            var cfg = MigrationConfiguration.LoadFromConfig();
            var repository = new DocumentRepository(cfg.SqlConnectionString);
            var migrator = new SharePointDocumentMigrator(repository);

            var result = migrator.Run(request);
            ViewBag.Result = result;

            return View("Index", request);
        }
    }
}
