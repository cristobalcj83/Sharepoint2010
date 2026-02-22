using System;
using System.IO;
using System.Web.Script.Serialization;
using SharePoint2010Migration.Web.Models;
using SharePoint2010Migration.Web.Services;

namespace SharePoint2010Migration.Runner
{
    internal class Program
    {
        private static int Main(string[] args)
        {
            try
            {
                var settingsPath = args.Length > 0 ? args[0] : @"..\SharePoint2010Migration.Web\App_Data\admin-settings.json";
                if (!File.Exists(settingsPath))
                {
                    Console.WriteLine("Settings file not found: " + settingsPath);
                    Console.WriteLine("Create settings by running the web app and saving /Admin/Index first.");
                    return 2;
                }

                var serializer = new JavaScriptSerializer();
                var settings = serializer.Deserialize<AdminSettingsModel>(File.ReadAllText(settingsPath));

                var request = new MigrationRunRequest
                {
                    SiteUrl = settings.SiteUrl,
                    LibraryTitle = settings.LibraryTitle,
                    Domain = settings.Domain,
                    UserName = settings.UserName,
                    Password = settings.Password,
                    ModifiedSinceUtc = null
                };

                var cfg = MigrationConfiguration.LoadFromConfig();
                var repository = new DocumentRepository(cfg.SqlConnectionString);
                var migrator = new SharePointDocumentMigrator(repository);
                var result = migrator.Run(request);

                Console.WriteLine("Processed: " + result.ProcessedCount);
                Console.WriteLine("Inserted/Updated: " + result.InsertedOrUpdatedCount);
                Console.WriteLine("Skipped non-PDF: " + result.SkippedNonPdfCount);
                Console.WriteLine("Errors: " + result.ErrorCount);
                Console.WriteLine("Last error: " + result.LastError);

                return result.ErrorCount > 0 ? 1 : 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Fatal error: " + ex);
                return 1;
            }
        }
    }
}
