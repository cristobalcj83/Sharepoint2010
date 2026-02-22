using System;
using System.IO;
using System.Web;
using System.Web.Script.Serialization;
using SharePoint2010Migration.Web.Models;

namespace SharePoint2010Migration.Web.Services
{
    public class AdminSettingsService
    {
        private const string SettingsFileName = "admin-settings.json";

        public AdminSettingsModel Load()
        {
            var path = GetSettingsPath();
            if (!File.Exists(path))
            {
                return GetDefault();
            }

            var serializer = new JavaScriptSerializer();
            var content = File.ReadAllText(path);
            var model = serializer.Deserialize<AdminSettingsModel>(content) ?? GetDefault();
            EnrichPathChecks(model);
            return model;
        }

        public void Save(AdminSettingsModel model)
        {
            var path = GetSettingsPath();
            var serializer = new JavaScriptSerializer();
            var content = serializer.Serialize(model);

            var folder = Path.GetDirectoryName(path);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }

            File.WriteAllText(path, content);
        }

        private static string GetSettingsPath()
        {
            var basePath = HttpContext.Current.Server.MapPath("~/App_Data");
            return Path.Combine(basePath, SettingsFileName);
        }

        private static AdminSettingsModel GetDefault()
        {
            var model = new AdminSettingsModel
            {
                SharePointClientDllPath = @"C:\\Program Files\\Common Files\\Microsoft Shared\\Web Server Extensions\\14\\ISAPI\\Microsoft.SharePoint.Client.dll",
                SharePointRuntimeDllPath = @"C:\\Program Files\\Common Files\\Microsoft Shared\\Web Server Extensions\\14\\ISAPI\\Microsoft.SharePoint.Client.Runtime.dll",
                SiteUrl = "http://sp2010vm/sites/migration",
                LibraryTitle = "Shared Documents",
                Domain = "CONTOSO",
                UserName = "spreader",
                Password = string.Empty
            };

            EnrichPathChecks(model);
            return model;
        }

        public static void EnrichPathChecks(AdminSettingsModel model)
        {
            model.IsClientDllFound = !string.IsNullOrWhiteSpace(model.SharePointClientDllPath) && File.Exists(model.SharePointClientDllPath);
            model.IsRuntimeDllFound = !string.IsNullOrWhiteSpace(model.SharePointRuntimeDllPath) && File.Exists(model.SharePointRuntimeDllPath);
        }
    }
}
