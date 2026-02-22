using System.Configuration;

namespace SharePoint2010Migration.Web.Services
{
    public class MigrationConfiguration
    {
        public string SqlConnectionString { get; private set; }

        public static MigrationConfiguration LoadFromConfig()
        {
            return new MigrationConfiguration
            {
                SqlConnectionString = ConfigurationManager.ConnectionStrings["MigrationDb"].ConnectionString
            };
        }
    }
}
