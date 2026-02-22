using System.Web.Mvc;
using System.Web.Routing;
using SharePoint2010Migration.Web.App_Start;

namespace SharePoint2010Migration.Web
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            RouteConfig.RegisterRoutes(RouteTable.Routes);
        }
    }
}
