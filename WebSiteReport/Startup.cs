using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(WebSiteReport.Startup))]
namespace WebSiteReport
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
