using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(CSVImport.Web.Startup))]
namespace CSVImport.Web
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
