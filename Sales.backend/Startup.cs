using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(Sales.backend.Startup))]
namespace Sales.backend
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
