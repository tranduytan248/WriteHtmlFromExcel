using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(WriteHtmlFromExcel.Startup))]
namespace WriteHtmlFromExcel
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
