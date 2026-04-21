using System.Web;
using System.Web.Optimization;

namespace Romaneio
{
    public class BundleConfig
    {
         public static void RegisterBundles(BundleCollection bundles)
         {
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                        "~/Content/js/jquery-{version}.js"));

            bundles.Add(new ScriptBundle("~/bundles/jqueryval").Include(
                        "~/Content/plugins/jquery-validate/jquery.validate*")); 
            bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
                        "~/Scripts/modernizr-*"));

            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                      "~/Content/css/bootstrap.min.js",
                      "~/Content/js/bootstrap.js"));

            bundles.Add(new StyleBundle("~/Content/css").Include(
                      "~/Content/css/bootstrap.css",
                      "~/Content/css/site.css"));
             
         }
    }
}
