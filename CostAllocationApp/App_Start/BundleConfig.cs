using System.Web;
using System.Web.Optimization;

namespace CostAllocationApp
{
    public class BundleConfig
    {
        // For more information on bundling, visit https://go.microsoft.com/fwlink/?LinkId=301862
        public static void RegisterBundles(BundleCollection bundles)
        {
            bundles.Add(new ScriptBundle("~/bundles/jquery").Include(
                        "~/Scripts/jquery-3.5.1.min.js"));


            bundles.Add(new ScriptBundle("~/bundles/bootstrap").Include(
                            "~/Scripts/popper.min.js",
                            "~/Scripts/bootstrap.min.js"));


            bundles.Add(new ScriptBundle("~/bundles/others").Include(
                            "~/Scripts/jquery.slimscroll.min.js",
                            "~/Scripts/select2.min.js",
                            "~/Scripts/moment.min.js",
                            "~/Scripts/bootstrap-datetimepicker.min.js",
                            "~/Scripts/jquery.dataTables.min.js",
                            "~/Scripts/dataTables.bootstrap4.min.js",
                            "~/Scripts/toastr.min.js"
 
                            //,
                            //"~/Scripts/loader.js",
                            //"~/Scripts/Chart.min.js"
                            ));

            bundles.Add(new ScriptBundle("~/bundles/custom").Include(
                            "~/Scripts/app.js"));


            bundles.Add(new StyleBundle("~/Content/css").Include(
                    "~/Content/bootstrap.min.css",
                    "~/Content/font-awesome.min.css",
                    "~/Content/line-awesome.min.css",
                    "~/Content/dataTables.bootstrap4.min.css",
                    "~/Content/select2.min.css",
                    "~/Content/bootstrap-datetimepicker.min.css",
                    "~/Content/toastr.min.css",
                    "~/Content/style_ver02.css"
                    //,
                    //"~/Content/custom.css"
                    ));

        }
    }
}
