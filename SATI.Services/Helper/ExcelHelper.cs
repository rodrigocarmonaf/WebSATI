using System.Configuration;
namespace SATI.Services.Helper
{
   public static class ExcelHelper
   {
        public static string StringConnectionExcel()
        {
            return ConfigurationManager.AppSettings["excel:conn"].ToString();
        }
        public static string PathUploadExcel()
        {
            return ConfigurationManager.AppSettings["folder:excel"].ToString();
        }

        public static string RangeColumsExcel()
        {
            return ConfigurationManager.AppSettings["range:exel"].ToString();
        }
    }
}
