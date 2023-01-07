using ExcelDna.Integration;

namespace yExcelFunctions
{
    public class Functions
    {
        [ExcelFunction(Description = "For Voyager Multi-Select: Combines values together with a carat ^")]
        public static string yCarat(object[] list)
        {
            return String.Join("^", list);
        }

        [ExcelFunction(Description = "For SQL: Combines values together in SQL formatted Strings")]
        public static string yIn(object[] list)
        {
            return $"'{String.Join("','", list)}'";
        }
    }
}