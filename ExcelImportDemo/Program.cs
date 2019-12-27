using ExcelImport;

namespace ExcelImportDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            var helper = new ExcelImportHelper<TestData>();

            //导出模板
            helper.ExportTemplate("c:\\test.xls");

            //导入数据
            var datas = helper.Import("c:\\test.xls");
        }
    }

    public class TestData : ExcelData
    {
        [ExcelHeader("名称", Comment = "Name")]
        public string Name { get; set; }

        [ExcelHeader("年龄", Comment = "Age")]
        public string Age { get; set; }
    }

}
