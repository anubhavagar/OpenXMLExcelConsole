using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXMLExcelConsole
{
    class Program
    {
        static void Main(string[] args)
        {

            Dictionary<string, string> ReplacemenDict = new Dictionary<string, string>();

            string[] SeriesLabels = { "Series One", "Series Two", "Series Three", "Series Four", "Series Five", "Series Six", "Series Seven", "Series Eight", "Series Nine", "Series Ten" };

            ReplacemenDict.Add("[ReportName]", "Cloud Inc Report");
            ReplacemenDict.Add("[CreatedBy]", "Cloud Team");
            ReplacemenDict.Add("[Company]", "Pitney Bowes");
            ReplacemenDict.Add("[TableName]", "Data new table");
            ReplacemenDict.Add("[ChartDataTableName]", "New Table");
            ReplacemenDict.Add("[ChartName]", "New Chart");
            ReplacemenDict.Add("[BriefDescription]", "Hello above is the custom server side chart");


            string templatefilepath = "C:\\Anubhav\\projects\\OpenXMLExcelConsole\\OpenXMLExcelConsole\\Template\\";
            string resultfilepath = "C:\\Anubhav\\projects\\OpenXMLExcelConsole\\OpenXMLExcelConsole\\Output\\";


            ExcelClassLibrary.ExcelClass obj1 = new ExcelClassLibrary.ExcelClass(templatefilepath, resultfilepath);

            Console.WriteLine(obj1.InitBookCreation("ExcelTemplate.xlsx", "CloudReport.xlsx", "TemplateSheetFile"));

            Console.WriteLine(obj1.AddSheetWithTable("varundatasheet", Program.GetDemoChartData(), ReplacemenDict));

            Console.WriteLine(obj1.AddSheetWithChart("mayankchartsheet", Program.GetDemoChartData(), SeriesLabels, ReplacemenDict));

           Console.WriteLine(obj1.EndBookCreation("CloudReport.xlsx"));

            Console.WriteLine("Press <Enter> to Exit");

            Console.ReadLine();

            //ExcelClassLibrary.testchart obj1 = new ExcelClassLibrary.testchart(templatefilepath, resultfilepath);
            //Console.WriteLine(obj1.AddSheetWithChart("mayankchartsheet", Program.GetDemoChartData(), SeriesLabels, ReplacemenDict));
            //Console.ReadLine();
        }

        // generate demo data for chart and table data
        public static List<List<object>> GetDemoChartData()
        {
            List<List<object>> list = new List<List<object>>();

            for (int i = 0; i < 10; i++)
            {

                List<object> sublist = new List<object>();

                sublist.Add(5);
                sublist.Add(10);
                sublist.Add(20);
                sublist.Add(25);

                list.Add(sublist);


            }
            return list;
        }
    }
}
