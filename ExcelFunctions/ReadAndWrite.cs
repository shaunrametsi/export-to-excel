using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace Simple.ToExcel.ExcelFunctions
{
    public class ReadAndWrite<T>
    {
        public string[] Headings { private get; set; }
        public IEnumerable<T> Content { private get; set; }
        public string Path { private get; set;}
        public async Task<string> Export()
        {
            
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add(Path);

                //example for adding static headings from a list of arays
                //worksheet.Cells["A1"].LoadFromArrays(new List<string[]>() { Headings });

                worksheet.Cells["A1"].LoadFromCollection(Content, true,TableStyles.Dark10);

                package.SaveAs(new FileInfo(Path));
            }

            return $"Data has been written to: {Path}";
        }
        public async Task ReadFromExcel(string path)
        {
            using (var package = new ExcelPackage(new FileInfo(path)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var table = "";

                /* Print[Heading] */

                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    var cell = worksheet.GetValue<string>(1, col);
                    table += cell + "\t";
                }
                table += "\n-----------------------------------------------------------\n";
                /* Print[seperator] */

                //table += "-----------------------------------------------";
                /* Print[Table] */

                for (int row = 2; row <= worksheet.Dimension.Columns; row++)
                {
                    for (int column = 1; column <= worksheet.Dimension.Columns; column++)
                    {
                        var cell = worksheet.GetValue<string>(row, column);
                        table += cell + "\t";
                    }
                    table += "\n";
                }

                Console.Write(table);
                
            }
        }

    }
}
