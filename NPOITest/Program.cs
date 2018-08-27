using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOITest
{
    //
    //  ROWS     TIME      SETTINGS
    //  1M       5min      buffer=100000  compressed=true
    //  2M                 buffer=100000  compressed=true
    //
    //

    class Program
    {
        static void Main(string[] args)
        {
            var watch = new Stopwatch();
            watch.Start();

            using (var outStream = new FileStream("test.xlsx", FileMode.Create))
            {
                var wb = new XSSFWorkbook();
                var swb = new NPOI.XSSF.Streaming.SXSSFWorkbook(wb, 100000, true);
                var sh = swb.CreateSheet();
                for (int rownum = 0; rownum < 1048576; rownum++)
                {
                    var row = sh.CreateRow(rownum);
                    for (int cellnum = 0; cellnum < 10; cellnum++)
                    {
                        var cell = row.CreateCell(cellnum);
                        cell.SetCellValue(Guid.NewGuid().ToString());
                    }
                }

                //wb.Close();
                swb.Write(outStream);
            }

            watch.Stop();
            Console.WriteLine("Completed in {0} seconds", watch.Elapsed.TotalSeconds);
        }
    }
}
