using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelPro
{
    class Program
    {
        static void Main(string[] args)
        {

            //
            ExcelReader excelReader = new ExcelReader(@"D:\Study\ExcelPro\Book1.xlsx", sheetName :"Sheet1", ncolumnStartIndex:5, timeInterval: 1);
            //excelReader.Timer_Elapsed(null, null);

            excelReader.StartProcessing();


            Console.ReadLine();
        }
    }
}
