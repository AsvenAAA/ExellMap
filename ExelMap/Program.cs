using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using ClosedXML.Attributes;
using ClosedXML.Excel;
using System.IO;

namespace ExelMap
{
 class Program
 {
     static void Main(string[] args)
     {
         //ExellData eData = new ExellData();

         var exelbook = new XLWorkbook("C:/Users/Asven/Desktop/TestExell.xlsx");
         var exelsheet = exelbook.Worksheet(1);
         var rows = exelsheet.RangeUsed().RowsUsed();

         //string rowsNumb;
         var encod = Encoding.GetEncoding(1251);

         //Читает строки из файла и заносит их в список
         List<string> stIntxt = new List<string>();
         StreamReader reader = new StreamReader(@"C:/Users/Asven/Desktop/TestTxt.txt", encod);
         string st = null;
         while (st != "end")
         {
             st = reader.ReadLine();
             if (st != "" && st != "end")
             {
                 stIntxt.Add(st);
             }
         }
         reader.Close();

            //список объектов ExellData
            List<ExellData> eDataObj = new List<ExellData>();
         foreach (var row in rows)
         {
             eDataObj.Add(new ExellData() { Cell = $"{row.Cell(2).Value}" });
         }

         //список строк для приравнивания строки к строке
         List<string> stObj = new List<string>();
         foreach (var row in rows)
         {
             stObj.Add($"{row.Cell(2).Value}");
         }

         foreach (string sT in stObj)
         {
             for (int i = 0; i < 12; i++)
             {
                 string sTxt = stIntxt[i];
                 if (sTxt == sT)
                 {
                     exelsheet.Cell($"C{i + 1}").Value = stIntxt[i + 1];
                 }

             }
         }
         exelbook.Save();
         Console.ReadKey();

     }
 }
 class ExellData
 {
     public string Cell { get; set; }
 }

}
