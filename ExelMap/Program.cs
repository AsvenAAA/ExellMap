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

         var exelbook = new XLWorkbook("C:/Users/alkin/Desktop/TestExell.xlsx");
         var exelsheet = exelbook.Worksheet(1);
         //exelsheet.Cell("B3").Value = "Hello World!";
         var rows = exelsheet.RangeUsed().RowsUsed();

         //string rowsNumb;
         var encod = Encoding.GetEncoding(1251);
         //StreamReader reader = new StreamReader(@"C:/Users/Asven/Desktop/readWrite.txt", encod);
         //string st;
         //while ((st = reader.ReadLine()) != "Evgen")
         //{
         //    Console.WriteLine(st);
         //}
         //reader.Close();

         //using (StreamReader reader = new StreamReader(@"C:/Users/Asven/Desktop/readWrite.txt", encod))
         //{
         //    string st;
         //    while((st = reader.ReadLine()) != "Evgen")
         //        Console.WriteLine(st);
         //}

         //Читает строки из файла и заносит их в список
         List<string> stIntxt = new List<string>();
         StreamReader reader = new StreamReader(@"C:/Users/alkin/Desktop/TestTxt.txt", encod);
         string st = "";
         while (st != "end")
         {
             if (st != " ")
             {
                 st = reader.ReadLine();
                 stIntxt.Add(st);
             }

         }

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
             stObj.Add($"{row.Cell(1).Value}");
         }


         foreach (string sT in stObj)
         {
             //надо брать поочередно строки из списка, 
             //и сравнивать со всеми строками из потока, пока не найдешь нужное, 
             //возможно строки из потока нужно будет пометсить куда либо, 
             //чтобы можно было сравнивать несколько раз
             //foreach(string sTxt in stIntxt)
             //{
             //    if (sT == sTxt)
             //    {
             //        var rangeMain = exelbook.Range("B2:B6");
             //        exelsheet.Cell("C1").Value = 
             //    }

             //}

             for (int i = 0; i < 16; i++)
             {
                 string sTxt = stIntxt[i];
                 bool boo = sT == sTxt;
                 if (boo)
                 {
                     exelsheet.Cell($"C{i + 1}").Value = stIntxt[i + 1];
                 }

             }
         }

         reader.Close();

         exelbook.Save();

         Console.ReadKey();

     }
 }

 class ExellData
 {
     public string Cell { get; set; }
 }

}
