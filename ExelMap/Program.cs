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
         //Открываем xlsx файл, выбираем первый лист, задаем кодировку Windows-1251
         var exellbook = new XLWorkbook("C:/Users/Asven/Desktop/TestExell.xlsx");
         var exellsheet = exellbook.Worksheet(1);
         var encod = Encoding.GetEncoding(1251);

         //Читает строки из файла и заносит их в список, игнорирует пустые строки, end нужен в качестве точки окончания
         List<string> stIntxt = new List<string>();
         StreamReader reader = new StreamReader(@"C:/Users/Asven/Desktop/TestTxt.txt", encod);
         string txtString = null;
         while (txtString != "end")
         {
             txtString = reader.ReadLine();
             if (txtString != "" && txtString != "end")
             {
                 stIntxt.Add(txtString);
             }
         }
         reader.Close();

         //список строк, которые считываются из второго столбца exell
         var rows = exellsheet.RangeUsed().RowsUsed();//вроде все не null строки
         List<string> stObj = new List<string>();
         foreach (var row in rows)
         {
             stObj.Add($"{row.Cell(2).Value}");
         }

         //Сравнивает строки из txt со страками из xlsx
         int count = 2;
         foreach (string stringInExell in stObj)
         {
             for (int i = 0; i < stIntxt.Count; i++)
             {
                 string stringTxt = stIntxt[i];
                 if (stringTxt == stringInExell)
                 {
                     string readWriteCount = stIntxt[i + 1].Remove(0, 9);
                     int readWriteCountInt = int.Parse(readWriteCount);
                     switch (readWriteCountInt)
                     {
                         case 0:
                             exellsheet.Cell($"C{count}").Value = "readonly";
                             break;
                         case 1:
                             exellsheet.Cell($"C{count}").Value = "writable";
                             break;
                         default:
                             exellsheet.Cell($"C{count}").Value = "Invalid variable value";
                             break;
                     }
                     count++;
                     break;
                 }

             }
         }
        var exellbook2 = new XLWorkbook("C:/Users/Asven/Desktop/TestExell2.xlsx");
        var exellsheet2 = exellbook2.Worksheet(1);
        var rows2 = exellsheet2.RangeUsed().RowsUsed();
        List<string> stObj2 = new List<string>();
        foreach (var row2 in rows2)
        {
            stObj2.Add($"{row2.Cell(2).Value}");
        }
        List<string> Obj = new List<string>();
        foreach (var row2 in rows2)
        {
            Obj.Add($"{row2.Cell(4).Value}");
        }
        count = 1;
        foreach (string sT in stObj)
        {
            for(int i = 0; i < stObj2.Count; i++)
             {
                    if (sT == stObj2[i])
                    {
                        exellsheet.Cell($"D{count}").Value = Obj[i];
                        count++;
                        break;
                    }
            }
        }
        
        exellbook.Save();
        Console.ReadKey();

     }
 }
}
