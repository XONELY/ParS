using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using HtmlAgilityPack;
using Microsoft.Office.Interop.Word;
using static System.Net.Mime.MediaTypeNames;

namespace ParS
{
    internal class Program
    {
        public static string directory = @"D:\Desktop\DEKART\Выписки\";

        static void Main()
        {
            string typeOne = "";
            string typeTwo = ".ИНФ";

            for (int i = 0; i < Raters.rater.Length; i++)
            {
                GetHtml.Start(i);
                WordEditor.EditFile($"{directory}{Raters.rater[i].name}.docx", i, directory, typeOne);
                WordEditor.EditFile($"{directory}{Raters.rater[i].name}.docx",i,directory,typeTwo);
                
            }
        }
    }
}


