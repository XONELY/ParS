using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ParS
{
    internal class WordEditor
    {
        public static string insNumber;
        public static string insDateFrom;
        public static string insDateTo;
        public static object date;

        public static void EditFile(string fileDir, int counter, string saveDir, string type)
        {
            Application app = new Application();
            Document doc = new Document();
            object replace = 2;
            object fileName = $"{Program.directory}{Raters.rater[counter].name}{type}.docx";
            object falseValue = false;
            object missing = Type.Missing;


            void SaveAndClose()
            {
                string createNew = $"{saveDir}{Raters.rater[counter].name} {type}" + $" на {date}.docx";
                doc.SaveAs2(createNew);
                doc.Close(true);
                Marshal.ReleaseComObject(doc);

            }
            void Open()
            {
                doc = app.Documents.Open(ref fileName, ref missing, ref falseValue,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);

                if (type != ".ИНФ") //typeOne
                {
                    string fieldDate = "{DATE}";
                    app.Selection.Find.Execute(fieldDate, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref date,
                    ref replace, ref missing, ref missing, ref missing, ref missing);

                    SaveAndClose();
                }
                else  //typeTWO
                {
                    object[] replaceFields =
                        {
                    "{INS_NUMBER}",
                    "{INS_DATE_FROM}",
                    "{INS_DATE_TO}",
                    "{DATE}",
                };

                    object[] replaceWith = new object[4];
                    replaceWith[0] = insNumber;
                    replaceWith[1] = insDateFrom;
                    replaceWith[2] = insDateTo;
                    replaceWith[3] = date;

                    for (int i = 0; i < replaceFields.Length; i++)
                    {
                        app.Selection.Find.Execute(replaceFields[i], ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith[i],
                        ref replace, ref missing, ref missing, ref missing, ref missing);
                    }
                    SaveAndClose();
                }
            }
            Open();
        }
    }
}
