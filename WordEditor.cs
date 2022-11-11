using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;

namespace ParS
{
    internal class WordEditor
    {
        public string insNumber;
        public string insDateFrom;
        public string insDateTo;
        public object date;
        object replace = 2;

        Application app = new Application();
        Document doc = new Document();


        public void SaveAndClose(string fileDir, string personName)
        {
            string createNew = $"{fileDir}" + $"{personName}.docx";
            doc.SaveAs2(createNew);
            doc.Close(true);


            Marshal.ReleaseComObject(doc);

        }

        public void EditFileOne(string fileDir, string personName)
        {
            object fileName = fileDir;
            object falseValue = false;
            object missing = Type.Missing;

            //Открытие
            doc = app.Documents.Open(ref fileName, ref missing, ref falseValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);

            string fieldDate = "{DATE}";

            app.Selection.Find.Execute(fieldDate, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing, ref date,
            ref replace, ref missing, ref missing, ref missing, ref missing);

            SaveAndClose($"{fileDir.Substring(0, fileDir.Length - 5)}", "Выписка СРО" + $" на {date}");
        }

        public void EditFileTwo(string fileDir, string personName)
        {
            object fileName = fileDir;
            object falseValue = false;
            object missing = Type.Missing;

            //Открытие
            doc = app.Documents.Open(ref fileName, ref missing, ref falseValue,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing, ref missing, ref missing,
            ref missing, ref missing, ref missing);


            //Поля для замены текста 
            object[] replaceFields = new object[4];
            replaceFields[0] = "{INS_NUMBER}";
            replaceFields[1] = "{INS_DATE_FROM}";
            replaceFields[2] = "{INS_DATE_TO}";
            replaceFields[3] = "{DATE}";

            object[] replaceWith = new object[4];
            replaceWith[0] = insNumber;
            replaceWith[1] = insDateFrom;
            replaceWith[2] = insDateTo;
            replaceWith[3] = date;

            //Замена текста через цикл
            for (int i = 0; i < replaceFields.Length; i++)
            {
                app.Selection.Find.Execute(replaceFields[i], ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceWith[i],
                ref replace, ref missing, ref missing, ref missing, ref missing);
            }
            SaveAndClose($"{fileDir.Substring(0, fileDir.Length - 5)}", " ИНФ.Выписка СРО" + $" на {date}");

        }
    }
}
