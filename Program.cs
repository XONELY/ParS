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
        static HtmlWeb web = new HtmlWeb();
        static WordEditor editor = new WordEditor();

        static void Main()
        {
            string dateToday = Convert.ToString(DateTime.Today);
            editor.date = dateToday.Substring(0, dateToday.Length - 8);
            string infoNode = "//table[@class='uk-table uk-table-small bordered uk-table-divider uk-table-striped']";
            string fileDirection = @"D:\Desktop\DEKART\Выписки\";

            string[] urls = new string[]
            {
                @"https://srosoyz.ru/sro/members/b5f29bc4-eba2-4360-857d-35c2605ea67d",
                @"https://srosoyz.ru/sro/members/1ea650cc-d6e5-45a6-8d58-cafdc21a3170",
                @"https://srosoyz.ru/sro/members/188d1a1a-46f0-487e-a007-e4b39778cba8",
                @"https://srosoyz.ru/sro/members/ad018a30-c0eb-41b3-9a58-2f065bdbf01a",
                @"https://srosoyz.ru/sro/members/df2734fd-38e1-40c9-bdc2-f5543501913f",
                @"https://srosoyz.ru/sro/members/20d5a756-e2b5-486e-b171-cdf98c184e42",
                @"https://srosoyz.ru/sro/members/4fb729b5-6437-4405-b499-996212603b1e",
                @"https://srosoyz.ru/sro/members/6e0b3243-be5e-437d-838b-a9cc0a4bab47",
                @"https://srosoyz.ru/sro/members/6dc9a145-e352-475a-967c-e8e1fe359957",
                @"https://srosoyz.ru/sro/members/8557a249-e339-4c02-87cb-acc5010fc7dd",
                @"https://srosoyz.ru/sro/members/ac553078-69e6-4e40-a071-03d806c43d8c",
                @"https://srosoyz.ru/sro/members/30a3d0b8-35c4-422b-a278-050e3d59fd0c", 
                @"https://srosoyz.ru/sro/members/37e14b39-835e-439e-8aa2-c8f5d2376d11", 
                @"https://srosoyz.ru/sro/members/91297dd0-9005-4d71-8f4f-d1a93ccd2972", 
            };
            string[] names = new string[]
            {
                "Белых",
                "Данченко",
                "Измакова",
                "Богуцкая",
                "Федорова",
                "Забелина",
                "Пожарская",
                "Бородкин",
                "Калошин",
                "Лошков",
                "Швец",
                "Сергеенко",
                "Пяткова",
                "Пестова",
            };


            for (int i = 0; i < names.Length; i++)
            {
                if (i == 5 || i == 9)
                {
                    var htmlDoc = web.Load(urls[i]);
                    var nodes = htmlDoc.DocumentNode.SelectNodes(infoNode);
                    var insuranse = nodes[4].SelectNodes("tr/td");

                    editor.insNumber = insuranse[3].InnerHtml;
                    editor.insDateFrom = insuranse[0].InnerHtml;
                    editor.insDateTo = insuranse[1].InnerHtml;
                }
                else
                {
                    var htmlDoc = web.Load(urls[i]);
                    var nodes = htmlDoc.DocumentNode.SelectNodes(infoNode);
                    var insuranse = nodes[5].SelectNodes("tr/td");

                    editor.insNumber = insuranse[3].InnerHtml;
                    editor.insDateFrom = insuranse[0].InnerHtml;
                    editor.insDateTo = insuranse[1].InnerHtml;
                }

                editor.EditFileOne($@"{fileDirection}{names[i]}" + ".docx", names[i]);
                editor.EditFileTwo($@"{fileDirection}{names[i]}" + ".ИНФ.docx", names[i]);
            }
        }
    }
}


