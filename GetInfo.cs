using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParS
{
    internal class GetHtml
    {
        static HtmlWeb web = new HtmlWeb();
        static string infoNode = "//table[@class='uk-table uk-table-small bordered uk-table-divider uk-table-striped']";
        public static void Start(int i)
        {
            if (i == 5 || i == 9)
            {
                var htmlDoc = web.Load(Raters.rater[i].url);
                var nodes = htmlDoc.DocumentNode.SelectNodes(infoNode);
                var insuranse = nodes[4].SelectNodes("tr/td");

                WordEditor.insNumber = insuranse[3].InnerHtml;
                WordEditor.insDateFrom = insuranse[0].InnerHtml;
                WordEditor.insDateTo = insuranse[1].InnerHtml;
            }
            else
            {
                var htmlDoc = web.Load(Raters.rater[i].url);
                var nodes = htmlDoc.DocumentNode.SelectNodes(infoNode);
                var insuranse = nodes[5].SelectNodes("tr/td");

                WordEditor.insNumber = insuranse[3].InnerHtml;
                WordEditor.insDateFrom = insuranse[0].InnerHtml;
                WordEditor.insDateTo = insuranse[1].InnerHtml;
            }
        }
    }
}
