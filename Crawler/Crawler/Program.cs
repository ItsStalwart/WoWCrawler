using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using HtmlAgilityPack;
using OfficeOpenXml;
using System.IO;

namespace Crawler
{

    public class gearPiece
    {
        public string itemSlot;
        public int itemLevel;
        public string itemName;
        public string itemSrc;
        public DateTime itemAcquired;
    }

    public class Character
    {
        public string name;
        public string realm;
        public List<gearPiece> charEquipment;
    }

    


    class Program
    {
        static List<gearPiece> charItems = new List<gearPiece>();
        static List<Character> Characters = new List<Character>();
        static List<string> charactersToLoad = new List<string>() {};



        static void Main(string[] args)
        {

            fetchNewData();
            //createSheet();
            Console.ReadLine();
        }

        static List<String> SlotNames = new List<String>();
        static List<Task> tasks = new List<Task>();
        private static async Task fetchNewData() {
            Console.WriteLine("Crawling started for "+ charactersToLoad.Count + " characters. Please do not close or shutdown the program. Data will not be saved until crawling is finished.");
            foreach (string t in charactersToLoad)
            {
              tasks.Add(StartCrawlerAsync(t));
            }
            Task.WaitAll(tasks.ToArray());
            createSheet("yourFileName");
        }
                private static async Task StartCrawlerAsync(string url)
        {
            var httpClient = new HttpClient();
            var html = await httpClient.GetStringAsync(url);

            var htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(html);

            var slots = htmlDocument.DocumentNode.Descendants("div")
                .Where(node => node.GetAttributeValue("class","").Equals("slotHeader")).ToList();

            var itemTable = htmlDocument.DocumentNode.Descendants("div")
                .Where(node => node.GetAttributeValue("class", "").Equals("slotItems")).ToList();

            var findOwner = htmlDocument.DocumentNode.Descendants("h1").FirstOrDefault().InnerText;

            var findRealm = htmlDocument.DocumentNode.Descendants("a").
                Where(node => node.GetAttributeValue("class", "").Equals("nav_link")).FirstOrDefault().InnerText;

            var provChar = new Character();
            provChar.name = findOwner;
            provChar.realm = findRealm;
            provChar.charEquipment = new List<gearPiece>();

            foreach (var slot in itemTable)
            {
                gearPiece equipment = new gearPiece();
                equipment.itemLevel = Int32.Parse(slot.InnerText.Substring(0, 3));
                equipment.itemName = slot.InnerText.Substring(3);
                String newUrl = slot.Descendants("a").FirstOrDefault().ChildAttributes("href").FirstOrDefault().Value;
                var secondaryClient = new HttpClient();
                var secondaryHtml = await secondaryClient.GetStringAsync(newUrl);

                var secondaryDocument = new HtmlDocument();
                secondaryDocument.LoadHtml(secondaryHtml);

                var findSource = secondaryDocument.DocumentNode.Descendants("dd")
                    .Where(node => node.GetAttributeValue("style", "").Equals("color: #00FF00")).ToList();

                var findSlot = secondaryDocument.DocumentNode.Descendants("dd")
                    .Where(node => node.GetAttributeValue("class", "").Equals("db-left")).FirstOrDefault().InnerText;
                equipment.itemSlot = findSlot;

                if (findSource.Count == 0)
                {
                    equipment.itemSrc = "Normal";
                }
                else
                {
                    equipment.itemSrc = findSource[0].InnerText;
                }
                equipment.itemAcquired = DateTime.Now;
                provChar.charEquipment.Add(equipment);         
            }
            Characters.Add(provChar);
            Console.WriteLine("Succesfully Crawled " + Characters.Count + " characters so far.");
            Console.WriteLine((charactersToLoad.Count - Characters.Count) + " remain to be crawled on this session.");

        }

        public static void createSheet(string fileName)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Worksheet1");
                

                var headerRow = new List<string[]>()
                 {
                    new string[] { "Owner Character", "Item Name", "Item Slot", "Difficulty Tag", "Item Level", "Date of Acquisition" }
                 };

                
                
                // Determine the header range (e.g. A1:D1)
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                var linha = new List<string[]>();

                foreach (Character boneco in Characters)
                {
                    for (int i = 0; i < boneco.charEquipment.Count; i++)
                    {
                        
                            

                            linha.Add(new string[] { boneco.name + " @ " + boneco.realm, boneco.charEquipment[i].itemName, boneco.charEquipment[i].itemSlot, boneco.charEquipment[i].itemSrc, boneco.charEquipment[i].itemLevel.ToString(), boneco.charEquipment[i].itemAcquired.ToString() });

                            string referencia = "A2:" + Char.ConvertFromUtf32(linha[0].Length + 64) + "2";

                           // Console.ReadLine();

                            worksheet.Cells[referencia].LoadFromArrays(linha);

                        

                    }

                }

                    // Popular header row data
                    worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                FileInfo excelFile = new FileInfo(@"C:\Users\luan_\Desktop\"+  fileName + ".xlsx");
                excel.SaveAs(excelFile);
                Console.WriteLine("Crawling Finished. Program may now be stopped.");
            }
        }


    }

}

