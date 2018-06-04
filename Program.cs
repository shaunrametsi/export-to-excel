using Simple.ToExcel.Classes;
using Simple.ToExcel.ExcelFunctions;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace EPPLUS_1._1
{
    class Program
    {
        private static void Main(string[] args)
        {
            /* Export Pirates to Excel */
            Display("Pirates", () => { ExportPirates();});
            /* Export Marines to Excel */
            Console.WriteLine("\n");
            Display("Marines",()=> { ExportMarines();});
            Thread.Sleep(10000);

        }
        public static void Display(string type, Action function)
        {
            WriteColor($"********************************************************\n*{type}\n********************************************************",ConsoleColor.DarkRed);
            Console.ForegroundColor = ConsoleColor.White;
            function();
            WriteColor("\nexported!",ConsoleColor.Blue);
        }

        public static void WriteColor(string text, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(text);
        }
        public async static Task ExportPirates()
        {
            var pirates = new List<Pirate>()
            {
                new Pirate("Luffy","Mugiwara",500000000,false,"Schichibukai"),
                new Pirate("Zoro","Kaizoku-Gari",320000000,false,"Unknown"),
                new Pirate("Ussop","Sogeking",177000000,false,"Unknown"),
                new Pirate("Sanji","Kuroashi",177000000,false,"Unknown"),
                new Pirate("Nami","Neko no dorobo",177000000,false,"Unknown"),
                new Pirate("Chopper","Kottonkyandiraba",177000000,false,"Unknown"),
                new Pirate("Robin","Akuma no ko",177000000,false,"Unknown"),
                new Pirate("Franky","Cyborg",177000000,false,"Unknown"),
                new Pirate("Brook","Soul King",177000000,false,"Unknown")
            };

            var excel = new ReadAndWrite<Pirate>();
            excel.Content = pirates;
            excel.Path = "pirates.xlsx";

            await excel.Export();
            await excel.ReadFromExcel("pirates.xlsx");

        }
        public async static Task ExportMarines()
        {
            var marines = new List<Marine>()
            {
                new Marine("Garp","The Fist","Vice Admiral[ex]","Headquartes"),
                new Marine("Sengoku","Great Buddah","Fleet Admiral[ex]","Headquartes"),
                new Marine("Akainu","Red Dog","Fleet Admiral","Headquartes"),
                new Marine("Aokiji","Blue pheasant","Admiral[ex]","Headquarters"),
                new Marine("Smoker","Chaser","Vice Admiral","New World"),
                new Marine("Tashigi","--Unknown--","Captain","New World"),
                new Marine("Coby","--Unknown--","--Unknown--","--Unknown--")
            };

            var excel = new ReadAndWrite<Marine>();
            excel.Content = marines;
            excel.Path = "marines.xlsx";

            await excel.Export();
            await excel.ReadFromExcel("marines.xlsx");

        }

    }
}