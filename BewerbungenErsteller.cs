using Microsoft.Office.Interop.Word;
using System;
using System.IO;

namespace BewerbungenErsteller
{
    internal class Program
    {
        public static string TemplatePath = @"I:\Template\";
        public static string SavePath = @"I:\Bewerbungen\";
        public static string CompanyName = "";
        public static string CompanyAddress = "";
        public static string CompanyPLZ = "";
        public static string CompanyCity = "";
        public static string CompanyRecruiterGender = "";
        public static string CompanyRecruiterName = "";
        public static string Date = DateTime.Now.ToString("dd.MM.yyyy");
        object readOnly = false;
        object isVisible = true;
        object missing = System.Reflection.Missing.Value;
        public static void Main(string[] args)
        {
            Start();
        }
        public static void Start()
        {
            Console.Clear();
            Console.Write("Bitte Gib den Firmen Namen ein: ");
            CompanyName = Console.ReadLine();
            Console.Write("\nBitte Gib die Firmen Addresse Straße ein: ");
            CompanyAddress = Console.ReadLine();
            Console.Write("\nBitte Gib die Firmen PLZ ein: ");
            CompanyPLZ = Console.ReadLine();
            Console.Write("\nBitte Gib die Firmen Stadt ein: ");
            CompanyCity = Console.ReadLine();
            Console.Write("\nBitte Gib an, ob der Ansprech Partner, Männlich oder Weiblich ist(m/w)(Für keinen, gib w an): ");
            CompanyRecruiterGender = Console.ReadLine();
            if (CompanyRecruiterGender != "")
            {
                Console.WriteLine("\nBitte Gib den Namen des Ansprech Partners ein(für Keinen, gib x an): ");
                CompanyRecruiterName = Console.ReadLine();
            }
            Validate();
        }
        public static void Validate()
        {
            Console.WriteLine("Sind Diese Infos Korrekt?");
            Console.WriteLine(CompanyName);
            Console.WriteLine(CompanyAddress + "");
            Console.WriteLine(CompanyPLZ + " " + CompanyCity);
            Console.WriteLine();
            if (CompanyRecruiterGender == "w") Console.Write("Sehr geehrte ");
            else Console.Write("Sehr geehrter ");
            if (CompanyRecruiterName == "x") Console.Write("Damen und Herren");
            else Console.Write(CompanyRecruiterName);

            Console.WriteLine();
            Console.WriteLine("(y/n)");
            if (Console.ReadLine() == "y")
            {
                Directory.CreateDirectory(SavePath + CompanyName);
                GetCV();
                MakeDocument();
                GetOtherDocuments();
            }
            else Start();
            Console.ReadKey();
        }
        public static void GetOtherDocuments()
        {
            string[] files = Directory.GetFiles(TemplatePath, "*.pdf");
            foreach (string file in files)
            {
                File.Copy(file, SavePath + CompanyName);
            }
        }
        public static void GetCV()
        {
            Application fileOpen = new Application();
            Microsoft.Office.Interop.Word.Document document = fileOpen.Documents.Open(TemplatePath + "LebenslaufTemplate.docx", ReadOnly: false);
            fileOpen.Visible = true;
            document.Activate();
            FindAndReplace(fileOpen, "{Date}", Date);
            document.SaveAs2(SavePath + "\\" + CompanyName + "\\Lebenslauf.docx");
            fileOpen.Quit();
        }
        public static void MakeDocument()
        {
            Application fileOpen = new Application();
            Microsoft.Office.Interop.Word.Document document = fileOpen.Documents.Open(TemplatePath + "AnschreibenTemplate.docx", ReadOnly: false);
            fileOpen.Visible = true;
            document.Activate();
            FindAndReplace(fileOpen, "{CompanyName}", CompanyName);
            FindAndReplace(fileOpen, "{Address-Street}", CompanyAddress);
            FindAndReplace(fileOpen, "{Address-PLZ}", CompanyPLZ);
            FindAndReplace(fileOpen, "{Address-City}", CompanyCity);
            FindAndReplace(fileOpen, "{Date}", Date);
            if(CompanyRecruiterGender == "w") FindAndReplace(fileOpen, "{m/f}", "");
            else FindAndReplace(fileOpen, "{m/f}", "r");
            if(CompanyRecruiterName != "x") FindAndReplace(fileOpen, "{Company-Recruiter-Name}", CompanyRecruiterName);
            else FindAndReplace(fileOpen, "{Company-Recruiter-Name}", "Damen und Herren");
            document.SaveAs2(SavePath + "\\" + CompanyName + "\\Anschreiben.docx");
            fileOpen.Quit();
        }
        static void FindAndReplace(Microsoft.Office.Interop.Word.Application fileOpen, object findText, object replaceWithText)
        {
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
    }
}
