using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using System.Xml.Linq;

namespace KreiranjeNaloga
{
    internal class Program
    {
        private const string OutputFileName = "rezultat.xml";
        private static string? _companyName;
        private static string? _companyCity;
        private static string? _accountId;
        private static string? _bankId;
        private static string? _bankName;
        private static XElement? _companyInfo;
        private static XElement? _accountinfo;

        static void Main(string[] args)
        {
            Init();

            var inputFileName = FindExcelSheetFile();
            if (string.IsNullOrEmpty(inputFileName))
            {
                Console.WriteLine($"Nije pronadjen nijedan *.xlsx fajl");
                Console.ReadKey();
                return;
            }

            if (!TryOpenSheet(inputFileName, out var wb) || wb == null)
            {
                Console.WriteLine($"Nije moguce otvoriti fajl {inputFileName}");
                Console.ReadKey();
                return;
            }
            var ws = wb.Worksheets.FirstOrDefault();
            var firstRowUsed = ws.FirstRowUsed();
            var row = firstRowUsed.RowBelow();
            
            var xml = new XDocument();
            xml.Declaration = new XDeclaration("1.0", "UTF-8", "");
            var paymentOrders = new XElement("pmtorderrq");

            do
            {
                var cell = row.FirstCellUsed();
                var naziv = cell.Value.ToString();
                var adresa = (cell = cell.CellRight()).Value.ToString();
                var tekuci = (cell = cell.CellRight()).Value.ToString();
                var sifraUplate = (cell = cell.CellRight()).Value.ToString();
                var model = (cell = cell.CellRight()).Value.ToString();
                var pozivNaBroj = (cell = cell.CellRight()).Value.ToString();
                var iznos = (cell = cell.CellRight()).Value.ToString() ?? "0.00";
                var svrhaUplate = (cell = cell.CellRight()).Value.ToString();

                if(!TryFixAndValidateBankAccountNumber(ref tekuci))
                {
                    Console.WriteLine($"Broj racuna {tekuci} za primaoca {naziv} nije ispravan");
                    Console.ReadKey();
                    return;
                }

                iznos = iznos.Contains('.') ? iznos : $"{iznos}.00";

                XElement paymentOrder = CreateOrder(naziv, adresa, tekuci, sifraUplate, model, pozivNaBroj, iznos, svrhaUplate);

                paymentOrders.Add(paymentOrder);
            }
            while (!(row = row.RowBelow()).IsEmpty() && !row.FirstCellUsed().HasFormula);

            xml.Add(paymentOrders);

            WriteToFile(xml);

            Console.Write("Pritisnuti bilo koji taster za kraj: ");
            Console.ReadKey();
        }

        private static bool TryFixAndValidateBankAccountNumber(ref string? tekuci)
        {
            if (string.IsNullOrEmpty(tekuci))
            {
                return false;
            }
            var parts = tekuci.Split('-');
            if (parts.Length != 3)
            {
                return false;
            }
            if (parts[1].Length < 13)
            {
                tekuci = AddLeadingZeroes(parts);
            }

            return true;
        }

        private static string? AddLeadingZeroes(string[] parts)
        {
            var brojNulaKojiNedostaje = 13 - parts[1].Length;
            var srednjiDeoTekucegRacuna = parts[1];
            for (int i = 0; i < brojNulaKojiNedostaje; i++)
            {
                srednjiDeoTekucegRacuna = $"0{srednjiDeoTekucegRacuna}";
            }

            return $"{parts[0]}-{srednjiDeoTekucegRacuna}-{parts[2]}";
        }

        private static bool TryOpenSheet(string inputFileName, out XLWorkbook? wb)
        {
            wb = null;
            try
            {
                wb = new XLWorkbook(inputFileName);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }

        private static XElement CreateOrder(object naziv, object adresa, object tekuci, object sifraUplate, object model, object pozivNaBroj, object iznos, object svrhaUplate)
        {
            var payeecompanyinfo = new XElement("payeecompanyinfo",
                new XElement("name", naziv),
                new XElement("city", adresa)
            );

            var payeeaccountinfo = new XElement("payeeaccountinfo",
                new XElement("acctid", tekuci),
                new XElement("bankid", ((string)tekuci).Substring(0, 3)),
                new XElement("bankname", $"{naziv} {adresa}")
            );

            var paymentOrder = new XElement("pmtorder",
                _companyInfo,
                _accountinfo,
                payeecompanyinfo,
                payeeaccountinfo,
                new XElement("trnuid", ""),
                new XElement("dtdue", DateOnly.FromDateTime(DateTime.Now).ToString("yyyy-MM-dd")), // "2022-11-04"
                new XElement("trnamt", iznos),
                new XElement("trnplace", "online"),
                new XElement("purpose", svrhaUplate),
                new XElement("purposecode", sifraUplate),
                new XElement("curdef", "RSD"),
                new XElement("refmodel", ""),
                new XElement("refnumber", ""),
                new XElement("payeerefmodel", model),
                new XElement("payeerefnumber", pozivNaBroj),
                new XElement("urgency", "ACH"),
                new XElement("priority", "50")
            );
            return paymentOrder;
        }

        private static string? FindExcelSheetFile()
        {
            var currentDirectoryName = Directory.GetCurrentDirectory();
            string[] files = Directory.GetFiles(currentDirectoryName, "*.xlsx");
            var inputFileName = files.FirstOrDefault();
            Console.WriteLine($"Naziv pronadjenog fajla: {inputFileName}");
            return inputFileName;
        }

        private static void Init()
        {
            var configuration = new ConfigurationBuilder()
                .AddJsonFile($"appsettings.json");

            var config = configuration.Build();

            _companyName = config.GetSection("companyinfo").GetSection("name").Value;
            _companyCity = config.GetSection("companyinfo").GetSection("city").Value;
            _accountId = config.GetSection("accountinfo").GetSection("acctid").Value;
            _bankId = config.GetSection("accountinfo").GetSection("bankid").Value;
            _bankName = config.GetSection("accountinfo").GetSection("bankname").Value;

            _companyInfo =
                new XElement("companyinfo",
                    new XElement("name", _companyName),
                    new XElement("city", _companyCity)
                );
            _accountinfo =
                new XElement("accountinfo",
                    new XElement("acctid", _accountId),
                    new XElement("bankid", _bankId),
                    new XElement("bankname", _bankName)
                );
        }

        private static void WriteToFile(XDocument xml)
        {
            try
            {
                var file = File.Open(OutputFileName, FileMode.OpenOrCreate);
                xml.Save(file);
                Console.WriteLine($"Fajl je uspesno kreiran: \"{OutputFileName}\"");
            }
            catch (Exception)
            {
                Console.WriteLine($"Greska prilikom kreiranja fajla \"{OutputFileName}\"");
                Console.ReadKey();
            }
        }
    }
}