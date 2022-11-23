using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using System.Reflection;
using System.Xml.Linq;

namespace KreiranjeNaloga
{
    internal class Program
    {
        private const string OutputFileName = "nalog_za_uplatu.xml";
        private static string? _companyName;
        private static string? _companyCity;
        private static string? _accountId;
        private static string? _bankId;
        private static string? _bankName;
        private static XElement? _companyInfo;
        private static XElement? _accountinfo;

        static async Task Main(string[] args)
        {
            Init();

            var inputFileName = FindExcelSheetFile();
            if (string.IsNullOrEmpty(inputFileName))
            {
                Console.WriteLine($"Nije pronadjen nijedan *.xlsx fajl");
                Console.ReadKey();
                return;
            }

            if (!TryOpenSheet(inputFileName, out var wb))
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
                var naziv = cell.Value;
                var adresa = (cell = cell.CellRight()).Value;
                var tekuci = (cell = cell.CellRight()).Value;
                var sifraUplate = (cell = cell.CellRight()).Value;
                var model = (cell = cell.CellRight()).Value;
                var pozivNaBroj = (cell = cell.CellRight()).Value;
                var iznos = (cell = cell.CellRight()).Value;
                var svrhaUplate = (cell = cell.CellRight()).Value;

                XElement paymentOrder = CreateOrder(naziv, adresa, tekuci, sifraUplate, model, pozivNaBroj, iznos, svrhaUplate);

                paymentOrders.Add(paymentOrder);
            }
            while (!(row = row.RowBelow()).IsEmpty() && !row.FirstCellUsed().HasFormula);

            xml.Add(paymentOrders);

            WriteToFile(xml);
        }

        private static bool TryOpenSheet(string inputFileName, out XLWorkbook? wb)
        {
            wb = null;
            try
            {
                wb = new XLWorkbook("TABELA PLAĆANJA novembar.xlsx");
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
            var currentDirectoryName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string[] files = Directory.GetFiles(currentDirectoryName, "*.xlsx");
            var inputFileName = files.FirstOrDefault();
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
                Console.ReadKey();
            }
            catch (Exception)
            {
                Console.WriteLine($"Greska prilikom kreiranja fajla \"{OutputFileName}\"");
                Console.ReadKey();
            }
        }

        private static XDocument CreateXML(XElement? companyInfo, XElement? accountinfo)
        {
            return new XDocument(
                new XDeclaration("1.0", "UTF-8", ""),
                new XElement("pmtorderrq",
                    new XElement("pmtorder",
                        companyInfo,
                        accountinfo,
                        new XElement("payeecompanyinfo",
                            new XElement("name", "TEST"),
                            new XElement("city", "GRAD")
                        ),
                        new XElement("payeeaccountinfo",
                            new XElement("acctid", "330-1111111111111-58"),
                            new XElement("bankid", "330"),
                            new XElement("bankname", "TEST TEST")
                        ),
                        new XElement("trnuid", ""),
                        new XElement("dtdue", "2022-11-04"),
                        new XElement("trnamt", "12345.50"),
                        new XElement("trnplace", "online"),
                        new XElement("purpose", "Test"),
                        new XElement("purposecode", "248"),
                        new XElement("curdef", "RSD"),
                        new XElement("refmodel", ""),
                        new XElement("refnumber", ""),
                        new XElement("payeerefmodel", "97"),
                        new XElement("payeerefnumber", "3591000000040654654"),
                        new XElement("urgency", "ACH"),
                        new XElement("priority", "50")
                    )
                )
            );
        }
    }
}