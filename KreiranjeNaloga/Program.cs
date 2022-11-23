using Microsoft.Extensions.Configuration;
using System.Xml.Linq;

namespace KreiranjeNaloga
{
    internal class Program
    {
        private const string Path = "nalog_za_uplatu.xml";
        private static string? _companyName;
        private static string? _companyCity;
        private static string? _accountId;
        private static string? _bankId;
        private static string? _bankName;
        private static XElement _companyInfo;
        private static XElement _accountinfo;

        static async Task Main(string[] args)
        {
            Init();

            // TODO: Read and parse excel sheet

            XDocument xml = CreateXML(_companyInfo, _accountinfo);

            WriteToFile(xml);
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
                var file = File.Open(Path, FileMode.OpenOrCreate);
                xml.Save(file);
                Console.WriteLine($"Fajl je uspesno kreiran: \"{Path}\"");
            }
            catch (Exception)
            {
                Console.WriteLine($"Greska prilikom kreiranja fajla \"{Path}\"");
            }
        }

        private static XDocument CreateXML(XElement companyInfo, XElement accountinfo)
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