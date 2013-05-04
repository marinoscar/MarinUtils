using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;

namespace BNCRSaldosToExcel.Tests
{
    [TestFixture]
    public class WhenFileIsProcessed
    {


        private StreamReader _stream;

        [TestFixtureSetUp]
        public void TestSetup()
        {
            _stream = new StreamReader(@".\Tests\ConsultaSaldos.csv");
        }

        [Test]
        public void ItShouldCreateAValidExcelFile()
        {
            var excel = new SaldosToExcel(_stream);
            var file = new FileStream("SingleExcel.xlsx", FileMode.OpenOrCreate);
            var bits = excel.ToSingleExcel();
            file.Write(bits, 0, bits.Count());          
        }

        [Test]
        public void ItShouldCreateAValidExcelGroupByMonthFile()
        {
            var excel = new SaldosToExcel(_stream);
            var file = new FileStream("MonthlyExcel.xlsx", FileMode.OpenOrCreate);
            var bits = excel.ToExcelGroupByMonth();
            file.Write(bits, 0, bits.Count());
        }

        [TestFixtureTearDown]
        public void TestTearDown()
        {
            _stream.Close();
        }
    }
}
