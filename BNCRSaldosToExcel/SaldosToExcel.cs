using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BNCRSaldosToExcel
{
    /// <summary>
    /// Converts a BNCR saldos file into an excel file grouped by month
    /// </summary>
    public class SaldosToExcel
    {

        public SaldosToExcel(string fileName)
            : this(new StringReader(fileName))
        {

        }

        public SaldosToExcel(TextReader contents)
        {
            FileContent = SanitazeContents(contents.ReadToEnd());
            Package = new ExcelPackage();
        }

        public List<string> FileContent { get; private set; }
        public ExcelPackage Package { get; private set; }

        private List<string> SanitazeContents(string contents)
        {
            var items =  contents.Replace(",", "").Replace(";", ",").Split(new[] { '\n' });
            return items.Skip(1).ToList();
        }

        private List<BncrSaldo> GetSaldos()
        {
            return FileContent.Select(BncrSaldo.Parse).Where(i => i != null).ToList();
        }

        private void LoadWorksheet(List<BncrSaldo> saldos, ExcelWorksheet worksheet)
        {
            FormatHeaderRow(worksheet);
            for (var i = 0; i < saldos.Count; i++)
            {
                AssignSaldo(saldos[i], worksheet, i + 2);
            }
        }

        public byte[] ToSingleExcel()
        {
            var saldos = GetSaldos();
            var worksheet = Package.Workbook.Worksheets.Add("Datos");
            LoadWorksheet(saldos, worksheet);
            return Package.GetAsByteArray();
        }

        public byte[] ToExcelGroupByMonth()
        {
            var saldos = GetSaldos();
            var years = saldos.Select(y => y.Fecha.Year).Distinct();
            foreach (var year in years)
            {
                var months = saldos.Where(m => m.Fecha.Year == year).Select(m => m.Fecha.Month).Distinct();
                foreach (var month in months)
                {
                    var worksheet = Package.Workbook.Worksheets.Add((new DateTime(year, month, 1)).ToString("yyyy-MMM"));
                    var monthlySaldos = saldos.Where(s => s.Fecha.Year == year && s.Fecha.Month == month).ToList();
                    LoadWorksheet(monthlySaldos, worksheet);
                }
            }
            return Package.GetAsByteArray();
        }

        private void FormatHeaderRow(ExcelWorksheet worksheet)
        {
            var excelRow = worksheet.Row(0);
            excelRow.Style.Font.Bold = true;
            worksheet.Cells[1, 1].Value = "Fecha";
            worksheet.Cells[1, 1].Style.Numberformat.Format = "dd/mm/yyyy";
            worksheet.Cells[1, 2].Value = "Oficina";
            worksheet.Cells[1, 3].Value = "Numero Doc";
            worksheet.Cells[1, 4].Value = "Debito";
            worksheet.Cells[1, 5].Style.Numberformat.Format = "#,##0.00";
            worksheet.Cells[1, 5].Value = "Credito";
            worksheet.Cells[1, 6].Style.Numberformat.Format = "#,##0.00";
            worksheet.Cells[1, 6].Value = "Descripcion";


        }

        private void AssignSaldo(BncrSaldo saldo, ExcelWorksheet worksheet, int row)
        {
            worksheet.Cells[row, 1].Value = saldo.Fecha;
            worksheet.Cells[row, 2].Value = saldo.Oficina;
            worksheet.Cells[row, 3].Value = saldo.NumDocumento;
            worksheet.Cells[row, 4].Value = saldo.Debito;
            worksheet.Cells[row, 5].Value = saldo.Credito;
            worksheet.Cells[row, 6].Value = saldo.Descripcion;
        }

    }

    public class BncrSaldo
    {
        public string Oficina { get; set; }
        public DateTime Fecha { get; set; }
        public string NumDocumento { get; set; }
        public double Debito { get; set; }
        public double Credito { get; set; }
        public string Descripcion { get; set; }

        public static BncrSaldo Parse(string value)
        {
            var items = value.Split(",".ToArray());
            if (items.Count() < 6) return null;
            if (string.IsNullOrWhiteSpace(items[0])) return null;
            return new BncrSaldo()
                {
                    Oficina = items[0],
                    Fecha = Convert.ToDateTime(items[1]),
                    NumDocumento = items[2],
                    Debito = Convert.ToDouble(string.IsNullOrWhiteSpace(items[3]) ? "0" : items[3]),
                    Credito = Convert.ToDouble(string.IsNullOrWhiteSpace(items[4]) ? "0" : items[4]),
                    Descripcion = items[5]
                };
        }
    }
}
