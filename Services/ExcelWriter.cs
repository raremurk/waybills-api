using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Globalization;
using WaybillsAPI.Interfaces;
using WaybillsAPI.Models;
using WaybillsAPI.ReportsModels.CostPrice;
using WaybillsAPI.ReportsModels.MonthTotal;

namespace WaybillsAPI.Services
{
    public class ExcelWriter : IExcelWriter
    {
        public byte[] GenerateCostPriceReport(CostPriceReport report, int year, int month)
        {
            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Себестоимость");
            SetPrinterMargins(sheet);
            sheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells.Style.Font.Size = 14;
            sheet.Row(1).Height = 54;
            sheet.Row(2).Height = 21;
            sheet.Column(1).Width = GetTrueColumnWidth(15);
            sheet.Column(2).Width = GetTrueColumnWidth(35);
            sheet.Column(3).Width = GetTrueColumnWidth(35);
            sheet.Column(2).Style.Numberformat.Format = "0.00";
            sheet.Column(3).Style.Numberformat.Format = "0.00";

            sheet.Cells["A1:C1"].Merge = true;
            sheet.Cells["A1"].Style.Font.Size = 18;
            sheet.Cells["A1:C2"].Style.Font.Bold = true;
            sheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["B2:C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            sheet.Cells["A2:C2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            sheet.Cells["A1"].Value = $"{CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(month)} {year} по путевым листам трактористов";
            sheet.Cells["A2"].Value = "ШПЗ";
            sheet.Cells["B2"].Value = "Условный эталонный гектар";
            sheet.Cells["C2"].Value = "Себестоимость";

            var index = 3;
            foreach (var costCode in report.CostCodes)
            {
                sheet.Row(index).Height = 21;
                sheet.Cells[$"A{index}"].Value = costCode.ProductionCostCode;
                sheet.Cells[$"B{index}"].Value = costCode.ConditionalReferenceHectares;
                sheet.Cells[$"C{index}"].Value = costCode.CostPrice;
                index++;
            }
            sheet.Row(index).Height = 21;
            sheet.Cells[$"A{index}"].Value = "Итого";
            sheet.Cells[$"B{index}"].Value = report.ConditionalReferenceHectares;
            sheet.Cells[$"C{index}"].Value = report.CostPrice;
            sheet.Cells[$"A{index}:C{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            sheet.Cells[$"A{index}:C{index}"].Style.Font.Bold = true;

            return package.GetAsByteArray();
        }

        public byte[] GenerateMonthTotal(List<Waybill> waybills)
        {

            var monthTotalReportByDrivers = new MonthTotalReport(waybills);
            var monthTotalReportByTransports = new MonthTotalReport(waybills, false);

            var package = new ExcelPackage();
            AddMonthTotalReportSheet(monthTotalReportByDrivers, true);
            AddMonthTotalReportSheet(monthTotalReportByTransports, false);

            void AddMonthTotalReportSheet(MonthTotalReport monthTotalReport, bool byDrivers)
            {
                var headers = byDrivers ? ("По водителям", "Таб.", "Водитель", "Гар.", "Транспорт") : ("По транспорту", "Гар.", "Транспорт", "Таб.", "Водитель");
                var sheet = package.Workbook.Worksheets.Add(headers.Item1);
                SetPrinterMargins(sheet);
                sheet.PrinterSettings.Orientation = eOrientation.Landscape;
                sheet.PrinterSettings.PaperSize = ePaperSize.A4;
                sheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells.Style.Font.Size = 8;
                sheet.Row(1).Height = 9;
                sheet.Row(2).Height = 9;
                sheet.Column(1).Width = GetTrueColumnWidth(4);
                sheet.Column(2).Width = GetTrueColumnWidth(15);
                sheet.Column(3).Width = GetTrueColumnWidth(4);
                sheet.Column(4).Width = GetTrueColumnWidth(15);
                sheet.Column(5).Width = GetTrueColumnWidth(3);
                sheet.Column(6).Width = GetTrueColumnWidth(3);
                sheet.Column(7).Width = GetTrueColumnWidth(5);
                sheet.Column(8).Width = GetTrueColumnWidth(7);
                sheet.Column(9).Width = GetTrueColumnWidth(7);
                sheet.Column(10).Width = GetTrueColumnWidth(7);
                sheet.Column(11).Width = GetTrueColumnWidth(5);
                sheet.Column(12).Width = GetTrueColumnWidth(6);
                sheet.Column(13).Width = GetTrueColumnWidth(6);
                sheet.Column(14).Width = GetTrueColumnWidth(9);
                sheet.Column(15).Width = GetTrueColumnWidth(6);
                sheet.Column(16).Width = GetTrueColumnWidth(7);
                sheet.Column(17).Width = GetTrueColumnWidth(5);
                sheet.Column(18).Width = GetTrueColumnWidth(5);

                sheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(7).Style.Numberformat.Format = "0.0";
                sheet.Column(8).Style.Numberformat.Format = "0.00";
                sheet.Column(9).Style.Numberformat.Format = "0.00";
                sheet.Column(10).Style.Numberformat.Format = "0.00";
                sheet.Column(12).Style.Numberformat.Format = "0.0";
                sheet.Column(13).Style.Numberformat.Format = "0.0";
                sheet.Column(14).Style.Numberformat.Format = "0.000";
                sheet.Column(15).Style.Numberformat.Format = "0.00";
                sheet.Column(16).Style.Numberformat.Format = "0.00";

                var index = 1;
                sheet.Cells[$"A{index}:R{index + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                sheet.Cells[$"A{index}:R{index + 1}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index}:R{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"Q{index}:R{index + 1}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index + 1}:R{index + 1}, L{index}:M{index}, Q{index}:R{index}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                sheet.Cells[$"B{index}:B{index + 1}"].Merge = true;
                sheet.Cells[$"D{index}:D{index + 1}"].Merge = true;
                sheet.Cells[$"F{index}:F{index + 1}"].Merge = true;
                sheet.Cells[$"G{index}:G{index + 1}"].Merge = true;
                sheet.Cells[$"H{index}:H{index + 1}"].Merge = true;
                sheet.Cells[$"I{index}:I{index + 1}"].Merge = true;
                sheet.Cells[$"J{index}:J{index + 1}"].Merge = true;
                sheet.Cells[$"L{index}:M{index}, Q{index}:R{index}"].Merge = true;

                sheet.Cells[$"A{index}"].Value = headers.Item2;
                sheet.Cells[$"B{index}"].Value = headers.Item3;
                sheet.Cells[$"C{index}"].Value = headers.Item4;
                sheet.Cells[$"D{index}"].Value = headers.Item5;
                sheet.Cells[$"A{index + 1}, C{index + 1}"].Value = "№";
                sheet.Cells[$"E{index}"].Value = "Пут.";
                sheet.Cells[$"E{index + 1}"].Value = "лист";
                sheet.Cells[$"F{index}"].Value = "Дни";
                sheet.Cells[$"G{index}"].Value = "Часы";
                sheet.Cells[$"H{index}"].Value = "Сумма";
                sheet.Cells[$"I{index}"].Value = "Премия";
                sheet.Cells[$"J{index}"].Value = "Выходные";
                sheet.Cells[$"K{index}"].Value = "Число";
                sheet.Cells[$"K{index + 1}"].Value = "ездок";
                sheet.Cells[$"L{index}"].Value = "Пробег";
                sheet.Cells[$"L{index + 1}"].Value = "Всего";
                sheet.Cells[$"M{index + 1}"].Value = "С грузом";
                sheet.Cells[$"N{index}"].Value = "Перевезено";
                sheet.Cells[$"N{index + 1}"].Value = "груза, тонн";
                sheet.Cells[$"O{index}"].Value = "Нормо-";
                sheet.Cells[$"O{index + 1}"].Value = "смена";
                sheet.Cells[$"P{index}"].Value = "Условный";
                sheet.Cells[$"P{index + 1}"].Value = "гектар";
                sheet.Cells[$"Q{index}"].Value = "Расход ГСМ";
                sheet.Cells[$"Q{index + 1}"].Value = "Факт";
                sheet.Cells[$"R{index + 1}"].Value = "Норма";

                index++;
                foreach (var monthTotal in monthTotalReport.DetailedEntityMonthTotals)
                {
                    index++;

                    sheet.Cells[$"A{index}"].Value = monthTotal.EntityCode;
                    sheet.Cells[$"B{index}"].Value = monthTotal.EntityName;
                    foreach (var subTotal in monthTotal.SubTotals)
                    {
                        sheet.Cells[$"C{index}"].Value = subTotal.EntityCode;
                        sheet.Cells[$"D{index}"].Value = subTotal.EntityName;
                        sheet.Row(index).Height = 9;
                        PrintMonthTotalValues(subTotal);
                        index++;
                    }

                    if (monthTotal.SubTotals.Count > 1)
                    {
                        sheet.Cells[$"C{index}:D{index}"].Merge = true;
                        sheet.Cells[$"C{index}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[$"C{index}:R{index}"].Style.Font.Bold = true;
                        sheet.Cells[$"C{index}:R{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        sheet.Cells[$"C{index}"].Value = "Итого";
                        sheet.Row(index).Height = 9;
                        PrintMonthTotalValues(monthTotal);
                        index++;
                    }
                    sheet.Row(index).Height = 4.5;
                }
                index++;
                sheet.Row(index).Height = 9;
                sheet.Cells[$"A{index}:R{index}"].Style.Font.Bold = true;
                sheet.Cells[$"A{index}:R{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index}"].Value = "Всего";
                PrintMonthTotalValues(monthTotalReport);

                void PrintMonthTotalValues(MonthTotal total)
                {
                    sheet.Cells[$"E{index}"].Value = total.WaybillsCount;
                    sheet.Cells[$"F{index}"].Value = total.Days;
                    sheet.Cells[$"G{index}"].Value = total.Hours;
                    sheet.Cells[$"H{index}"].Value = total.Earnings;
                    sheet.Cells[$"I{index}"].Value = total.Bonus;
                    sheet.Cells[$"J{index}"].Value = total.Weekend;
                    sheet.Cells[$"K{index}"].Value = total.NumberOfRuns;
                    sheet.Cells[$"L{index}"].Value = total.TotalMileage;
                    sheet.Cells[$"M{index}"].Value = total.TotalMileageWithLoad;
                    sheet.Cells[$"N{index}"].Value = total.TransportedLoad;
                    sheet.Cells[$"O{index}"].Value = total.NormShift;
                    sheet.Cells[$"P{index}"].Value = total.ConditionalReferenceHectares;
                    sheet.Cells[$"Q{index}"].Value = total.FactFuelConsumption;
                    sheet.Cells[$"R{index}"].Value = total.NormalFuelConsumption;
                }
            }
            return package.GetAsByteArray();
        }

        public byte[] GenerateShortWaybills(List<Waybill> waybills)
        {
            var groupedWaybills = waybills.GroupBy(x => x.DriverId).OrderBy(x => x.First().Driver.LastName);

            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Путевые листы кратко");
            SetPrinterMargins(sheet);
            sheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells.Style.Font.Name = "Times New Roman";
            sheet.Cells.Style.Font.Size = 7;
            sheet.Column(1).Width = GetTrueColumnWidth(2.86);
            sheet.Column(2).Width = GetTrueColumnWidth(11.57);
            sheet.Column(3).Width = GetTrueColumnWidth(3);
            sheet.Column(4).Width = GetTrueColumnWidth(2.71);
            sheet.Column(5).Width = GetTrueColumnWidth(4.29);
            sheet.Column(6).Width = GetTrueColumnWidth(2);
            sheet.Column(7).Width = GetTrueColumnWidth(3.29);
            sheet.Column(8).Width = GetTrueColumnWidth(5);
            sheet.Column(9).Width = GetTrueColumnWidth(3.29);
            sheet.Column(10).Width = GetTrueColumnWidth(4.86);
            sheet.Column(11).Width = GetTrueColumnWidth(4.86);
            sheet.Column(12).Width = GetTrueColumnWidth(7.86);
            sheet.Column(13).Width = GetTrueColumnWidth(5);
            sheet.Column(14).Width = GetTrueColumnWidth(6);
            sheet.Column(15).Width = GetTrueColumnWidth(3.57);
            sheet.Column(16).Width = GetTrueColumnWidth(3.57);
            sheet.Column(17).Width = GetTrueColumnWidth(3.57);

            sheet.Column(7).Style.Numberformat.Format = "0.0";
            sheet.Column(8).Style.Numberformat.Format = "0.00";
            sheet.Column(10).Style.Numberformat.Format = "0.0";
            sheet.Column(11).Style.Numberformat.Format = "0.0";
            sheet.Column(12).Style.Numberformat.Format = "0.000";
            sheet.Column(13).Style.Numberformat.Format = "0.00";
            sheet.Column(14).Style.Numberformat.Format = "0.00";

            sheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            sheet.DefaultRowHeight = 9;

            var index = 1;
            PrintHeader();
            var pageHeight = 18;
            var maxPageHeight = 838;

            foreach (var driverWaybills in groupedWaybills)
            {
                var driverWaybillsHeight = driverWaybills.Count() * 9;
                if (pageHeight + driverWaybillsHeight > maxPageHeight)
                {
                    sheet.Row(index).PageBreak = true;
                    pageHeight = 18 + driverWaybillsHeight;
                    index++;
                    PrintHeader();
                }
                else
                {
                    pageHeight += driverWaybillsHeight + 9;
                }
                index += 2;

                foreach (var waybill in driverWaybills)
                {
                    sheet.Cells[$"A{index}"].Value = waybill.Driver.PersonnelNumber;
                    sheet.Cells[$"B{index}"].Value = waybill.Driver.ShortFullName();
                    sheet.Cells[$"C{index}"].Value = waybill.Transport.Code;
                    sheet.Cells[$"D{index}"].Value = waybill.Date.ToString("dd.MM");
                    sheet.Cells[$"E{index}"].Value = waybill.Number;
                    sheet.Cells[$"F{index}"].Value = waybill.Days;
                    sheet.Cells[$"G{index}"].Value = waybill.Hours;
                    sheet.Cells[$"H{index}"].Value = waybill.Earnings + waybill.Bonus + waybill.Weekend;
                    sheet.Cells[$"I{index}"].Value = waybill.Operations.Sum(x => x.NumberOfRuns);
                    sheet.Cells[$"J{index}"].Value = waybill.Operations.Sum(x => x.TotalMileage);
                    sheet.Cells[$"K{index}"].Value = waybill.Operations.Sum(x => x.TotalMileageWithLoad);
                    sheet.Cells[$"L{index}"].Value = waybill.Operations.Sum(x => x.TransportedLoad);
                    sheet.Cells[$"M{index}"].Value = waybill.Operations.Sum(x => x.NormShift);
                    sheet.Cells[$"N{index}"].Value = waybill.Operations.Sum(x => x.ConditionalReferenceHectares);
                    sheet.Cells[$"O{index}"].Value = waybill.FactFuelConsumption;
                    sheet.Cells[$"P{index}"].Value = waybill.NormalFuelConsumption;
                    sheet.Cells[$"Q{index}"].Value = waybill.NormalFuelConsumption - waybill.FactFuelConsumption;
                    index++;
                }

                var total = new MonthTotal(driverWaybills);
                sheet.Cells[$"A{index}"].Value = "Итого";
                sheet.Cells[$"E{index}"].Value = total.WaybillsCount;
                sheet.Cells[$"F{index}"].Value = total.Days;
                sheet.Cells[$"G{index}"].Value = total.Hours;
                sheet.Cells[$"H{index}"].Value = total.Earnings + total.Bonus + total.Weekend;
                sheet.Cells[$"I{index}"].Value = total.NumberOfRuns;
                sheet.Cells[$"J{index}"].Value = total.TotalMileage;
                sheet.Cells[$"K{index}"].Value = total.TotalMileageWithLoad;
                sheet.Cells[$"L{index}"].Value = total.TransportedLoad;
                sheet.Cells[$"M{index}"].Value = total.NormShift;
                sheet.Cells[$"N{index}"].Value = total.ConditionalReferenceHectares;
                sheet.Cells[$"O{index}"].Value = total.FactFuelConsumption;
                sheet.Cells[$"P{index}"].Value = total.NormalFuelConsumption;
                sheet.Cells[$"Q{index}"].Value = total.NormalFuelConsumption - total.FactFuelConsumption;

                sheet.Cells[$"A{index}:Q{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index}:Q{index}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }
            return package.GetAsByteArray();

            void PrintHeader()
            {
                sheet.Cells[$"A{index}:Q{index + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                sheet.Cells[$"A{index}:Q{index + 1}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index}:Q{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"Q{index}:Q{index + 1}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index + 1}:Q{index + 1}, J{index}:K{index}, O{index}:Q{index}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                sheet.Cells[$"J{index}:K{index}, O{index}:Q{index}"].Merge = true;
                sheet.Cells[$"B{index}:B{index + 1}"].Merge = true;
                sheet.Cells[$"C{index}:C{index + 1}"].Merge = true;
                sheet.Cells[$"D{index}:D{index + 1}"].Merge = true;
                sheet.Cells[$"F{index}:F{index + 1}"].Merge = true;
                sheet.Cells[$"G{index}:G{index + 1}"].Merge = true;
                sheet.Cells[$"H{index}:H{index + 1}"].Merge = true;
                sheet.Cells[$"A{index}"].Value = "Таб.";
                sheet.Cells[$"A{index + 1}"].Value = "№";
                sheet.Cells[$"B{index}"].Value = "ФИО";
                sheet.Cells[$"C{index}"].Value = "Шифр";
                sheet.Cells[$"D{index}"].Value = "Дата";
                sheet.Cells[$"E{index}"].Value = "Номер";
                sheet.Cells[$"E{index + 1}"].Value = "листа";
                sheet.Cells[$"F{index}"].Value = "Дни";
                sheet.Cells[$"G{index}"].Value = "Часы";
                sheet.Cells[$"H{index}"].Value = "Зарплата";
                sheet.Cells[$"I{index}"].Value = "Число";
                sheet.Cells[$"I{index + 1}"].Value = "ездок";
                sheet.Cells[$"J{index}"].Value = "Пробег";
                sheet.Cells[$"J{index + 1}"].Value = "Всего";
                sheet.Cells[$"K{index + 1}"].Value = "С грузом";
                sheet.Cells[$"L{index}"].Value = "Перевезено";
                sheet.Cells[$"L{index + 1}"].Value = "груза, тонн";
                sheet.Cells[$"M{index}"].Value = "Нормо-";
                sheet.Cells[$"M{index + 1}"].Value = "смена";
                sheet.Cells[$"N{index}"].Value = "Условный";
                sheet.Cells[$"N{index + 1}"].Value = "гектар";
                sheet.Cells[$"O{index}"].Value = "Расход ГСМ";
                sheet.Cells[$"O{index + 1}"].Value = "Факт";
                sheet.Cells[$"P{index + 1}"].Value = "Норма";
                sheet.Cells[$"Q{index + 1}"].Value = "Экон";
            }
        }

        public byte[] GenerateDetailedWaybills(List<Waybill> waybills)
        {
            var groupedWaybills = waybills.GroupBy(x => new { x.DriverId, x.TransportId }).OrderBy(x => x.First().Driver.LastName);

            var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Путевые листы детально");
            SetPrinterMargins(sheet);
            sheet.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            sheet.Cells.Style.Font.Name = "Times New Roman";
            sheet.Cells.Style.Font.Size = 8;
            sheet.Column(1).Width = GetTrueColumnWidth(9);
            sheet.Column(2).Width = GetTrueColumnWidth(4.29);
            sheet.Column(3).Width = GetTrueColumnWidth(5);
            sheet.Column(4).Width = GetTrueColumnWidth(5);
            sheet.Column(5).Width = GetTrueColumnWidth(8);
            sheet.Column(6).Width = GetTrueColumnWidth(5);
            sheet.Column(7).Width = GetTrueColumnWidth(5);
            sheet.Column(8).Width = GetTrueColumnWidth(5);
            sheet.Column(9).Width = GetTrueColumnWidth(5.29);
            sheet.Column(10).Width = GetTrueColumnWidth(5.29);
            sheet.Column(11).Width = GetTrueColumnWidth(7.29);
            sheet.Column(12).Width = GetTrueColumnWidth(7);
            sheet.Column(13).Width = GetTrueColumnWidth(7);

            sheet.DefaultRowHeight = 9;

            var index = 0;
            var maxPageHeight = 838;
            foreach (var driverWaybills in groupedWaybills)
            {
                var total = new MonthTotal(driverWaybills);

                index++;

                sheet.Row(index).Height = 12;
                sheet.Row(index + 1).Height = 12;
                sheet.Row(index + 2).Height = 9;
                sheet.Row(index + 3).Height = 9;
                sheet.Row(index + 4).Height = 9;
                sheet.Row(index + 5).Height = 12;

                sheet.Cells[$"A{index}:M{index + 5}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"A{index}:M{index + 5}"].Style.WrapText = true;
                sheet.Cells[$"A{index + 5}:M{index + 5}"].Style.Font.Bold = true;

                sheet.Cells[$"A{index}:M{index + 5}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index}:M{index + 1}, A{index + 2}:F{index + 5}, G{index + 2}, G{index + 3}, G{index + 5}, H{index + 2}:J{index + 5}," +
                    $"K{index + 2}, K{index + 5}, L{index + 2}:M{index + 5}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"M{index}:M{index + 5}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index + 5}:M{index + 5}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                sheet.Cells[$"A{index}:E{index + 1}, F{index}:G{index + 1}, H{index}:I{index + 1}, J{index}:K{index}, J{index + 1}:K{index + 1}," +
                    $"L{index}:M{index}, L{index + 1}:M{index + 1}"].Merge = true;
                sheet.Cells[$"A{index + 2}:B{index + 4}, C{index + 2}:C{index + 4}, D{index + 2}:D{index + 4}, E{index + 2}:E{index + 4}," +
                    $"F{index + 2}:G{index + 2}, F{index + 3}:F{index + 4}, H{index + 2}:I{index + 4}, J{index + 2}:J{index + 4}," +
                    $"L{index + 2}:M{index + 2}, L{index + 3}:L{index + 4}, M{index + 3}:M{index + 4}"].Merge = true;
                sheet.Cells[$"A{index + 5}:B{index + 5}, H{index + 5}:I{index + 5}"].Merge = true;

                sheet.Cells[$"A{index}"].Value = "Итого";
                sheet.Cells[$"A{index}"].Style.Font.Size = 18;
                sheet.Cells[$"A{index}"].Style.Font.Bold = true;

                sheet.Cells[$"F{index}"].Value = "Сумма заработка";
                sheet.Cells[$"H{index}"].Value = total.Earnings;
                sheet.Cells[$"H{index}"].Style.Font.Size = 12;
                sheet.Cells[$"H{index}"].Style.Font.Bold = true;

                sheet.Cells[$"J{index}"].Value = "Премия";
                sheet.Cells[$"L{index}"].Value = total.Bonus;
                sheet.Cells[$"L{index}"].Style.Font.Size = 10;
                sheet.Cells[$"L{index}"].Style.Font.Bold = true;

                sheet.Cells[$"J{index + 1}"].Value = "Выходные";
                sheet.Cells[$"L{index + 1}"].Value = total.Weekend;
                sheet.Cells[$"L{index + 1}"].Style.Font.Size = 10;
                sheet.Cells[$"L{index + 1}"].Style.Font.Bold = true;

                sheet.Cells[$"A{index + 2}"].Value = "Путевых листов";
                sheet.Cells[$"C{index + 2}"].Value = "Дней";
                sheet.Cells[$"D{index + 2}"].Value = "Часов";
                sheet.Cells[$"E{index + 2}"].Value = "Число ездок";
                sheet.Cells[$"F{index + 2}"].Value = "Пробег";
                sheet.Cells[$"F{index + 3}"].Value = "Всего";
                sheet.Cells[$"G{index + 3}"].Value = "В т.ч. с";
                sheet.Cells[$"G{index + 4}"].Value = "грузом";
                sheet.Cells[$"H{index + 2}"].Value = "Перевезено груза, тонн";
                sheet.Cells[$"J{index + 2}"].Value = "Нормо-смена";
                sheet.Cells[$"K{index + 2}"].Value = "Условный";
                sheet.Cells[$"K{index + 3}"].Value = "эталонный";
                sheet.Cells[$"K{index + 4}"].Value = "гектар";
                sheet.Cells[$"L{index + 2}"].Value = "Расход ГСМ";
                sheet.Cells[$"L{index + 3}"].Value = "По факту";
                sheet.Cells[$"M{index + 3}"].Value = "По норме";

                sheet.Cells[$"A{index + 5}"].Value = total.WaybillsCount;
                sheet.Cells[$"C{index + 5}"].Value = total.Days;
                sheet.Cells[$"D{index + 5}"].Value = total.Hours;
                sheet.Cells[$"E{index + 5}"].Value = total.NumberOfRuns;
                sheet.Cells[$"F{index + 5}"].Value = total.TotalMileage;
                sheet.Cells[$"G{index + 5}"].Value = total.TotalMileageWithLoad;
                sheet.Cells[$"H{index + 5}"].Value = total.TransportedLoad;
                sheet.Cells[$"J{index + 5}"].Value = total.NormShift;
                sheet.Cells[$"K{index + 5}"].Value = total.ConditionalReferenceHectares;
                sheet.Cells[$"L{index + 5}"].Value = total.FactFuelConsumption;
                sheet.Cells[$"M{index + 5}"].Value = total.NormalFuelConsumption;

                var pageHeight = 63;
                var pagesCount = 1;
                index += 5;
                foreach (var waybill in driverWaybills)
                {
                    var waybillHeight = 108 + waybill.Operations.Count * 9;
                    if (pageHeight + waybillHeight > maxPageHeight)
                    {
                        sheet.Row(index).PageBreak = true;
                        pagesCount++;
                        pageHeight = waybillHeight - 9;
                    }
                    else
                    {
                        pageHeight += waybillHeight;
                        index++;
                    }
                    index++;

                    sheet.Cells[$"A{index}:M{index + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    sheet.Cells[$"A{index}:M{index + 2}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{index}:M{index}"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                    sheet.Cells[$"M{index}:M{index + 2}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{index}:M{index + 2}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    sheet.Row(index).Height = 9;
                    sheet.Row(index + 1).Height = 9;
                    sheet.Row(index + 2).Height = 9;
                    sheet.Row(index + 3).Height = 4.5;
                    sheet.Row(index + 4).Height = 9;
                    sheet.Row(index + 5).Height = 9;
                    sheet.Row(index + 6).Height = 9;

                    sheet.Cells[$"A{index}:D{index}"].Merge = true;
                    sheet.Cells[$"F{index}:H{index}"].Merge = true;
                    sheet.Cells[$"I{index}:K{index}"].Merge = true;
                    sheet.Cells[$"L{index}:M{index}"].Merge = true;

                    sheet.Cells[$"A{index}"].Value = $"Путевой лист № {waybill.Number}";
                    sheet.Cells[$"E{index}"].Value = "Шифр";
                    sheet.Cells[$"F{index}"].Value = "Транспорт";
                    sheet.Cells[$"I{index}"].Value = "ГСМ";
                    sheet.Cells[$"L{index}"].Value = "Расход ГСМ";

                    index++;

                    sheet.Cells[$"A{index}:D{index}"].Merge = true;
                    sheet.Cells[$"E{index}:E{index + 1}"].Merge = true;
                    sheet.Cells[$"F{index}:H{index}"].Merge = true;
                    sheet.Cells[$"A{index}"].Value = waybill.FullDate;
                    sheet.Cells[$"E{index}"].Value = waybill.Transport.Code;
                    sheet.Cells[$"F{index}"].Value = waybill.Transport.Name;
                    sheet.Cells[$"I{index}"].Value = "Начало";
                    sheet.Cells[$"J{index}"].Value = "Выдано";
                    sheet.Cells[$"K{index}"].Value = "Конец";
                    sheet.Cells[$"L{index}"].Value = "По факту";
                    sheet.Cells[$"M{index}"].Value = "По норме";

                    sheet.Cells[$"A{index}"].Style.Font.Bold = true;
                    sheet.Cells[$"E{index}"].Style.Font.Bold = true;
                    sheet.Cells[$"F{index}"].Style.Font.Bold = true;

                    index++;

                    sheet.Cells[$"A{index}:D{index}"].Merge = true;
                    sheet.Cells[$"F{index}:G{index}"].Merge = true;
                    sheet.Cells[$"A{index}"].Value = waybill.Driver.ShortFullName();
                    sheet.Cells[$"F{index}"].Value = "Табельный №";
                    sheet.Cells[$"H{index}"].Value = waybill.Driver.PersonnelNumber;
                    sheet.Cells[$"I{index}"].Value = waybill.StartFuel;
                    sheet.Cells[$"J{index}"].Value = waybill.FuelTopUp;
                    sheet.Cells[$"K{index}"].Value = waybill.EndFuel;
                    sheet.Cells[$"L{index}"].Value = waybill.FactFuelConsumption;
                    sheet.Cells[$"M{index}"].Value = waybill.NormalFuelConsumption;

                    sheet.Cells[$"A{index}, H{index}:M{index}"].Style.Font.Bold = true;

                    index += 2;

                    sheet.Cells[$"A{index}:A{index + 2}"].Merge = true;
                    sheet.Cells[$"B{index}:B{index + 2}"].Merge = true;
                    sheet.Cells[$"C{index + 1}:C{index + 2}"].Merge = true;
                    sheet.Cells[$"C{index}:D{index}"].Merge = true;
                    sheet.Cells[$"E{index}:E{index + 2}"].Merge = true;
                    sheet.Cells[$"F{index + 1}:F{index + 2}"].Merge = true;
                    sheet.Cells[$"G{index + 1}:G{index + 2}"].Merge = true;
                    sheet.Cells[$"H{index + 1}:I{index + 1}"].Merge = true;
                    sheet.Cells[$"H{index + 2}:I{index + 2}"].Merge = true;
                    sheet.Cells[$"F{index}:I{index}"].Merge = true;
                    sheet.Cells[$"J{index}:J{index + 2}"].Merge = true;
                    sheet.Cells[$"L{index}:M{index}"].Merge = true;
                    sheet.Cells[$"L{index + 1}:L{index + 2}"].Merge = true;
                    sheet.Cells[$"M{index + 1}:M{index + 2}"].Merge = true;

                    sheet.Cells[$"A{index}:M{index + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    sheet.Cells[$"A{index}:M{index + 2}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{index}:j{index + 1}, K{index}, L{index}:M{index + 1}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"M{index}:M{index + 2}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    sheet.Cells[$"A{index}"].Value = "ШПЗ";
                    sheet.Cells[$"B{index}"].Style.WrapText = true;
                    sheet.Cells[$"B{index}"].Value = "Число ездок";
                    sheet.Cells[$"C{index}"].Value = "Пробег";
                    sheet.Cells[$"C{index + 1}"].Value = "Всего";
                    sheet.Cells[$"D{index + 1}"].Value = "В т.ч. с";
                    sheet.Cells[$"D{index + 2}"].Value = "грузом";

                    sheet.Cells[$"E{index}"].Style.WrapText = true;
                    sheet.Cells[$"E{index}"].Value = "Перевезено груза, тонн";
                    sheet.Cells[$"F{index}"].Value = "Сделано";
                    sheet.Cells[$"F{index + 1}"].Value = "Норма";
                    sheet.Cells[$"G{index + 1}"].Value = "Факт.";
                    sheet.Cells[$"H{index + 1}"].Value = "Длина ездки с";
                    sheet.Cells[$"H{index + 2}"].Value = "грузом";
                    sheet.Cells[$"J{index}"].Style.WrapText = true;
                    sheet.Cells[$"J{index}"].Value = "Нормо-смена";

                    sheet.Cells[$"K{index}"].Value = "Условный";
                    sheet.Cells[$"K{index + 1}"].Value = "эталонный";
                    sheet.Cells[$"K{index + 2}"].Value = "гектар";

                    sheet.Cells[$"L{index}"].Value = "Норма расхода ГСМ";
                    sheet.Cells[$"L{index + 1}"].Value = "За 1 ед.";
                    sheet.Cells[$"M{index + 1}"].Value = "Всего";

                    index += 3;

                    foreach (var operation in waybill.Operations)
                    {
                        sheet.Row(index).Height = 9;
                        sheet.Row(index).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[$"A{index}:M{index}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        sheet.Cells[$"A{index}:M{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        sheet.Cells[$"M{index}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        sheet.Cells[$"H{index}:I{index}"].Merge = true;
                        sheet.Cells[$"A{index}"].Value = operation.ProductionCostCode;
                        sheet.Cells[$"B{index}"].Value = operation.NumberOfRuns == 0 ? "" : operation.NumberOfRuns;
                        sheet.Cells[$"C{index}"].Value = operation.TotalMileage == 0d ? "" : operation.TotalMileage;
                        sheet.Cells[$"D{index}"].Value = operation.TotalMileageWithLoad == 0d ? "" : operation.TotalMileageWithLoad;
                        sheet.Cells[$"E{index}"].Value = operation.TransportedLoad == 0d ? "" : operation.TransportedLoad;
                        sheet.Cells[$"F{index}"].Value = operation.Norm == 0d ? "" : operation.Norm;
                        sheet.Cells[$"G{index}"].Value = operation.Fact == 0d ? "" : operation.Fact;
                        sheet.Cells[$"H{index}"].Value = operation.MileageWithLoad == 0d ? "" : operation.MileageWithLoad;
                        sheet.Cells[$"J{index}"].Value = operation.NormShift == 0d ? "" : operation.NormShift;
                        sheet.Cells[$"K{index}"].Value = operation.ConditionalReferenceHectares == 0d ? "" : operation.ConditionalReferenceHectares;
                        sheet.Cells[$"L{index}"].Value = operation.FuelConsumptionPerUnit == 0d ? "" : operation.FuelConsumptionPerUnit;
                        sheet.Cells[$"M{index}"].Value = operation.TotalFuelConsumption == 0d ? "" : operation.TotalFuelConsumption;

                        index++;
                    }

                    sheet.Cells[$"A{index}:M{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Row(index).Height = 4.5;
                    index++;
                    sheet.Row(index).Height = 9;
                    sheet.Row(index + 1).Height = 9;
                    sheet.Row(index + 2).Height = 9;
                    sheet.Row(index + 3).Height = 9;

                    sheet.Cells[$"B{index}:C{index}"].Merge = true;
                    sheet.Cells[$"B{index + 1}:C{index + 1}"].Merge = true;
                    sheet.Cells[$"B{index + 2}:C{index + 2}"].Merge = true;
                    sheet.Cells[$"B{index + 3}:C{index + 3}"].Merge = true;
                    sheet.Cells[$"F{index}:G{index}"].Merge = true;
                    sheet.Cells[$"F{index + 1}:G{index + 1}"].Merge = true;
                    sheet.Cells[$"F{index + 2}:G{index + 2}"].Merge = true;
                    sheet.Cells[$"F{index + 3}:G{index + 3}"].Merge = true;
                    sheet.Cells[$"I{index}:J{index}"].Merge = true;
                    sheet.Cells[$"L{index}:M{index}"].Merge = true;
                    sheet.Cells[$"I{index + 2}:I{index + 3}"].Merge = true;
                    sheet.Cells[$"J{index + 2}:J{index + 3}"].Merge = true;
                    sheet.Cells[$"K{index + 2}:K{index + 3}"].Merge = true;
                    sheet.Cells[$"L{index + 2}:L{index + 3}"].Merge = true;
                    sheet.Cells[$"M{index + 2}:M{index + 3}"].Merge = true;

                    sheet.Cells[$"A{index}"].Value = "Количество";
                    sheet.Cells[$"B{index}"].Value = "Расценка";
                    sheet.Cells[$"D{index}"].Value = "Сумма";
                    sheet.Cells[$"E{index}"].Value = "Количество";
                    sheet.Cells[$"F{index}"].Value = "Расценка";
                    sheet.Cells[$"H{index}"].Value = "Сумма";
                    sheet.Cells[$"I{index}"].Value = "Отработано";
                    sheet.Cells[$"I{index + 1}"].Value = "Дней";
                    sheet.Cells[$"J{index + 1}"].Value = "Часов";
                    sheet.Cells[$"K{index}"].Value = "Сумма";
                    sheet.Cells[$"K{index + 1}"].Value = "заработка";
                    sheet.Cells[$"L{index}"].Value = "Дополнительно";
                    sheet.Cells[$"L{index + 1}"].Value = "Премия";
                    sheet.Cells[$"M{index + 1}"].Value = "Выходные";

                    sheet.Cells[$"A{index}:H{index}, I{index}:M{index + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    sheet.Cells[$"A{index}:M{index + 3}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{index}:J{index + 3}, K{index}, K{index + 2}, L{index}:M{index + 2}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"M{index}:M{index + 3}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{index + 3}:M{index + 3}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                    index++;

                    for (int i = 0; i < waybill.Calculations.Count; i++)
                    {
                        if (i < 3)
                        {
                            sheet.Cells[$"A{index + i}"].Value = waybill.Calculations[i].Quantity;
                            sheet.Cells[$"B{index + i}"].Value = waybill.Calculations[i].Price;
                            sheet.Cells[$"D{index + i}"].Value = waybill.Calculations[i].Sum;
                            continue;
                        }

                        sheet.Cells[$"E{index + i - 3}"].Value = waybill.Calculations[i].Quantity;
                        sheet.Cells[$"F{index + i - 3}"].Value = waybill.Calculations[i].Price;
                        sheet.Cells[$"H{index + i - 3}"].Value = waybill.Calculations[i].Sum;
                    }

                    index++;

                    sheet.Cells[$"I{index}:M{index}"].Style.Font.Bold = true;
                    sheet.Cells[$"I{index}"].Value = waybill.Days;
                    sheet.Cells[$"J{index}"].Value = waybill.Hours;
                    sheet.Cells[$"K{index}"].Value = waybill.Earnings;
                    sheet.Cells[$"L{index}"].Value = waybill.Bonus;
                    sheet.Cells[$"M{index}"].Value = waybill.Weekend;

                    index++;
                }

                sheet.Row(index).PageBreak = true;
                if (pagesCount % 2 != 0)
                {
                    index++;
                    sheet.Row(index).PageBreak = true;
                    sheet.Cells[$"A{index}"].Value = "";
                }
            }
            return package.GetAsByteArray();
        }

        private static void SetPrinterMargins(ExcelWorksheet sheet)
        {
            sheet.PrinterSettings.LeftMargin = 0.7M;
            sheet.PrinterSettings.RightMargin = 0.5M;
            sheet.PrinterSettings.TopMargin = 0.2M;
            sheet.PrinterSettings.BottomMargin = 0.2M;
            sheet.PrinterSettings.HeaderMargin = 0M;
            sheet.PrinterSettings.FooterMargin = 0M;
        }

        private static double GetTrueColumnWidth(double width)
        {
            //DEDUCE WHAT THE COLUMN WIDTH WOULD REALLY GET SET TO
            double z;
            if (width >= (1 + 2 / 3))
            {
                z = Math.Round((Math.Round(7 * (width - 1 / 256), 0) - 5) / 7, 2);
            }
            else
            {
                z = Math.Round((Math.Round(12 * (width - 1 / 256), 0) - Math.Round(5 * width, 0)) / 12, 2);
            }

            //HOW FAR OFF? (WILL BE LESS THAN 1)
            double errorAmt = width - z;

            //CALCULATE WHAT AMOUNT TO TACK ONTO THE ORIGINAL AMOUNT TO RESULT IN THE CLOSEST POSSIBLE SETTING 
            double adj;
            if (width >= (1 + 2 / 3))
            {
                adj = (Math.Round(7 * errorAmt - 7 / 256, 0)) / 7;
            }
            else
            {
                adj = ((Math.Round(12 * errorAmt - 12 / 256, 0)) / 12) + (2 / 12);
            }

            //RETURN A SCALED-VALUE THAT SHOULD RESULT IN THE NEAREST POSSIBLE VALUE TO THE TRUE DESIRED SETTING
            if (z > 0)
            {
                return width + adj;
            }

            return 0d;
        }
    }
}
