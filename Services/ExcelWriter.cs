using OfficeOpenXml;
using OfficeOpenXml.Style;
using WaybillsAPI.Interfaces;
using WaybillsAPI.Models;

namespace WaybillsAPI.Services
{
    public class ExcelWriter : IExcelWriter
    {
        public byte[] Generate(List<Waybill> waybills)
        {
            var total = GetWaybillsTotal(waybills);

            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets.Add("Путевые листы");
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

            sheet.Row(1).Height = 12;
            sheet.Row(2).Height = 12;
            sheet.Row(3).Height = 9;
            sheet.Row(4).Height = 21.75;
            sheet.Row(5).Height = 12;

            sheet.Cells["A1:M5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Cells["A1:M5"].Style.WrapText = true;
            sheet.Cells["A5:M5"].Style.Font.Bold = true;

            sheet.Cells["A1:M5"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            sheet.Cells["A1:M5"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            sheet.Cells["A1:M5"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            sheet.Cells["A1:M5"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            sheet.Cells["A1:E2, F1:G2, H1:I2, J1:K1, J2:K2, L1:M1, L2:M2"].Merge = true;
            sheet.Cells["A3:B4, C3:C4, D3:D4, E3:E4, F3:G3, H3:I4, J3:J4, K3:K4, L3:M3"].Merge = true;
            sheet.Cells["A5:B5, H5:I5"].Merge = true;

            sheet.Cells["A1"].Value = "Итого";
            sheet.Cells["A1"].Style.Font.Size = 18;
            sheet.Cells["A1"].Style.Font.Bold = true;

            sheet.Cells["F1"].Value = "Сумма заработка";
            sheet.Cells["H1"].Value = total.Earnings;
            sheet.Cells["H1"].Style.Font.Size = 12;
            sheet.Cells["H1"].Style.Font.Bold = true;

            sheet.Cells["J1"].Value = "Выходные";
            sheet.Cells["L1"].Value = total.Weekend;
            sheet.Cells["L1"].Style.Font.Size = 10;
            sheet.Cells["L1"].Style.Font.Bold = true;

            sheet.Cells["J2"].Value = "Премия";
            sheet.Cells["L2"].Value = total.Bonus;
            sheet.Cells["L2"].Style.Font.Size = 10;
            sheet.Cells["L2"].Style.Font.Bold = true;

            sheet.Cells["A3"].Value = "Путевых листов";
            sheet.Cells["C3"].Value = "Дней";
            sheet.Cells["D3"].Value = "Часов";
            sheet.Cells["E3"].Value = "Число ездок";
            sheet.Cells["F3"].Value = "Пробег";
            sheet.Cells["F4"].Value = "Всего";
            sheet.Cells["G4"].Value = "В т.ч. с грузом";
            sheet.Cells["H3"].Value = "Перевезено груза, тонн";
            sheet.Cells["J3"].Value = "Нормо-смена";
            sheet.Cells["K3"].Value = "Условный эталонный гектар";
            sheet.Cells["L3"].Value = "Расход ГСМ";
            sheet.Cells["L4"].Value = "По факту";
            sheet.Cells["M4"].Value = "По норме";

            sheet.Cells["A5"].Value = total.WaybillsCount;
            sheet.Cells["C5"].Value = total.Days;
            sheet.Cells["D5"].Value = total.Hours;
            sheet.Cells["E5"].Value = total.NumberOfRuns;
            sheet.Cells["F5"].Value = total.TotalMileage;
            sheet.Cells["G5"].Value = total.TotalMileageWithLoad;
            sheet.Cells["H5"].Value = total.TransportedLoad;
            sheet.Cells["J5"].Value = total.NormShift;
            sheet.Cells["K5"].Value = total.ConditionalReferenceHectares;
            sheet.Cells["L5"].Value = total.FactFuelConsumption;
            sheet.Cells["M5"].Value = total.NormalFuelConsumption;

            var index = 5;
            var pageHeight = 66.75;
            foreach (var waybill in waybills)
            {
                var waybillHeight = 113.25 + waybill.Operations.Count * 9;
                if (pageHeight + waybillHeight > 756)
                {
                    sheet.Row(index).PageBreak = true;
                    pageHeight = waybillHeight;
                }
                else
                {
                    pageHeight += waybillHeight;
                    index++;
                }
                index++;

                sheet.Cells[$"A{index}:M{index + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                sheet.Cells[$"A{index}:M{index}"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                sheet.Cells[$"A{index}:M{index + 2}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index}:M{index + 2}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"M{index}:M{index + 2}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                sheet.Row(index).Height = 9;
                sheet.Row(index + 1).Height = 9;
                sheet.Row(index + 2).Height = 9;
                sheet.Row(index + 3).Height = 6;
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

                sheet.Cells[$"A{index}"].Style.Font.Bold = true;
                sheet.Cells[$"H{index}"].Style.Font.Bold = true;
                sheet.Cells[$"I{index}"].Style.Font.Bold = true;
                sheet.Cells[$"J{index}"].Style.Font.Bold = true;
                sheet.Cells[$"K{index}"].Style.Font.Bold = true;
                sheet.Cells[$"L{index}"].Style.Font.Bold = true;
                sheet.Cells[$"M{index}"].Style.Font.Bold = true;

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

                sheet.Cells[$"A{index}:M{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index}:M{index + 2}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"M{index}:M{index + 2}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"C{index}:D{index}, F{index}:I{index}, L{index}:M{index}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

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
                    sheet.Cells[$"A{index}:M{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{index}:M{index}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{index}:M{index}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"A{index}:M{index}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[$"H{index}:I{index}"].Merge = true;
                    sheet.Cells[$"A{index}"].Value = operation.ProductionCostCode;
                    sheet.Cells[$"B{index}"].Value = operation.NumberOfRuns;
                    sheet.Cells[$"C{index}"].Value = operation.TotalMileage;
                    sheet.Cells[$"D{index}"].Value = operation.TotalMileageWithLoad;
                    sheet.Cells[$"E{index}"].Value = operation.TransportedLoad;
                    sheet.Cells[$"F{index}"].Value = operation.Norm;
                    sheet.Cells[$"G{index}"].Value = operation.Fact;
                    sheet.Cells[$"H{index}"].Value = operation.MileageWithLoad;
                    sheet.Cells[$"J{index}"].Value = operation.NormShift;
                    sheet.Cells[$"K{index}"].Value = operation.ConditionalReferenceHectares;
                    sheet.Cells[$"L{index}"].Value = operation.FuelConsumptionPerUnit;
                    sheet.Cells[$"M{index}"].Value = operation.TotalFuelConsumption;

                    index++;
                }

                sheet.Row(index).Height = 6;
                index++;
                sheet.Row(index).Height = 9;
                sheet.Row(index + 1).Height = 9;
                sheet.Row(index + 2).Height = 9;
                sheet.Row(index + 3).Height = 9;

                sheet.Cells[$"A{index + 1}:H{index + 3}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

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
                sheet.Cells[$"L{index + 1}"].Value = "Выходные";
                sheet.Cells[$"M{index + 1}"].Value = "Премия";

                sheet.Cells[$"A{index}:M{index}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"I{index + 1}:M{index + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"A{index}:J{index + 3}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index}:J{index + 3}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index}:J{index + 3}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"A{index}:J{index + 3}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                sheet.Cells[$"K{index}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"K{index + 2}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"K{index + 3}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                sheet.Cells[$"L{index}:M{index + 3}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"L{index}:M{index + 3}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"L{index}:M{index + 3}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[$"L{index}:M{index + 3}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                for (int i = 0; i < waybill.Calculations.Count; i++)
                {
                    if (i < 3)
                    {
                        sheet.Cells[$"A{index + i + 1}"].Value = waybill.Calculations[i].Quantity;
                        sheet.Cells[$"B{index + i + 1}"].Value = waybill.Calculations[i].Price;
                        sheet.Cells[$"D{index + i + 1}"].Value = waybill.Calculations[i].Sum;
                        continue;
                    }

                    sheet.Cells[$"E{index + i - 2}"].Value = waybill.Calculations[i].Quantity;
                    sheet.Cells[$"F{index + i - 2}"].Value = waybill.Calculations[i].Price;
                    sheet.Cells[$"H{index + i - 2}"].Value = waybill.Calculations[i].Sum;
                }

                sheet.Cells[$"I{index + 2}"].Style.Font.Bold = true;
                sheet.Cells[$"J{index + 2}"].Style.Font.Bold = true;
                sheet.Cells[$"K{index + 2}"].Style.Font.Bold = true;
                sheet.Cells[$"L{index + 2}"].Style.Font.Bold = true;
                sheet.Cells[$"M{index + 2}"].Style.Font.Bold = true;

                sheet.Cells[$"I{index + 2}"].Value = waybill.Days;
                sheet.Cells[$"J{index + 2}"].Value = waybill.Hours;
                sheet.Cells[$"K{index + 2}"].Value = waybill.Earnings;
                sheet.Cells[$"L{index + 2}"].Value = waybill.Weekend;
                sheet.Cells[$"M{index + 2}"].Value = waybill.Bonus;

                index += 3;
            }
            return package.GetAsByteArray();
        }

        private WaybillsTotal GetWaybillsTotal(List<Waybill> waybills)
        {
            var total = new WaybillsTotal();
            foreach (var waybill in waybills)
            {
                total.Earnings += waybill.Earnings;
                total.Weekend += waybill.Weekend;
                total.Bonus += waybill.Bonus;
                total.WaybillsCount++;
                total.Days += waybill.Days;
                total.Hours += waybill.Hours;
                total.ConditionalReferenceHectares += waybill.ConditionalReferenceHectares;
                total.FactFuelConsumption += waybill.FactFuelConsumption;
                total.NormalFuelConsumption += waybill.NormalFuelConsumption;

                foreach (var operation in waybill.Operations)
                {
                    total.NumberOfRuns += operation.NumberOfRuns;
                    total.TotalMileage += operation.TotalMileage;
                    total.TotalMileageWithLoad += operation.TotalMileageWithLoad;
                    total.TransportedLoad += operation.TransportedLoad;
                    total.NormShift += operation.NormShift;
                }
            }

            return total;
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

        private class WaybillsTotal
        {
            public double Earnings { get; set; }
            public double Weekend { get; set; }
            public double Bonus { get; set; }
            public int WaybillsCount { get; set; }
            public int Days { get; set; }
            public double Hours { get; set; }
            public int NumberOfRuns { get; set; }
            public double TotalMileage { get; set; }
            public double TotalMileageWithLoad { get; set; }
            public double TransportedLoad { get; set; }
            public double NormShift { get; set; }
            public double ConditionalReferenceHectares { get; set; }
            public int FactFuelConsumption { get; set; }
            public int NormalFuelConsumption { get; set; }
        }
    }
}
