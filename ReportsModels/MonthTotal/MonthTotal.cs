using WaybillsAPI.Models;

namespace WaybillsAPI.ReportsModels.MonthTotal
{
    public class MonthTotal
    {
        public int WaybillsCount { get; private set; }
        public int Days { get; private set; }
        public double Hours { get; private set; }

        public double Earnings { get; private set; }
        public double Weekend { get; private set; }
        public double Bonus { get; private set; }

        public int FactFuelConsumption { get; private set; }
        public int NormalFuelConsumption { get; private set; }

        public int NumberOfRuns { get; private set; }
        public double TotalMileage { get; private set; }
        public double TotalMileageWithLoad { get; private set; }
        public double TransportedLoad { get; private set; }
        public double NormShift { get; private set; }
        public double ConditionalReferenceHectares { get; private set; }

        public MonthTotal(IEnumerable<Waybill> waybills)
        {
            foreach (var waybill in waybills)
            {
                WaybillsCount++;
                Days += waybill.Days;
                Hours += waybill.Hours;
                Earnings += waybill.Earnings;
                Weekend += waybill.Weekend;
                Bonus += waybill.Bonus;
                FactFuelConsumption += waybill.FactFuelConsumption;
                NormalFuelConsumption += waybill.NormalFuelConsumption;

                foreach (var operation in waybill.Operations)
                {
                    NumberOfRuns += operation.NumberOfRuns;
                    TotalMileage += operation.TotalMileage;
                    TotalMileageWithLoad += operation.TotalMileageWithLoad;
                    TransportedLoad += operation.TransportedLoad;
                    NormShift += operation.NormShift;
                    ConditionalReferenceHectares += operation.ConditionalReferenceHectares;
                }
            }
            RoundValues();
        }

        protected MonthTotal() { }
        protected void Initialize(IEnumerable<MonthTotal> totals)
        {
            foreach (var total in totals)
            {
                WaybillsCount += total.WaybillsCount;
                Days += total.Days;
                Hours += total.Hours;
                Earnings += total.Earnings;
                Weekend += total.Weekend;
                Bonus += total.Bonus;
                FactFuelConsumption += total.FactFuelConsumption;
                NormalFuelConsumption += total.NormalFuelConsumption;
                NumberOfRuns += total.NumberOfRuns;
                TotalMileage += total.TotalMileage;
                TotalMileageWithLoad += total.TotalMileageWithLoad;
                TransportedLoad += total.TransportedLoad;
                NormShift += total.NormShift;
                ConditionalReferenceHectares += total.ConditionalReferenceHectares;
            }
            RoundValues();
        }

        private void RoundValues()
        {
            Hours = Math.Round(Hours, 1);
            Earnings = Math.Round(Earnings, 2);
            Weekend = Math.Round(Weekend, 2);
            Bonus = Math.Round(Bonus, 2);
            TotalMileage = Math.Round(TotalMileage, 1);
            TotalMileageWithLoad = Math.Round(TotalMileageWithLoad, 1);
            TransportedLoad = Math.Round(TransportedLoad, 3);
            NormShift = Math.Round(NormShift, 2);
            ConditionalReferenceHectares = Math.Round(ConditionalReferenceHectares, 2);
        }
    }
}