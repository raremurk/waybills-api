using WaybillsAPI.Models;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.ReportsModels.Fuel.Transports
{
    public class DetailedTransportFuelMonthTotal : TransportFuelMonthTotal
    {
        public TransportDTO Transport { get; init; }
        public List<TransportFuelSubTotal> SubTotals { get; init; } = [];

        public DetailedTransportFuelMonthTotal(IEnumerable<Waybill> waybills)
        {
            if (waybills is null || !waybills.Any()) throw new ArgumentException("The collection of waybills is null or empty.");

            var transport = waybills.First().Transport ?? throw new ArgumentException("Waybills must include transport.");

            if (waybills.Any(x => x.TransportId != transport.Id)) throw new ArgumentException("Waybills must belong to one transport.");

            if (waybills.Any(x => x.Driver is null)) throw new ArgumentException("Waybills must include driver.");

            Transport = new TransportDTO(transport);
            var waybillsGroupedByMonths = waybills.OrderBy(x => x.Date).GroupBy(x => x.Date.ToString("MMyyyy"));
            foreach (var group in waybillsGroupedByMonths)
            {
                var workStreak = new List<Waybill> { group.First() };
                foreach (var waybill in group.Skip(1))
                {
                    if (waybill.DriverId != workStreak.Last().Driver!.Id)
                    {
                        SubTotals.Add(new SubTotal(new DriverDTO(workStreak.Last().Driver!), workStreak));
                        workStreak.Clear();
                    }
                    workStreak.Add(waybill);
                }
                SubTotals.Add(new SubTotal(new DriverDTO(workStreak.Last().Driver!), workStreak));
            }
            StartFuel = waybillsGroupedByMonths.Last().First().StartFuel;
            InitializeFromSubTotals();
        }

        private void InitializeFromSubTotals()
        {
            FuelTopUp = SubTotals.Sum(x => x.FuelTopUp);
            FactFuelConsumption = SubTotals.Sum(x => x.FactFuelConsumption);
            EndFuel = SubTotals.Last().EndFuel;
            Deviation = Math.Abs(StartFuel + FuelTopUp - FactFuelConsumption - EndFuel);
        }


        private class SubTotal : TransportFuelSubTotal
        {
            public SubTotal(DriverDTO driver, IEnumerable<Waybill> waybills) : base(driver)
            {
                StartFuel = waybills.First().StartFuel;
                FuelTopUp = waybills.Sum(x => x.FuelTopUp);
                FactFuelConsumption = waybills.Sum(x => x.FactFuelConsumption);
                EndFuel = waybills.Last().EndFuel;
                Deviation = Math.Abs(StartFuel + FuelTopUp - FactFuelConsumption - EndFuel);
            }
        }
    }
}
