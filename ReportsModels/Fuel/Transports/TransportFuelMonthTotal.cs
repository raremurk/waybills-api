using WaybillsAPI.Models;

namespace WaybillsAPI.ReportsModels.Fuel.Transports
{
    public class TransportFuelMonthTotal
    {
        public int StartFuel { get; private set; }
        public int FuelTopUp { get; private set; }
        public int FactFuelConsumption { get; private set; }
        public int EndFuel { get; private set; }
        public int Deviation { get; private set; }

        public TransportFuelMonthTotal(IEnumerable<Waybill> waybills)
        {
            var currentPeriodWaybills = waybills.Where(x => x.Date.Year == x.SalaryYear && x.Date.Month == x.SalaryMonth);
            var waybillsForWork = currentPeriodWaybills.Any() ? currentPeriodWaybills.OrderBy(x => x.Date) : waybills.OrderBy(x => x.Date);

            StartFuel = waybillsForWork.First().StartFuel;
            FuelTopUp = waybills.Sum(x => x.FuelTopUp);
            FactFuelConsumption = waybills.Sum(x => x.FactFuelConsumption);
            EndFuel = waybillsForWork.Last().EndFuel;
            Deviation = Math.Abs(StartFuel + FuelTopUp - FactFuelConsumption - EndFuel);
        }
    }
}