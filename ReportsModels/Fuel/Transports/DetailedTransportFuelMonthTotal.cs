using WaybillsAPI.Models;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.ReportsModels.Fuel.Transports
{
    public class DetailedTransportFuelMonthTotal : TransportFuelMonthTotal
    {
        public TransportDTO Transport { get; private set; }
        public IEnumerable<TransportFuelSubTotal> SubTotals { get; private set; }

        public DetailedTransportFuelMonthTotal(TransportDTO transport, Dictionary<int, DriverDTO> driversDTO, IEnumerable<Waybill> waybills) : base(waybills)
        {
            Transport = transport;

            var currentPeriodWaybills = new List<Waybill>();
            var previousPeriodWaybills = new List<Waybill>();
            foreach (var waybill in waybills.OrderBy(x => x.Date))
            {
                if (waybill.Date.Year == waybill.SalaryYear && waybill.Date.Month == waybill.SalaryMonth)
                {
                    currentPeriodWaybills.Add(waybill);
                }
                else
                {
                    previousPeriodWaybills.Add(waybill);
                }
            }

            var allPeriodWaybills = new List<List<Waybill>>();
            if (currentPeriodWaybills.Count != 0)
            {
                allPeriodWaybills.Add(currentPeriodWaybills);
            }

            if (previousPeriodWaybills.Count != 0)
            {
                allPeriodWaybills.Add(previousPeriodWaybills);
            }

            var subTotals = new List<TransportFuelSubTotal>();

            foreach (var periodWaybilss in allPeriodWaybills)
            {
                var currentDriverWaybills = new List<Waybill>() { periodWaybilss[0] };
                for (int i = 1; i < periodWaybilss.Count; i++)
                {
                    if (periodWaybilss[i].DriverId == periodWaybilss[i - 1].DriverId)
                    {
                        currentDriverWaybills.Add(periodWaybilss[i]);
                    }
                    else
                    {
                        subTotals.Add(new TransportFuelSubTotal(driversDTO[periodWaybilss[i - 1].DriverId], currentDriverWaybills));
                        currentDriverWaybills.Clear();
                        currentDriverWaybills.Add(periodWaybilss[i]);
                    }
                }
                subTotals.Add(new TransportFuelSubTotal(driversDTO[periodWaybilss.Last().DriverId], currentDriverWaybills));
            }
            SubTotals = subTotals;
        }
    }
}
