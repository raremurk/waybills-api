using WaybillsAPI.Models;

namespace WaybillsAPI.ReportsModels
{
    public class DriverMonthTotal(Driver driver, List<TransportMonthTotal> transportTotals) : MonthTotal(transportTotals.Select(x => (MonthTotal)x).ToList())
    {
        public string DriverFullName { get; private set; } = driver.ShortFullName();
        public int DriverPersonnelNumber { get; private set; } = driver.PersonnelNumber;
        public List<TransportMonthTotal> TransportTotals { get; private set; } = transportTotals;
    }
}
