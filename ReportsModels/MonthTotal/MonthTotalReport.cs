using WaybillsAPI.Models;

namespace WaybillsAPI.ReportsModels.MonthTotal
{
    public class MonthTotalReport : MonthTotal
    {
        public IEnumerable<DetailedEntityMonthTotal> DetailedEntityMonthTotals { get; private set; }

        public MonthTotalReport(IEnumerable<Waybill> waybills, bool byDrivers = true)
        {
            DetailedEntityMonthTotals = byDrivers ?
                waybills.GroupBy(x => x.DriverId).Select(x => new DetailedEntityMonthTotal(x, x.First().Driver!)).OrderBy(x => x.EntityName) :
                waybills.GroupBy(x => x.TransportId).Select(x => new DetailedEntityMonthTotal(x, x.First().Transport!)).OrderBy(x => x.EntityName);

            Initialize(DetailedEntityMonthTotals);
        }
    }
}