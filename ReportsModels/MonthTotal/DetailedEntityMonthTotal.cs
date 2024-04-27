using WaybillsAPI.Models;

namespace WaybillsAPI.ReportsModels.MonthTotal
{
    public class DetailedEntityMonthTotal : MonthTotal
    {
        public string EntityName { get; private set; }
        public int EntityCode { get; private set; }
        public List<EntityMonthTotal> SubTotals { get; private set; } = [];

        public DetailedEntityMonthTotal(IEnumerable<Waybill> waybills, Driver driver)
        {
            ArgumentNullException.ThrowIfNull(driver);
            if (waybills.Any(x => x.DriverId != driver.Id)) throw new ArgumentException("Waybills must belong to one driver.");

            EntityName = driver.ShortFullName();
            EntityCode = driver.PersonnelNumber;
            foreach (var subGroup in waybills.GroupBy(x => x.TransportId))
            {
                var transport = subGroup.First().Transport ?? throw new ArgumentException("Waybills must include transport.");
                SubTotals.Add(new EntityMonthTotal(transport.Name, transport.Code, subGroup));
            }
            Initialize(SubTotals);
        }

        public DetailedEntityMonthTotal(IEnumerable<Waybill> waybills, Transport transport)
        {
            ArgumentNullException.ThrowIfNull(transport);
            if (waybills.Any(x => x.TransportId != transport.Id)) throw new ArgumentException("Waybills must belong to one transport.");

            EntityName = transport.Name;
            EntityCode = transport.Code;
            foreach (var subGroup in waybills.GroupBy(x => x.DriverId))
            {
                var driver = subGroup.First().Driver ?? throw new ArgumentException("Waybills must include driver.");
                SubTotals.Add(new EntityMonthTotal(driver.ShortFullName(), driver.PersonnelNumber, subGroup));
            }
            Initialize(SubTotals);
        }
    }
}
