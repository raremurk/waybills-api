using WaybillsAPI.ViewModels;

namespace WaybillsAPI.ReportsModels.Fuel.Drivers
{
    public class DetailedDriverFuelMonthTotal(DriverDTO driver, IEnumerable<DriverFuelSubTotal> subTotals) : DriverFuelMonthTotal(subTotals)
    {
        public DriverDTO Driver { get; private set; } = driver;
        public IEnumerable<DriverFuelSubTotal> SubTotals { get; private set; } = subTotals;
    }
}
