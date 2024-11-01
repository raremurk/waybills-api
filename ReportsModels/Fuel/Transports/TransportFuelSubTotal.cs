using WaybillsAPI.ViewModels;

namespace WaybillsAPI.ReportsModels.Fuel.Transports
{
    public abstract class TransportFuelSubTotal(DriverDTO driver) : TransportFuelMonthTotal
    {
        public DriverDTO Driver { get; protected set; } = driver ?? throw new ArgumentNullException(nameof(driver));
    }
}
