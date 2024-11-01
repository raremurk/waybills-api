using WaybillsAPI.ViewModels;

namespace WaybillsAPI.ReportsModels.Fuel.Drivers
{
    public abstract class DriverFuelSubTotal(TransportDTO transport) : DriverFuelMonthTotal
    {
        public TransportDTO Transport { get; protected set; } = transport ?? throw new ArgumentNullException(nameof(transport));
    }
}
