using WaybillsAPI.Models;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.ReportsModels.Fuel.Drivers
{
    public class DriverFuelSubTotal(TransportDTO transport, IEnumerable<Waybill> waybills) : DriverFuelMonthTotal(waybills)
    {
        public TransportDTO Transport { get; private set; } = transport;
    }
}
