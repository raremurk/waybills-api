using WaybillsAPI.Models;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.ReportsModels.Fuel.Transports
{
    public class TransportFuelSubTotal(DriverDTO driver, IEnumerable<Waybill> waybills) : TransportFuelMonthTotal(waybills)
    {
        public DriverDTO Driver { get; private set; } = driver;
    }
}
