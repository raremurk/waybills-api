using WaybillsAPI.Models;

namespace WaybillsAPI.ReportsModels
{
    public class TransportMonthTotal(Transport transport, List<Waybill> waybills) : MonthTotal(waybills)
    {
        public string TransportName { get; private set; } = transport.Name;
        public int TransportCode { get; private set; } = transport.Code;
    }
}
