using WaybillsAPI.Models;

namespace WaybillsAPI.ReportsModels
{
    public class EntityMonthTotal(string name, int code, List<Waybill> waybills) : MonthTotal(waybills)
    {
        public string EntityName { get; private set; } = name;
        public int EntityCode { get; private set; } = code;
    }
}
