using WaybillsAPI.Models;

namespace WaybillsAPI.ReportsModels.MonthTotal
{
    public class EntityMonthTotal(string entityName, int entityCode, IEnumerable<Waybill> waybills) : MonthTotal(waybills)
    {
        public string EntityName { get; private set; } = entityName;
        public int EntityCode { get; private set; } = entityCode;
    }
}
