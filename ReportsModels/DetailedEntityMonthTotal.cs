namespace WaybillsAPI.ReportsModels
{
    public class DetailedEntityMonthTotal(string name, int code, IEnumerable<EntityMonthTotal> subTotals) : MonthTotal(subTotals)
    {
        public string EntityName { get; private set; } = name;
        public int EntityCode { get; private set; } = code;
        public IEnumerable<EntityMonthTotal> SubTotals { get; private set; } = subTotals;
    }
}
