namespace WaybillsAPI.ViewModels
{
    public class CostCodeInfo(string productionCostCode, double conditionalReferenceHectares)
    {
        public string ProductionCostCode { get; set; } = productionCostCode;
        public double ConditionalReferenceHectares { get; set; } = conditionalReferenceHectares;
        public double CostPrice { get; set; }
    }
}
