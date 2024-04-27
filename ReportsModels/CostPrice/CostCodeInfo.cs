namespace WaybillsAPI.ReportsModels.CostPrice
{
    public class CostCodeInfo
    {
        public string ProductionCostCode { get; private set; }
        public double ConditionalReferenceHectares { get; private set; }
        public double CostPrice { get; private set; }
        public IEnumerable<WaybillIdentifier> WaybillIdentifiers { get; private set; }

        public CostCodeInfo(string productionCostCode, CostCodeInfoCreation costCodeInfoCreation, double price)
        {
            ProductionCostCode = productionCostCode;
            ConditionalReferenceHectares = Math.Round(costCodeInfoCreation.ConditionalReferenceHectares, 2);
            CostPrice = Math.Round(ConditionalReferenceHectares * price, 2);
            WaybillIdentifiers = costCodeInfoCreation.WaybillIdentifiers.Select(x => new WaybillIdentifier(x.Key, x.Value)).OrderBy(x => x.Number);
        }
    }

    public class CostCodeInfoCreation(double conditionalReferenceHectares, int waybillId, int waybillNumber)
    {
        public double ConditionalReferenceHectares { get; set; } = conditionalReferenceHectares;
        public Dictionary<int, int> WaybillIdentifiers { get; set; } = new() { { waybillId, waybillNumber } };
    }

    public class WaybillIdentifier(int id, int number)
    {
        public int Id { get; set; } = id;
        public int Number { get; set; } = number;
    }
}
