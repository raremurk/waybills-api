using WaybillsAPI.Models;

namespace WaybillsAPI.ReportsModels.CostPrice
{
    public class CostPriceReport
    {
        public double ConditionalReferenceHectares { get; private set; }
        public double CostPrice { get; private set; }
        public IEnumerable<CostCodeInfo> CostCodes { get; private set; }

        public CostPriceReport(IEnumerable<Waybill> waybills, double price)
        {
            var costPrices = new Dictionary<string, CostCodeInfoCreation>();
            foreach (var waybill in waybills)
            {
                ConditionalReferenceHectares += waybill.ConditionalReferenceHectares;
                foreach (var operation in waybill.Operations)
                {
                    if (operation.Norm == 0)
                    {
                        continue;
                    }
                    if (costPrices.TryGetValue(operation.ProductionCostCode, out CostCodeInfoCreation? value))
                    {
                        value.ConditionalReferenceHectares += operation.ConditionalReferenceHectares;
                        value.WaybillIdentifiers.TryAdd(waybill.Id, waybill.Number);
                        continue;
                    }
                    costPrices.Add(operation.ProductionCostCode, new(operation.ConditionalReferenceHectares, waybill.Id, waybill.Number));
                }
            }
            CostPrice = Math.Round(ConditionalReferenceHectares * price, 2);
            CostCodes = costPrices.Select(x => new CostCodeInfo(x.Key, x.Value, price)).OrderBy(c => c.ProductionCostCode);
        }
    }
}
