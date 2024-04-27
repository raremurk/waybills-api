using WaybillsAPI.Models;
using WaybillsAPI.ReportsModels.CostPrice;

namespace WaybillsAPI.Interfaces
{
    public interface IExcelWriter
    {
        public byte[] GenerateCostPriceReport(CostPriceReport report, int year, int month);

        public byte[] GenerateMonthTotal(List<Waybill> waybills);

        public byte[] GenerateShortWaybills(List<Waybill> waybills);

        public byte[] GenerateDetailedWaybills(List<Waybill> waybills);
    }
}
