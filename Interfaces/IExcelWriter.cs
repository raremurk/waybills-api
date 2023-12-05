using WaybillsAPI.Models;

namespace WaybillsAPI.Interfaces
{
    public interface IExcelWriter
    {
        public byte[] Generate(List<Waybill> waybills);
    }
}
