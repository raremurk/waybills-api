using WaybillsAPI.Omnicomm.OmnicommModels;

namespace WaybillsAPI.Interfaces
{
    public interface IOmnicommFuelService
    {
        public Task<OmnicommFuelReport> GetOmnicommFuelReport(DateOnly date, int omnicommId);
    }
}
