namespace WaybillsAPI.Omnicomm.OmnicommModels
{
    public class OmnicommData
    {
        public IEnumerable<OmnicommVehicleData> VehicleDataList { get; set; }
        public OmnicommTotalFuel TotalFuel { get; set; }
    }
}
