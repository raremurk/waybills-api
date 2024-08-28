using WaybillsAPI.Omnicomm.OmnicommModels;

namespace WaybillsAPI.ReportsModels
{
    public class FuelFromOmnicomm(OmnicommVehicleData vehicleData)
    {
        public string TransportName { get; private set; } = vehicleData.Name;
        public double StartFuel { get; private set; } = vehicleData.Fuel.StartVolume / 10 ?? 0d;
        public double FuelTopUp { get; private set; } = vehicleData.Fuel.Refuelling / 10 ?? 0d;
        public double EndFuel { get; private set; } = vehicleData.Fuel.EndVolume / 10 ?? 0d;
        public double FuelConsumption { get; private set; } = vehicleData.Fuel.FuelConsumption / 10 ?? 0d;
        public double Draining { get; private set; } = vehicleData.Fuel.Draining / 10 ?? 0d;
    }
}
