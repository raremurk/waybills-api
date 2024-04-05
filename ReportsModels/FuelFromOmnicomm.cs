namespace WaybillsAPI.ReportsModels
{
    public class FuelFromOmnicomm
    {
        public string TransportName { get; set; }
        public double StartFuel { get; set; }
        public double FuelTopUp { get; set; }
        public double EndFuel { get; set; }
        public double FuelConsumption { get; set; }
        public double Draining { get; set; }
    }
}
