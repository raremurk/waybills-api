namespace WaybillsAPI.ReportsModels.Fuel.Transports
{
    public abstract class TransportFuelMonthTotal
    {
        public int StartFuel { get; protected set; }
        public int FuelTopUp { get; protected set; }
        public int FactFuelConsumption { get; protected set; }
        public int EndFuel { get; protected set; }
        public int Deviation { get; protected set; }
    }
}