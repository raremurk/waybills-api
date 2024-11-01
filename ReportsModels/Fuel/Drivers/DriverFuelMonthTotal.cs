namespace WaybillsAPI.ReportsModels.Fuel.Drivers
{
    public abstract class DriverFuelMonthTotal
    {
        public int FuelTopUp { get; protected set; }
        public int FactFuelConsumption { get; protected set; }
        public int NormalFuelConsumption { get; protected set; }
        public int FuelEconomy { get; protected set; }
    }
}
