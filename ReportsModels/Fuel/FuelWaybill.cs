using WaybillsAPI.ViewModels;

namespace WaybillsAPI.ReportsModels.Fuel
{
    public class FuelWaybill
    {
        public int Id { get; set; }

        public DateOnly Date { get; set; }
        public int Days { get; private set; }

        public int StartFuel { get; set; }
        public int FuelTopUp { get; set; }
        public int EndFuel { get; set; }
        public int FactFuelConsumption { get; set; }
        public int NormalFuelConsumption { get; set; }
        public int FuelEconomy { get; set; }

        public DriverDTO? Driver { get; set; }
        public TransportDTO? Transport { get; set; }
    }
}
