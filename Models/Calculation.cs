using WaybillsAPI.CreationModels;

namespace WaybillsAPI.Models
{
    public class Calculation
    {
        public int Id { get; init; }
        public double Quantity { get; init; }
        public double Price { get; init; }
        public double Sum { get; init; }

        public int WaybillId { get; init; }
        public Waybill? Waybill { get; init; }

        private Calculation() { }
        public Calculation(CalculationCreation calculationCreation)
        {
            Id = calculationCreation.Id;
            Quantity = Math.Round(calculationCreation.Quantity, 3);
            Price = Math.Round(calculationCreation.Price, 3);
            Sum = Math.Round(Quantity * Price, 2);
        }
    }
}
