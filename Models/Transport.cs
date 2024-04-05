namespace WaybillsAPI.Models
{
    public class Transport
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Code { get; set; }
        public double Coefficient { get; set; }
        public int OmnicommId { get; set; }

        public List<Waybill> Waybills { get; set; }
    }
}
