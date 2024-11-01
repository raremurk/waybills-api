using WaybillsAPI.Models;

namespace WaybillsAPI.ViewModels
{
    public class TransportDTO
    {
        public int Id { get; init; }
        public string Name { get; init; } = "";
        public int Code { get; init; }
        public double Coefficient { get; init; }
        public int OmnicommId { get; init; }

        private TransportDTO() { }
        public TransportDTO(Transport transport)
        {
            ArgumentNullException.ThrowIfNull(transport);
            Id = transport.Id;
            Name = transport.Name;
            Code = transport.Code;
            Coefficient = transport.Coefficient;
            OmnicommId = transport.OmnicommId;
        }
    }
}
