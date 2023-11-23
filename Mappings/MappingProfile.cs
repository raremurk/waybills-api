using AutoMapper;
using WaybillsAPI.CreationModels;
using WaybillsAPI.Models;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.Mappings
{
    public class MappingProfile : Profile
    {
        public MappingProfile()
        {
            CreateMap<Driver, DriverDTO>().ReverseMap();
            CreateMap<Transport, TransportDTO>().ReverseMap();
            CreateMap<Waybill, WaybillDTO>().ReverseMap();
            CreateMap<Waybill, WaybillCreation>();
            CreateMap<Operation, OperationCreation>();
            CreateMap<Calculation, CalculationCreation>();
            CreateMap<Waybill, ShortWaybillDTO>()
                .ForMember("Date", opt => opt.MapFrom(x => x.FullDate));
            CreateMap<Operation, OperationDTO>().ReverseMap();
            CreateMap<Calculation, CalculationDTO>().ReverseMap();
        }
    }
}
