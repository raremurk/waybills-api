using AutoMapper;
using WaybillsAPI.CreationModels;
using WaybillsAPI.Models;
using WaybillsAPI.ReportsModels.Fuel;
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
            CreateMap<Waybill, FuelWaybill>()
                .ForMember("FuelEconomy", opt => opt.MapFrom(x => x.NormalFuelConsumption - x.FactFuelConsumption));
            CreateMap<Operation, OperationCreation>();
            CreateMap<Calculation, CalculationCreation>();
            CreateMap<Waybill, ShortWaybillDTO>();
            CreateMap<Operation, OperationDTO>().ReverseMap();
            CreateMap<Calculation, CalculationDTO>().ReverseMap();
        }
    }
}
