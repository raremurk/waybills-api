using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WaybillsAPI.Context;
using WaybillsAPI.Interfaces;
using WaybillsAPI.Models;
using WaybillsAPI.ReportsModels;
using WaybillsAPI.ReportsModels.CostPrice;
using WaybillsAPI.ReportsModels.Fuel.Drivers;
using WaybillsAPI.ReportsModels.Fuel.Transports;
using WaybillsAPI.ReportsModels.MonthTotal;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportsController(WaybillsContext context, IMapper mapper, IOmnicommFuelService omnicommFuelService) : ControllerBase
    {
        private readonly WaybillsContext _context = context;
        private readonly IMapper _mapper = mapper;
        private readonly IOmnicommFuelService _omnicommFuelService = omnicommFuelService;

        [HttpGet("omnicommFuel/{data}/{omnicommId}")]
        public async Task<ActionResult<FuelFromOmnicomm>> GetOmnicommFuel(DateOnly data, int omnicommId)
        {
            if (omnicommId == 0)
            {
                return Problem("Машина не подключена к Omnicomm");
            }

            var omnicommFuelReport = await _omnicommFuelService.GetOmnicommFuelReport(data, omnicommId);
            if (omnicommFuelReport.Code == 0)
            {
                var fuelData = omnicommFuelReport.Data.VehicleDataList.First();
                return new FuelFromOmnicomm()
                {
                    TransportName = fuelData.Name,
                    StartFuel = fuelData.Fuel.StartVolume / 10 ?? 0d,
                    FuelTopUp = fuelData.Fuel.Refuelling / 10 ?? 0d,
                    EndFuel = fuelData.Fuel.EndVolume / 10 ?? 0d,
                    FuelConsumption = fuelData.Fuel.FuelConsumption / 10 ?? 0d,
                    Draining = fuelData.Fuel.Draining / 10 ?? 0d
                };
            }
            else
            {
                return Problem(omnicommFuelReport.Message);
            }
        }

        [HttpGet("costCode/{year}/{month}/{price}")]
        public async Task<ActionResult<IEnumerable<CostCodeInfo>>> GetCostPriceReport(int year, int month, double price)
        {
            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Operations)
                .ToListAsync();

            var report = new CostPriceReport(waybills, price);
            return report.CostCodes.ToList();
        }

        [HttpGet("monthTotal/{year}/{month}/{mainEntity?}")]
        public async Task<ActionResult<IEnumerable<DetailedEntityMonthTotal>>> GetMonthTotal(int year, int month, string mainEntity = "driver")
        {
            var byDrivers = string.Equals(mainEntity, "driver", StringComparison.OrdinalIgnoreCase);
            var byTransports = string.Equals(mainEntity, "transport", StringComparison.OrdinalIgnoreCase);
            if (!byDrivers && !byTransports) return BadRequest();

            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Driver)
                .Include(x => x.Transport)
                .Include(x => x.Operations)
                .ToListAsync();

            var report = new MonthTotalReport(waybills, byDrivers);
            return report.DetailedEntityMonthTotals.ToList();
        }

        [HttpGet("driversFuelMonthTotal/{year}/{month}")]
        public async Task<ActionResult<IEnumerable<DetailedDriverFuelMonthTotal>>> GetDriversFuelMonthTotal(int year, int month)
        {
            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month).ToListAsync();

            var drivers = await _context.Drivers.ToDictionaryAsync(x => x.Id);
            var transports = await _context.Transport.ToDictionaryAsync(x => x.Id);
            DriverDTO getDriverDTO(int driverId) => _mapper.Map<Driver, DriverDTO>(drivers[driverId]);
            TransportDTO getTransportDTO(int transportId) => _mapper.Map<Transport, TransportDTO>(transports[transportId]);

            var driversFuelMonthTotals = waybills
                .GroupBy(x => x.DriverId)
                .Select(x => new DetailedDriverFuelMonthTotal(getDriverDTO(x.Key), x.GroupBy(x => x.TransportId).Select(x => new DriverFuelSubTotal(getTransportDTO(x.Key), x))))
                .OrderBy(x => x.Driver.LastName).ToList();

            return driversFuelMonthTotals;
        }

        [HttpGet("transportsFuelMonthTotal/{year}/{month}")]
        public async Task<ActionResult<IEnumerable<DetailedTransportFuelMonthTotal>>> GetTransportsFuelMonthTotal(int year, int month)
        {
            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Driver).ToListAsync();

            var drivers = await _context.Drivers.ToDictionaryAsync(x => x.Id);
            var driversDTO = _mapper.Map<Dictionary<int, Driver>, Dictionary<int, DriverDTO>>(drivers);

            var transports = await _context.Transport.ToDictionaryAsync(x => x.Id);
            TransportDTO getTransportDTO(int transportId) => _mapper.Map<Transport, TransportDTO>(transports[transportId]);

            var transportsFuelMonthTotals = waybills
                .GroupBy(x => x.TransportId)
                .Select(x => new DetailedTransportFuelMonthTotal(getTransportDTO(x.Key), driversDTO, x))
                .OrderBy(x => x.Transport.Name).ToList();

            return transportsFuelMonthTotals;
        }
    }
}
