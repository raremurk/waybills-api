using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WaybillsAPI.Context;
using WaybillsAPI.Helpers;
using WaybillsAPI.Interfaces;
using WaybillsAPI.ReportsModels;
using WaybillsAPI.ReportsModels.CostPrice;
using WaybillsAPI.ReportsModels.Fuel.Drivers;
using WaybillsAPI.ReportsModels.Fuel.Transports;
using WaybillsAPI.ReportsModels.MonthTotal;

namespace WaybillsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportsController(WaybillsContext context, IOmnicommFuelService omnicommFuelService) : ControllerBase
    {
        private readonly WaybillsContext _context = context;
        private readonly IOmnicommFuelService _omnicommFuelService = omnicommFuelService;

        [HttpGet("omnicommFuel/{date}/{omnicommId}")]
        public async Task<ActionResult<FuelFromOmnicomm>> GetOmnicommFuel(DateOnly date, int omnicommId)
        {
            if (omnicommId == 0)
            {
                return Problem("Машина не подключена к Omnicomm");
            }

            var omnicommFuelReport = await _omnicommFuelService.GetOmnicommFuelReport(date, omnicommId);
            if (omnicommFuelReport.Code == 0)
            {
                var fuelData = omnicommFuelReport.Data.VehicleDataList.First();
                return new FuelFromOmnicomm(fuelData);
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
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Driver)
                .Include(x => x.Transport)
                .ToListAsync();

            var driversFuelMonthTotals = waybills
                .GroupBy(x => x.DriverId)
                .Select(x => new DetailedDriverFuelMonthTotal(x))
                .OrderBy(x => x.Driver.LastName)
                .ToList();

            return driversFuelMonthTotals;
        }

        [HttpGet("transportsFuelMonthTotal/{year}/{month}")]
        public async Task<ActionResult<IEnumerable<DetailedTransportFuelMonthTotal>>> GetTransportsFuelMonthTotal(int year, int month)
        {
            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Driver)
                .Include(x => x.Transport)
                .ToListAsync();

            var transportsFuelMonthTotals = waybills
                .GroupBy(x => x.TransportId)
                .Select(x => new DetailedTransportFuelMonthTotal(x))
                .OrderBy(x => Helper.PadNumbers(x.Transport.Name))
                .ToList();

            return transportsFuelMonthTotals;
        }
    }
}
