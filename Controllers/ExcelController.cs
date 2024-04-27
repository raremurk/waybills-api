using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Globalization;
using WaybillsAPI.Context;
using WaybillsAPI.Interfaces;
using WaybillsAPI.ReportsModels.CostPrice;

namespace WaybillsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController(WaybillsContext context, IExcelWriter writer) : ControllerBase
    {
        private readonly string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        private readonly WaybillsContext _context = context;
        private readonly IExcelWriter _writer = writer;

        [HttpGet("costCode/{year}/{month}/{price}")]
        public async Task<ActionResult> GetCostPriceReport(int year, int month, double price)
        {
            var fileName = $"Себестоимость_{CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(month)}_{year}.xlsx";

            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Operations)
                .ToListAsync();

            var report = new CostPriceReport(waybills, price);
            var fileContents = _writer.GenerateCostPriceReport(report, year, month);
            return File(fileContents, contentType, fileName);
        }

        [HttpGet("monthTotal/{year}/{month}")]
        public async Task<ActionResult> GetMonthTotal(int year, int month)
        {
            var fileName = $"Итоги работы {CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(month)} {year}.xlsx";

            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Driver)
                .Include(x => x.Transport)
                .Include(x => x.Operations)
                .ToListAsync();

            var fileContents = _writer.GenerateMonthTotal(waybills);
            return File(fileContents, contentType, fileName);
        }

        [HttpGet("waybills/short/{year}/{month}")]
        public async Task<ActionResult> GetShortWaybills(int year, int month)
        {
            var fileName = $"Путевые листы кратко {CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(month)} {year}.xlsx";

            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Driver)
                .Include(x => x.Transport)
                .Include(x => x.Operations)
                .OrderBy(x => x.Date)
                .ToListAsync();

            var fileContents = _writer.GenerateShortWaybills(waybills);
            return File(fileContents, contentType, fileName);
        }


        [HttpGet("waybills/detailed/{year}/{month}/{driverId?}")]
        public async Task<ActionResult> GetDetailedWaybills(int year, int month, int driverId = 0)
        {
            var driverName = "все водители";
            if (driverId != 0)
            {
                var driver = await _context.Drivers.FindAsync(driverId);
                if (driver == null)
                {
                    return Problem("Водителя с заданным ID не существует.");
                }
                driverName = driver.ShortFullName();
            }

            var fileName = $"Путевые листы {driverName} {CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(month)} {year}.xlsx";

            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month && (driverId == 0 || x.DriverId == driverId))
                .Include(x => x.Driver)
                .Include(x => x.Transport)
                .Include(x => x.Operations)
                .Include(x => x.Calculations)
                .OrderBy(x => x.Date)
                .ToListAsync();

            var fileContents = _writer.GenerateDetailedWaybills(waybills);
            return File(fileContents, contentType, fileName);
        }
    }
}
