using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WaybillsAPI.Context;
using WaybillsAPI.Interfaces;
using WaybillsAPI.ReportsModels;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportsController(WaybillsContext context, IExcelWriter writer) : ControllerBase
    {
        private readonly WaybillsContext _context = context;
        private readonly IExcelWriter _writer = writer;

        [HttpGet("costCode/{year}/{month}/{price}")]
        public async Task<ActionResult<IEnumerable<CostCodeInfo>>> GetReport(int year, int month, double price)
        {
            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Operations).ToListAsync();

            var costPrices = new Dictionary<string, CostCodeInfoCreation>();
            foreach (var waybill in waybills)
            {
                foreach (var operation in waybill.Operations)
                {
                    if (operation.Norm == 0)
                    {
                        continue;
                    }
                    if (costPrices.TryGetValue(operation.ProductionCostCode, out CostCodeInfoCreation? value))
                    {
                        value.ConditionalReferenceHectares += operation.ConditionalReferenceHectares;
                        value.WaybillIdentifiers.TryAdd(waybill.Id, waybill.Number);
                        continue;
                    }
                    costPrices.Add(operation.ProductionCostCode, new(operation.ConditionalReferenceHectares, waybill.Id, waybill.Number));
                }
            }

            var report = costPrices.Select(x => new CostCodeInfo(x.Key, x.Value, price)).OrderBy(c => c.ProductionCostCode).ToList();
            return report;
        }

        [HttpGet("monthTotal/{year}/{month}")]
        public async Task<ActionResult<IEnumerable<DriverMonthTotal>>> GetMonthTotal(int year, int month)
        {
            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Operations).ToListAsync();

            var drivers = await _context.Drivers.ToDictionaryAsync(x => x.Id);
            var transports = await _context.Transport.ToDictionaryAsync(x => x.Id);

            var driverMonthTotals = waybills
                .GroupBy(x => x.DriverId)
                .Select(x => new DriverMonthTotal(drivers[x.Key], x.GroupBy(x => x.TransportId).Select(x => new TransportMonthTotal(transports[x.Key], [.. x])).ToList()))
                .OrderBy(x => x.DriverFullName).ToList();

            return driverMonthTotals;
        }
    }
}
