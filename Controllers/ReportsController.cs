using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WaybillsAPI.Context;
using WaybillsAPI.Interfaces;
using WaybillsAPI.Models;
using WaybillsAPI.ReportsModels;
using WaybillsAPI.ReportsModels.Fuel.Drivers;
using WaybillsAPI.ReportsModels.Fuel.Transports;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportsController(WaybillsContext context, IExcelWriter writer, IMapper mapper) : ControllerBase
    {
        private readonly WaybillsContext _context = context;
        private readonly IExcelWriter _writer = writer;
        private readonly IMapper _mapper = mapper;

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

        [HttpGet("monthTotal/{year}/{month}/{mainEntity?}")]
        public async Task<ActionResult<IEnumerable<DetailedEntityMonthTotal>>> GetMonthTotal(int year, int month, string mainEntity = "driver")
        {
            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month)
                .Include(x => x.Operations).ToListAsync();

            var drivers = await _context.Drivers.ToDictionaryAsync(x => x.Id);
            var transports = await _context.Transport.ToDictionaryAsync(x => x.Id);

            var mainGroupBy = (Waybill x) => x.DriverId;
            var childGroupBy = (Waybill x) => x.TransportId;
            var mainEntityName = (int id) => drivers[id].ShortFullName();
            var childEntityName = (int id) => transports[id].Name;
            var mainEntityCode = (int id) => drivers[id].PersonnelNumber;
            var childEntityCode = (int id) => transports[id].Code;
            if (string.Compare(mainEntity, "transport", StringComparison.InvariantCultureIgnoreCase) == 0)
            {
                (childGroupBy, mainGroupBy, childEntityName, mainEntityName, childEntityCode, mainEntityCode)
                    = (mainGroupBy, childGroupBy, mainEntityName, childEntityName, mainEntityCode, childEntityCode);
            }

            var detailedMonthTotals = waybills
                .GroupBy(mainGroupBy)
                .Select(x => new DetailedEntityMonthTotal(mainEntityName(x.Key), mainEntityCode(x.Key), x.GroupBy(childGroupBy).Select(x => new EntityMonthTotal(childEntityName(x.Key), childEntityCode(x.Key), [.. x]))))
                .OrderBy(x => x.EntityName).ToList();

            return detailedMonthTotals;
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
