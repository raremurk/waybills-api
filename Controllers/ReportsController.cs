﻿using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WaybillsAPI.Context;
using WaybillsAPI.Interfaces;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportsController(WaybillsContext context, IMapper mapper, IExcelWriter writer) : ControllerBase
    {
        private readonly WaybillsContext _context = context;
        private readonly IMapper _mapper = mapper;
        private readonly IExcelWriter _writer = writer;

        [HttpGet("reports/{month}")]
        public async Task<ActionResult<IEnumerable<CostCodeInfo>>> GetReport(int month)
        {
            var waybills = await _context.Waybills
                .Where(x => x.Date.Month == month)
                .Include(x => x.Operations).ToListAsync();

            var costPrices = new Dictionary<string, CostCodeInfo>();
            foreach (var waybill in waybills)
            {
                foreach (var operation in waybill.Operations)
                {
                    if (costPrices.TryGetValue(operation.ProductionCostCode, out CostCodeInfo? value))
                    {
                        value.ConditionalReferenceHectares += operation.ConditionalReferenceHectares;
                        continue;
                    }
                    var costPrice = new CostCodeInfo(operation.ProductionCostCode, operation.ConditionalReferenceHectares);
                    costPrices.Add(operation.ProductionCostCode, costPrice);
                }
            }

            var report = costPrices.Values.OrderBy(x => x.ProductionCostCode).ToList();
            report.ForEach(x =>
            {
                x.ConditionalReferenceHectares = Math.Round(x.ConditionalReferenceHectares, 2);
                x.CostPrice = Math.Round(x.ConditionalReferenceHectares * 27, 2);
            });
            return report;
        }

        [HttpGet]
        public async Task<ActionResult> GetWaybills()
        {
            var waybills = await _context.Waybills
                .Include(x => x.Driver)
                .Include(x => x.Transport)
                .Include(x => x.Operations)
                .Include(x => x.Calculations).ToListAsync();
            return File(_writer.Generate(waybills), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Report.xlsx");
        }
    }
}
