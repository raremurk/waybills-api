using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;
using Microsoft.EntityFrameworkCore;
using WaybillsAPI.Context;
using WaybillsAPI.CreationModels;
using WaybillsAPI.Interfaces;
using WaybillsAPI.Models;
using WaybillsAPI.ReportsModels.Fuel;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WaybillsController(WaybillsContext context, IMapper mapper, ISalaryPeriodService salaryPeriodService) : ControllerBase
    {
        private readonly WaybillsContext _context = context;
        private readonly IMapper _mapper = mapper;
        private readonly ISalaryPeriodService _salaryPeriodService = salaryPeriodService;

        [HttpGet("{year}/{month}/{driverId?}")]
        public async Task<ActionResult<IEnumerable<ShortWaybillDTO>>> GetWaybills(int year, int month, int driverId = 0)
        {
            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month && (driverId == 0 || x.DriverId == driverId))
                .Include(x => x.Driver)
                .Include(x => x.Transport).ToListAsync();

            var waybillsDTO = _mapper.Map<List<Waybill>, List<ShortWaybillDTO>>(waybills);
            return waybillsDTO;
        }

        [HttpGet("fuelOnly/{year}/{month}/{driverId?}/{transportId?}")]
        public async Task<ActionResult<IEnumerable<FuelWaybill>>> GetFuelWaybills(int year, int month, int driverId = 0, int transportId = 0)
        {
            if (driverId == 0 && transportId == 0)
            {
                return new List<FuelWaybill>();
            }

            var waybills = await _context.Waybills
                .Where(x => x.SalaryYear == year && x.SalaryMonth == month
                    && (x.DriverId == driverId || driverId == 0) && (x.TransportId == transportId || transportId == 0))
                .Include(x => x.Driver)
                .Include(x => x.Transport).ToListAsync();

            var fuelWaybills = _mapper.Map<List<Waybill>, List<FuelWaybill>>(waybills).OrderBy(x => x.Date).ToList();
            return fuelWaybills;
        }

        [HttpGet("{id}")]
        public async Task<ActionResult<WaybillDTO>> GetWaybill(int id)
        {
            var waybill = await _context.Waybills
                .Include(x => x.Driver)
                .Include(x => x.Transport)
                .Include(x => x.Operations)
                .Include(x => x.Calculations)
                .FirstOrDefaultAsync(x => x.Id == id);

            if (waybill == null)
            {
                return NotFound();
            }

            var waybillDTO = _mapper.Map<Waybill, WaybillDTO>(waybill);
            return waybillDTO;
        }

        [HttpPut("{id}")]
        public async Task<ActionResult<WaybillDTO>> PutWaybill(int id, WaybillCreation waybillCreation)
        {
            if (id != waybillCreation.Id)
            {
                return BadRequest();
            }

            var existingWaybill = await _context.Waybills
                .Include(x => x.Operations)
                .Include(x => x.Calculations)
                .FirstOrDefaultAsync(x => x.Id == id);
            if (existingWaybill == null)
            {
                return NotFound();
            }

            try
            {
                existingWaybill.Edit(_context, _salaryPeriodService, waybillCreation);
            }
            catch (ArgumentException exception)
            {
                return Problem(exception.Message);
            }

            await _context.SaveChangesAsync();

            var editedWaybill = _mapper.Map<Waybill, WaybillDTO>(existingWaybill);

            return Ok(editedWaybill);
        }

        [HttpPost]
        public async Task<ActionResult<WaybillDTO>> PostWaybill(WaybillCreation waybillCreation)
        {
            Waybill waybill;
            try
            {
                waybill = new Waybill(_context, _salaryPeriodService, waybillCreation);
            }
            catch (ArgumentException exception)
            {
                return Problem(exception.Message);
            }

            _context.Waybills.Add(waybill);
            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateException ex) when (ex.InnerException is SqliteException s && s.SqliteExtendedErrorCode == 2067)
            {
                return Problem($"Ошибка сохранения (такой путевой лист уже был создан).");
            }

            var createdWaybill = _mapper.Map<Waybill, WaybillDTO>(waybill);

            return CreatedAtAction("GetWaybill", new { id = createdWaybill.Id }, createdWaybill);
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteWaybill(int id)
        {
            var waybill = await _context.Waybills.FindAsync(id);
            if (waybill == null)
            {
                return NotFound();
            }

            _context.Waybills.Remove(waybill);
            await _context.SaveChangesAsync();

            return NoContent();
        }
    }
}
