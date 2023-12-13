﻿using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Globalization;
using WaybillsAPI.Context;
using WaybillsAPI.CreationModels;
using WaybillsAPI.Interfaces;
using WaybillsAPI.Models;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WaybillsController(WaybillsContext context, IMapper mapper, IExcelWriter writer) : ControllerBase
    {
        private readonly WaybillsContext _context = context;
        private readonly IMapper _mapper = mapper;
        private readonly IExcelWriter _writer = writer;

        [HttpGet("{year}/{month}/{driverId?}")]
        public async Task<ActionResult<IEnumerable<ShortWaybillDTO>>> GetWaybills(int year, int month, int driverId = 0)
        {
            var waybills = await _context.Waybills
                .Where(x => x.Date.Year == year && x.Date.Month == month && (driverId == 0 || x.DriverId == driverId))
                .Include(x => x.Driver)
                .Include(x => x.Transport).ToListAsync();

            var waybillsDTO = _mapper.Map<List<Waybill>, List<ShortWaybillDTO>>(waybills);
            return waybillsDTO;
        }

        [HttpGet("excel/{year}/{month}/{driverId}")]
        public async Task<ActionResult> GetExcelWithWaybills(int year, int month, int driverId)
        {
            var driver = await _context.Drivers.FindAsync(driverId);
            if (driver == null)
            {
                return Problem("Водителя с заданным ID не существует.");
            }
            var fileName = $"{driver.ShortFullName()}_{CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(month)}_{year}.xlsx";
            var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            var waybills = await _context.Waybills
                .Where(x => x.Date.Year == year && x.Date.Month == month && x.DriverId == driverId)
                .Include(x => x.Driver)
                .Include(x => x.Transport)
                .Include(x => x.Operations)
                .Include(x => x.Calculations)
                .OrderBy(x => x.Date).ToListAsync();
            return File(_writer.Generate(waybills), contentType, fileName);
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
        public async Task<IActionResult> PutWaybill(int id, WaybillCreation waybillCreation)
        {
            if (id != waybillCreation.Id)
            {
                return BadRequest();
            }

            var transport = _context.Transport.Find(waybillCreation.TransportId);
            if (transport is null)
            {
                return Problem("Транспорта с заданным ID не существует.");
            }
            var waybill = new Waybill(waybillCreation, transport.Coefficient);

            var existingWaybill = _context.Waybills
                .Include(p => p.Operations)
                .Include(p => p.Calculations)
                .FirstOrDefault(p => p.Id == id);
            if (existingWaybill != null)
            {
                _context.Entry(existingWaybill).CurrentValues.SetValues(waybill);
                existingWaybill.Operations.Clear();
                existingWaybill.Calculations.Clear();
                existingWaybill.Operations.AddRange(waybill.Operations);
                existingWaybill.Calculations.AddRange(waybill.Calculations);
            }

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!WaybillExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return NoContent();
        }

        [HttpPost]
        public async Task<ActionResult<WaybillDTO>> PostWaybill(WaybillCreation waybillCreation)
        {
            if (_context.Waybills.Any(x => x.Number == waybillCreation.Number))
            {
                return Problem($"Путевой лист №{waybillCreation.Number} уже существует.");
            }

            if (!_context.Drivers.Any(x => x.Id == waybillCreation.DriverId))
            {
                return Problem("Водителя с заданным ID не существует.");
            }

            var transport = _context.Transport.Find(waybillCreation.TransportId);
            if (transport is null)
            {
                return Problem("Транспорта с заданным ID не существует.");
            }

            var waybill = new Waybill(waybillCreation, transport.Coefficient);
            _context.Waybills.Add(waybill);
            await _context.SaveChangesAsync();
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

        private bool WaybillExists(int id)
        {
            return _context.Waybills.Any(e => e.Id == id);
        }
    }
}
