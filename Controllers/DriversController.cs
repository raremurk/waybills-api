using AutoMapper;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WaybillsAPI.Context;
using WaybillsAPI.Models;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DriversController(WaybillsContext context, IMapper mapper) : ControllerBase
    {
        private readonly WaybillsContext _context = context;
        private readonly IMapper _mapper = mapper;

        [HttpGet]
        public async Task<ActionResult<IEnumerable<DriverDTO>>> GetDrivers()
        {
            var drivers = await _context.Drivers.OrderBy(x => x.LastName).ToListAsync();
            var driversDTO = _mapper.Map<List<Driver>, List<DriverDTO>>(drivers);
            return driversDTO;
        }

        [HttpGet("{id}")]
        public async Task<ActionResult<DriverDTO>> GetDriver(int id)
        {
            var driver = await _context.Drivers.FindAsync(id);

            if (driver == null)
            {
                return NotFound();
            }

            var driverDTO = _mapper.Map<Driver, DriverDTO>(driver);
            return driverDTO;
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> PutDriver(int id, DriverDTO driverDTO)
        {
            if (id != driverDTO.Id)
            {
                return BadRequest();
            }

            var driver = _mapper.Map<DriverDTO, Driver>(driverDTO);
            _context.Entry(driver).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!DriverExists(id))
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
        public async Task<ActionResult<DriverDTO>> PostDriver(DriverDTO driverDTO)
        {
            var driver = _mapper.Map<DriverDTO, Driver>(driverDTO);
            _context.Drivers.Add(driver);
            await _context.SaveChangesAsync();
            var createdDriver = _mapper.Map<Driver, DriverDTO>(driver);

            return CreatedAtAction("GetDriver", new { id = createdDriver.Id }, createdDriver);
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteDriver(int id)
        {
            var driver = await _context.Drivers.FindAsync(id);
            if (driver == null)
            {
                return NotFound();
            }

            if (_context.Waybills.Any(x => x.DriverId == id))
            {
                return Problem("OBEMA", "", 403);
            }

            _context.Drivers.Remove(driver);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool DriverExists(int id)
        {
            return _context.Drivers.Any(e => e.Id == id);
        }
    }
}
