using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WaybillsAPI.Context;
using WaybillsAPI.Helpers;
using WaybillsAPI.Models;

namespace WaybillsAPI.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RatesController(WaybillsContext context) : ControllerBase
    {
        private readonly WaybillsContext _context = context;

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Rate>>> GetRates()
        {
            var rates = await _context.Rates.ToListAsync();
            var orderedRates = rates.OrderBy(x => Helper.PadNumbers(x.Name)).ToList();
            return orderedRates;
        }

        [HttpGet("{id}")]
        public async Task<ActionResult<Rate>> GetRate(int id)
        {
            var rate = await _context.Rates.FindAsync(id);

            if (rate == null)
            {
                return NotFound();
            }

            return rate;
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> PutRate(int id, Rate rate)
        {
            if (id != rate.Id)
            {
                return BadRequest();
            }

            _context.Entry(rate).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!RateExists(id))
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
        public async Task<ActionResult<Rate>> PostRate(Rate rate)
        {
            _context.Rates.Add(rate);
            await _context.SaveChangesAsync();

            return CreatedAtAction("GetRate", new { id = rate.Id }, rate);
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteRate(int id)
        {
            var rate = await _context.Rates.FindAsync(id);
            if (rate == null)
            {
                return NotFound();
            }

            _context.Rates.Remove(rate);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool RateExists(int id)
        {
            return _context.Rates.Any(e => e.Id == id);
        }
    }
}
