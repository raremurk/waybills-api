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
    public class TransportsController(WaybillsContext context, IMapper mapper) : ControllerBase
    {
        private readonly WaybillsContext _context = context;
        private readonly IMapper _mapper = mapper;

        [HttpGet]
        public async Task<ActionResult<IEnumerable<TransportDTO>>> GetTransport()
        {
            var transport = await _context.Transport.ToListAsync();
            var transportDTO = _mapper.Map<List<Transport>, List<TransportDTO>>(transport);
            return transportDTO;
        }

        [HttpGet("{id}")]
        public async Task<ActionResult<TransportDTO>> GetTransport(int id)
        {
            var transport = await _context.Transport.FindAsync(id);

            if (transport == null)
            {
                return NotFound();
            }

            var transportDTO = _mapper.Map<Transport, TransportDTO>(transport);
            return transportDTO;
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> PutTransport(int id, TransportDTO transportDTO)
        {
            if (id != transportDTO.Id)
            {
                return BadRequest();
            }

            var transport = _mapper.Map<TransportDTO, Transport>(transportDTO);
            _context.Entry(transport).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!TransportExists(id))
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
        public async Task<ActionResult<TransportDTO>> PostTransport(TransportDTO transportDTO)
        {
            var transport = _mapper.Map<TransportDTO, Transport>(transportDTO);

            _context.Transport.Add(transport);
            await _context.SaveChangesAsync();

            var createdTransport = _mapper.Map<Transport, TransportDTO>(transport);

            return CreatedAtAction("GetTransport", new { id = createdTransport.Id }, createdTransport);
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteTransport(int id)
        {
            var transport = await _context.Transport.FindAsync(id);
            if (transport == null)
            {
                return NotFound();
            }

            if (_context.Waybills.Any(x => x.TransportId == id))
            {
                return Problem("OBEMA", "",403);
            }

            _context.Transport.Remove(transport);
            await _context.SaveChangesAsync();

            return NoContent();
        }

        private bool TransportExists(int id)
        {
            return _context.Transport.Any(e => e.Id == id);
        }
    }
}
