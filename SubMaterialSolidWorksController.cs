using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WebPAIC_;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Swashbuckle.AspNetCore.Annotations; // Adicione esta linha

[Route("api/[controller]")]
[ApiController]
public class SubMaterialSolidWorksController : ControllerBase
{
    private readonly MyDbContext _context;

    public SubMaterialSolidWorksController(MyDbContext context)
    {
        _context = context;
    }

    /// <summary>
    /// Retorna uma lista de todos os Sub-Bancos de materiais, incluindo o Banco pai.
    /// </summary>
    /// <returns>Uma lista de Sub-Bancos de materiais.</returns>
    [HttpGet]
    [SwaggerOperation(Summary = "Obtém todos os Sub-Bancos", Description = "Retorna uma lista completa de todos os Sub-Bancos de materiais, com seus Bancos de dados associados.")]
    [ProducesResponseType(typeof(IEnumerable<SubMaterialSolidWorks>), 200)]
    public async Task<ActionResult<IEnumerable<SubMaterialSolidWorks>>> GetSubMaterialSolidWorks()
    {
        return await _context.Sub_banco.Include(s => s.MaterialSolidWorks).ToListAsync();
    }

    /// <summary>
    /// Retorna um Sub-Banco de materiais específico pelo seu ID.
    /// </summary>
    /// <param name="id">O ID do Sub-Banco.</param>
    /// <returns>O Sub-Banco encontrado ou NotFound se não existir.</returns>
    [HttpGet("{id}")]
    [SwaggerOperation(Summary = "Obtém um Sub-Banco por ID", Description = "Retorna um único Sub-Banco de materiais com base no ID fornecido, incluindo o Banco pai.")]
    [ProducesResponseType(typeof(SubMaterialSolidWorks), 200)]
    [ProducesResponseType(404)]
    public async Task<ActionResult<SubMaterialSolidWorks>> GetSubMaterialSolidWorks(Guid id)
    {
        var subMaterialSolidWorks = await _context.Sub_banco
                                                .Include(s => s.MaterialSolidWorks)
                                                .FirstOrDefaultAsync(s => s.id_sub == id);

        if (subMaterialSolidWorks == null)
        {
            return NotFound();
        }

        return subMaterialSolidWorks;
    }

    /// <summary>
    /// Cria um novo Sub-Banco de materiais.
    /// </summary>
    /// <param name="subMaterialSolidWorks">Os dados do Sub-Banco a ser criado, incluindo o IdMaterialSolidWorks.</param>
    /// <returns>O Sub-Banco recém-criado.</returns>
    [HttpPost]
    [SwaggerOperation(Summary = "Cria um novo Sub-Banco", Description = "Adiciona um novo Sub-Banco de materiais ao sistema, associado a um Banco de dados existente.")]
    [Consumes("application/json")]
    [ProducesResponseType(typeof(SubMaterialSolidWorks), 201)]
    [ProducesResponseType(400)]
    [ProducesResponseType(404)]
    public async Task<ActionResult<SubMaterialSolidWorks>> PostSubMaterialSolidWorks(SubMaterialSolidWorks subMaterialSolidWorks)
    {
        // Validação: id_banco é obrigatório
        if (subMaterialSolidWorks.IdMaterialSolidWorks == Guid.Empty)
        {
            return BadRequest("O ID do Banco (IdMaterialSolidWorks) é obrigatório para criar um SubBanco.");
        }

        // Verifica se o MaterialSolidWorks pai existe
        var materialSolidWorksExists = await _context.Banco_de_dados.AnyAsync(m => m.id_bank == subMaterialSolidWorks.IdMaterialSolidWorks);
        if (!materialSolidWorksExists)
        {
            return NotFound($"O Banco com ID '{subMaterialSolidWorks.IdMaterialSolidWorks}' não foi encontrado.");
        }

        if (subMaterialSolidWorks.id_sub == Guid.Empty)
        {
            subMaterialSolidWorks.id_sub = Guid.NewGuid();
        }

        _context.Sub_banco.Add(subMaterialSolidWorks);
        await _context.SaveChangesAsync();

        return CreatedAtAction(nameof(GetSubMaterialSolidWorks), new { id = subMaterialSolidWorks.id_sub }, subMaterialSolidWorks);
    }

    /// <summary>
    /// Atualiza um Sub-Banco de materiais existente.
    /// </summary>
    /// <param name="id">O ID do Sub-Banco a ser atualizado.</param>
    /// <param name="subMaterialSolidWorks">Os novos dados do Sub-Banco.</param>
    /// <returns>NoContent se a atualização for bem-sucedida.</returns>
    [HttpPut("{id}")]
    [SwaggerOperation(Summary = "Atualiza um Sub-Banco existente", Description = "Atualiza completamente os dados de um Sub-Banco de materiais.")]
    [Consumes("application/json")]
    [ProducesResponseType(204)]
    [ProducesResponseType(400)]
    [ProducesResponseType(404)]
    public async Task<IActionResult> PutSubMaterialSolidWorks(Guid id, SubMaterialSolidWorks subMaterialSolidWorks)
    {
        if (id != subMaterialSolidWorks.id_sub)
        {
            return BadRequest("O ID na URL não corresponde ao ID do SubBanco fornecido.");
        }

        // Validação: id_banco é obrigatório (caso seja alterado na atualização)
        if (subMaterialSolidWorks.IdMaterialSolidWorks == Guid.Empty)
        {
            return BadRequest("O ID do Banco (IdMaterialSolidWorks) é obrigatório.");
        }

        // Verifica se o MaterialSolidWorks pai existe
        var materialSolidWorksExists = await _context.Banco_de_dados.AnyAsync(m => m.id_bank == subMaterialSolidWorks.IdMaterialSolidWorks);
        if (!materialSolidWorksExists)
        {
            return NotFound($"O Banco com ID '{subMaterialSolidWorks.IdMaterialSolidWorks}' não foi encontrado.");
        }

        _context.Entry(subMaterialSolidWorks).State = EntityState.Modified;

        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateConcurrencyException)
        {
            if (!SubMaterialSolidWorksExists(id))
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

    /// <summary>
    /// Exclui um Sub-Banco de materiais.
    /// </summary>
    /// <param name="id">O ID do Sub-Banco a ser excluído.</param>
    /// <returns>NoContent se a exclusão for bem-sucedida.</returns>
    [HttpDelete("{id}")]
    [SwaggerOperation(Summary = "Exclui um Sub-Banco", Description = "Remove um Sub-Banco de materiais do sistema pelo seu ID.")]
    [ProducesResponseType(204)]
    [ProducesResponseType(404)]
    public async Task<IActionResult> DeleteSubMaterialSolidWorks(Guid id)
    {
        var subMaterialSolidWorks = await _context.Sub_banco.FindAsync(id);
        if (subMaterialSolidWorks == null)
        {
            return NotFound();
        }

        _context.Sub_banco.Remove(subMaterialSolidWorks);
        await _context.SaveChangesAsync();

        return NoContent();
    }

    private bool SubMaterialSolidWorksExists(Guid id)
    {
        return _context.Sub_banco.Any(e => e.id_sub == id);
    }
}