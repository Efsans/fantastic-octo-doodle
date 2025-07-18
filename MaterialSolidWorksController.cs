using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WebPAIC_;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Swashbuckle.AspNetCore.Annotations; // Adicione esta linha

[Route("api/[controller]")]
[ApiController]
public class MaterialSolidWorksController : ControllerBase
{
    private readonly MyDbContext _context;

    public MaterialSolidWorksController(MyDbContext context)
    {
        _context = context;
    }

    /// <summary>
    /// Retorna uma lista de todos os Bancos de dados de materiais, incluindo a Biblioteca pai.
    /// </summary>
    /// <returns>Uma lista de Bancos de dados de materiais.</returns>
    [HttpGet]
    [SwaggerOperation(Summary = "Obtém todos os Bancos de dados", Description = "Retorna uma lista completa de todos os Bancos de dados de materiais, com suas Bibliotecas associadas.")]
    [ProducesResponseType(typeof(IEnumerable<MaterialSolidWorks>), 200)]
    public async Task<ActionResult<IEnumerable<MaterialSolidWorks>>> GetMaterialSolidWorks()
    {
        return await _context.Banco_de_dados.Include(m => m.Bliblioteca).ToListAsync();
    }

    /// <summary>
    /// Retorna um Banco de dados de materiais específico pelo seu ID.
    /// </summary>
    /// <param name="id">O ID do Banco de dados.</param>
    /// <returns>O Banco de dados encontrado ou NotFound se não existir.</returns>
    [HttpGet("{id}")]
    [SwaggerOperation(Summary = "Obtém um Banco de dados por ID", Description = "Retorna um único Banco de dados de materiais com base no ID fornecido, incluindo a Biblioteca pai.")]
    [ProducesResponseType(typeof(MaterialSolidWorks), 200)]
    [ProducesResponseType(404)]
    public async Task<ActionResult<MaterialSolidWorks>> GetMaterialSolidWorks(Guid id)
    {
        var materialSolidWorks = await _context.Banco_de_dados
                                            .Include(m => m.Bliblioteca)
                                            .FirstOrDefaultAsync(m => m.id_bank == id);

        if (materialSolidWorks == null)
        {
            return NotFound();
        }

        return materialSolidWorks;
    }

    /// <summary>
    /// Cria um novo Banco de dados de materiais.
    /// </summary>
    /// <param name="materialSolidWorks">Os dados do Banco de dados a ser criado, incluindo o IdBliblioteca.</param>
    /// <returns>O Banco de dados recém-criado.</returns>
    [HttpPost]
    [SwaggerOperation(Summary = "Cria um novo Banco de dados", Description = "Adiciona um novo Banco de dados de materiais ao sistema, associado a uma Biblioteca existente.")]
    [Consumes("application/json")]
    [ProducesResponseType(typeof(MaterialSolidWorks), 201)]
    [ProducesResponseType(400)]
    [ProducesResponseType(404)]
    public async Task<ActionResult<MaterialSolidWorks>> PostMaterialSolidWorks(MaterialSolidWorks materialSolidWorks)
    {
        // Validação: id_biblioteca é obrigatório
        if (materialSolidWorks.IdBliblioteca == Guid.Empty)
        {
            return BadRequest("O ID da Biblioteca (IdBliblioteca) é obrigatório para criar um Banco.");
        }

        // Verifica se a Bliblioteca pai existe
        var blibliotecaExists = await _context.Bliblioteca_de_materiais.AnyAsync(b => b.id_lib == materialSolidWorks.IdBliblioteca);
        if (!blibliotecaExists)
        {
            return NotFound($"A Biblioteca com ID '{materialSolidWorks.IdBliblioteca}' não foi encontrada.");
        }

        if (materialSolidWorks.id_bank == Guid.Empty)
        {
            materialSolidWorks.id_bank = Guid.NewGuid();
        }

        _context.Banco_de_dados.Add(materialSolidWorks);
        await _context.SaveChangesAsync();

        return CreatedAtAction(nameof(GetMaterialSolidWorks), new { id = materialSolidWorks.id_bank }, materialSolidWorks);
    }

    /// <summary>
    /// Atualiza um Banco de dados de materiais existente.
    /// </summary>
    /// <param name="id">O ID do Banco de dados a ser atualizado.</param>
    /// <param name="materialSolidWorks">Os novos dados do Banco de dados.</param>
    /// <returns>NoContent se a atualização for bem-sucedida.</returns>
    [HttpPut("{id}")]
    [SwaggerOperation(Summary = "Atualiza um Banco de dados existente", Description = "Atualiza completamente os dados de um Banco de dados de materiais.")]
    [Consumes("application/json")]
    [ProducesResponseType(204)]
    [ProducesResponseType(400)]
    [ProducesResponseType(404)]
    public async Task<IActionResult> PutMaterialSolidWorks(Guid id, MaterialSolidWorks materialSolidWorks)
    {
        if (id != materialSolidWorks.id_bank)
        {
            return BadRequest("O ID na URL não corresponde ao ID do Banco fornecido.");
        }

        // Validação: id_biblioteca é obrigatório (caso seja alterado na atualização)
        if (materialSolidWorks.IdBliblioteca == Guid.Empty)
        {
            return BadRequest("O ID da Biblioteca (IdBliblioteca) é obrigatório.");
        }

        // Verifica se a Bliblioteca pai existe
        var blibliotecaExists = await _context.Bliblioteca_de_materiais.AnyAsync(b => b.id_lib == materialSolidWorks.IdBliblioteca);
        if (!blibliotecaExists)
        {
            return NotFound($"A Biblioteca com ID '{materialSolidWorks.IdBliblioteca}' não foi encontrada.");
        }

        _context.Entry(materialSolidWorks).State = EntityState.Modified;

        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateConcurrencyException)
        {
            if (!MaterialSolidWorksExists(id))
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
    /// Exclui um Banco de dados de materiais.
    /// </summary>
    /// <param name="id">O ID do Banco de dados a ser excluído.</param>
    /// <returns>NoContent se a exclusão for bem-sucedida.</returns>
    [HttpDelete("{id}")]
    [SwaggerOperation(Summary = "Exclui um Banco de dados", Description = "Remove um Banco de dados de materiais do sistema pelo seu ID.")]
    [ProducesResponseType(204)]
    [ProducesResponseType(404)]
    public async Task<IActionResult> DeleteMaterialSolidWorks(Guid id)
    {
        var materialSolidWorks = await _context.Banco_de_dados.FindAsync(id);
        if (materialSolidWorks == null)
        {
            return NotFound();
        }

        _context.Banco_de_dados.Remove(materialSolidWorks);
        await _context.SaveChangesAsync();

        return NoContent();
    }

    private bool MaterialSolidWorksExists(Guid id)
    {
        return _context.Banco_de_dados.Any(e => e.id_bank == id);
    }
}