using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WebPAIC_;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Swashbuckle.AspNetCore.Annotations; // Adicione esta linha para annotations do Swagger

[Route("api/[controller]")]
[ApiController]
public class BlibliotecaController : ControllerBase
{
    private readonly MyDbContext _context;

    public BlibliotecaController(MyDbContext context)
    {
        _context = context;
    }

    /// <summary>
    /// Retorna uma lista de todas as Bibliotecas de materiais.
    /// </summary>
    /// <returns>Uma lista de Bibliotecas.</returns>
    [HttpGet]
    [SwaggerOperation(Summary = "Obtém todas as Bibliotecas", Description = "Retorna uma lista completa de todas as Bibliotecas de materiais cadastradas.")]
    [ProducesResponseType(typeof(IEnumerable<Bliblioteca>), 200)]
    public async Task<ActionResult<IEnumerable<Bliblioteca>>> GetBlibliotecas()
    {
        return await _context.Bliblioteca_de_materiais.ToListAsync();
    }

    /// <summary>
    /// Retorna uma Biblioteca específica pelo seu ID.
    /// </summary>
    /// <param name="id">O ID da Biblioteca.</param>
    /// <returns>A Biblioteca encontrada ou NotFound se não existir.</returns>
    [HttpGet("{id}")]
    [SwaggerOperation(Summary = "Obtém uma Biblioteca por ID", Description = "Retorna uma única Biblioteca de materiais com base no ID fornecido.")]
    [ProducesResponseType(typeof(Bliblioteca), 200)]
    [ProducesResponseType(404)]
    public async Task<ActionResult<Bliblioteca>> GetBliblioteca(Guid id)
    {
        var bliblioteca = await _context.Bliblioteca_de_materiais.FindAsync(id);

        if (bliblioteca == null)
        {
            return NotFound();
        }

        return bliblioteca;
    }

    /// <summary>
    /// Cria uma nova Biblioteca de materiais.
    /// </summary>
    /// <param name="bliblioteca">Os dados da Biblioteca a ser criada.</param>
    /// <returns>A Biblioteca recém-criada.</returns>
    [HttpPost]
    [SwaggerOperation(Summary = "Cria uma nova Biblioteca", Description = "Adiciona uma nova Biblioteca de materiais ao sistema.")]
    [Consumes("application/json")]
    [ProducesResponseType(typeof(Bliblioteca), 201)]
    [ProducesResponseType(400)]
    public async Task<ActionResult<Bliblioteca>> PostBliblioteca(Bliblioteca bliblioteca)
    {
        // Se o id_lib for Guid.Empty, o EF Core vai gerar um novo Guid automaticamente.
        // Caso contrário, ele tentará usar o Guid fornecido (se for único).
        if (bliblioteca.id_lib == Guid.Empty)
        {
            bliblioteca.id_lib = Guid.NewGuid();
        }

        _context.Bliblioteca_de_materiais.Add(bliblioteca);
        await _context.SaveChangesAsync();

        return CreatedAtAction(nameof(GetBliblioteca), new { id = bliblioteca.id_lib }, bliblioteca);
    }

    /// <summary>
    /// Atualiza uma Biblioteca de materiais existente.
    /// </summary>
    /// <param name="id">O ID da Biblioteca a ser atualizada.</param>
    /// <param name="bliblioteca">Os novos dados da Biblioteca.</param>
    /// <returns>NoContent se a atualização for bem-sucedida.</returns>
    [HttpPut("{id}")]
    [SwaggerOperation(Summary = "Atualiza uma Biblioteca existente", Description = "Atualiza completamente os dados de uma Biblioteca de materiais.")]
    [Consumes("application/json")]
    [ProducesResponseType(204)]
    [ProducesResponseType(400)]
    [ProducesResponseType(404)]
    public async Task<IActionResult> PutBliblioteca(Guid id, Bliblioteca bliblioteca)
    {
        if (id != bliblioteca.id_lib)
        {
            return BadRequest("O ID na URL não corresponde ao ID da biblioteca fornecida.");
        }

        _context.Entry(bliblioteca).State = EntityState.Modified;

        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateConcurrencyException)
        {
            if (!BlibliotecaExists(id))
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
    /// Exclui uma Biblioteca de materiais.
    /// </summary>
    /// <param name="id">O ID da Biblioteca a ser excluída.</param>
    /// <returns>NoContent se a exclusão for bem-sucedida.</returns>
    [HttpDelete("{id}")]
    [SwaggerOperation(Summary = "Exclui uma Biblioteca", Description = "Remove uma Biblioteca de materiais do sistema pelo seu ID.")]
    [ProducesResponseType(204)]
    [ProducesResponseType(404)]
    public async Task<IActionResult> DeleteBliblioteca(Guid id)
    {
        var bliblioteca = await _context.Bliblioteca_de_materiais.FindAsync(id);
        if (bliblioteca == null)
        {
            return NotFound();
        }

        _context.Bliblioteca_de_materiais.Remove(bliblioteca);
        await _context.SaveChangesAsync();

        return NoContent();
    }

    private bool BlibliotecaExists(Guid id)
    {
        return _context.Bliblioteca_de_materiais.Any(e => e.id_lib == id);
    }
}