using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WebPAIC_;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Swashbuckle.AspNetCore.Annotations; // Adicione esta linha

[Route("api/[controller]")]
[ApiController]
public class MaterialsController : ControllerBase
{
    private readonly MyDbContext _context;

    public MaterialsController(MyDbContext context)
    {
        _context = context;
    }

    /// <summary>
    /// Retorna uma lista de todos os Materiais, incluindo suas hierarquias completas (Sub-Banco, Banco, Biblioteca).
    /// </summary>
    /// <returns>Uma lista de Materiais.</returns>
    [HttpGet]
    [SwaggerOperation(Summary = "Obtém todos os Materiais", Description = "Retorna uma lista completa de todos os Materiais, com suas respectivas hierarquias (Sub-Banco, Banco, Biblioteca).")]
    [ProducesResponseType(typeof(IEnumerable<Materials>), 200)]
    public async Task<ActionResult<IEnumerable<Materials>>> GetMaterials()
    {
        return await _context.Materiais
                             .Include(m => m.SubMaterialSolidWorks)
                                 .ThenInclude(smsw => smsw.MaterialSolidWorks)
                                     .ThenInclude(msw => msw.Bliblioteca)
                             .ToListAsync();
    }

    /// <summary>
    /// Retorna um Material específico pelo seu ID, incluindo sua hierarquia completa.
    /// </summary>
    /// <param name="id">O ID do Material.</param>
    /// <returns>O Material encontrado ou NotFound se não existir.</returns>
    [HttpGet("{id}")]
    [SwaggerOperation(Summary = "Obtém um Material por ID", Description = "Retorna um único Material com base no ID fornecido, incluindo sua hierarquia completa.")]
    [ProducesResponseType(typeof(Materials), 200)]
    [ProducesResponseType(404)]
    public async Task<ActionResult<Materials>> GetMaterial(Guid id)
    {
        var material = await _context.Materiais
                                     .Include(m => m.SubMaterialSolidWorks)
                                         .ThenInclude(smsw => smsw.MaterialSolidWorks)
                                             .ThenInclude(msw => msw.Bliblioteca)
                                     .FirstOrDefaultAsync(m => m.id_material == id);

        if (material == null)
        {
            return NotFound();
        }

        return material;
    }

    /// <summary>
    /// Cria um novo Material, associando-o a um Sub-Banco existente.
    /// </summary>
    /// <param name="material">Os dados do Material a ser criado, incluindo o IdSubMaterialSolidWorks.</param>
    /// <returns>O Material recém-criado.</returns>
    [HttpPost]
    [SwaggerOperation(Summary = "Cria um novo Material (individual)", Description = "Adiciona um novo Material ao sistema, associando-o a um Sub-Banco existente. **Este endpoint não cria a hierarquia de Banco ou Biblioteca automaticamente.**")]
    [Consumes("application/json")]
    [ProducesResponseType(typeof(Materials), 201)]
    [ProducesResponseType(400)]
    [ProducesResponseType(404)]
    public async Task<ActionResult<Materials>> PostMaterial(Materials material)
    {
        // Validação: id_subbanco é obrigatório
        if (material.IdSubMaterialSolidWorks == Guid.Empty)
        {
            return BadRequest("O ID do SubBanco (IdSubMaterialSolidWorks) é obrigatório para criar um Material.");
        }

        // Verifica se o SubMaterialSolidWorks pai existe
        var subMaterialSolidWorksExists = await _context.Sub_banco.AnyAsync(s => s.id_sub == material.IdSubMaterialSolidWorks);
        if (!subMaterialSolidWorksExists)
        {
            return NotFound($"O SubBanco com ID '{material.IdSubMaterialSolidWorks}' não foi encontrado.");
        }

        if (material.id_material == Guid.Empty)
        {
            material.id_material = Guid.NewGuid();
        }

        _context.Materiais.Add(material);
        await _context.SaveChangesAsync();

        return CreatedAtAction(nameof(GetMaterial), new { id = material.id_material }, material);
    }

    /// <summary>
    /// Cria um novo Material, opcionalmente criando ou vinculando sua hierarquia completa (Biblioteca, Banco, Sub-Banco).
    /// </summary>
    /// <param name="request">Os dados para criar o Material e, opcionalmente, sua hierarquia.</param>
    /// <returns>O Material recém-criado com sua hierarquia.</returns>
    [HttpPost("CreateFullHierarchy")]
    [SwaggerOperation(Summary = "Cria Material com hierarquia completa", Description = "Cria um novo Material e, opcionalmente, sua Biblioteca, Banco de Dados e Sub-Banco se eles não existirem ou não forem fornecidos pelos IDs. Esta é a rota para criação 'completa'.")]
    [Consumes("application/json")]
    [ProducesResponseType(typeof(Materials), 201)]
    [ProducesResponseType(400)]
    [ProducesResponseType(404)]
    public async Task<ActionResult<Materials>> CreateFullHierarchy([FromBody] CreateFullMaterialHierarchyRequest request)
    {
        // 1. Criar ou encontrar a Biblioteca
        Bliblioteca bliblioteca;
        if (request.BlibliotecaId.HasValue && request.BlibliotecaId.Value != Guid.Empty)
        {
            bliblioteca = await _context.Bliblioteca_de_materiais.FindAsync(request.BlibliotecaId.Value);
            if (bliblioteca == null)
            {
                return NotFound($"Biblioteca com ID '{request.BlibliotecaId.Value}' não encontrada.");
            }
        }
        else
        {
            if (string.IsNullOrEmpty(request.BlibliotecaName))
            {
                return BadRequest("Nome da Biblioteca é obrigatório se nenhum BlibliotecaId for fornecido.");
            }
            bliblioteca = new Bliblioteca { id_lib = Guid.NewGuid(), name = request.BlibliotecaName };
            _context.Bliblioteca_de_materiais.Add(bliblioteca);
        }

        // 2. Criar ou encontrar o Banco (MaterialSolidWorks)
        MaterialSolidWorks materialSolidWorks;
        if (request.BancoId.HasValue && request.BancoId.Value != Guid.Empty)
        {
            materialSolidWorks = await _context.Banco_de_dados.FindAsync(request.BancoId.Value);
            if (materialSolidWorks == null)
            {
                return NotFound($"Banco com ID '{request.BancoId.Value}' não encontrado.");
            }
            // Verifica se o banco encontrado pertence à biblioteca correta, se aplicável
            if (materialSolidWorks.IdBliblioteca != bliblioteca.id_lib)
            {
                return BadRequest($"O Banco com ID '{request.BancoId.Value}' não pertence à Biblioteca com ID '{bliblioteca.id_lib}'.");
            }
        }
        else
        {
            if (string.IsNullOrEmpty(request.BancoName))
            {
                return BadRequest("Nome do Banco é obrigatório se nenhum BancoId for fornecido.");
            }
            materialSolidWorks = new MaterialSolidWorks
            {
                id_bank = Guid.NewGuid(),
                IdBliblioteca = bliblioteca.id_lib,
                name = request.BancoName
            };
            _context.Banco_de_dados.Add(materialSolidWorks);
        }

        // 3. Criar ou encontrar o SubBanco (SubMaterialSolidWorks)
        SubMaterialSolidWorks subMaterialSolidWorks;
        if (request.SubBancoId.HasValue && request.SubBancoId.Value != Guid.Empty)
        {
            subMaterialSolidWorks = await _context.Sub_banco.FindAsync(request.SubBancoId.Value);
            if (subMaterialSolidWorks == null)
            {
                return NotFound($"SubBanco com ID '{request.SubBancoId.Value}' não encontrado.");
            }
            // Verifica se o subbanco encontrado pertence ao banco correto, se aplicável
            if (subMaterialSolidWorks.IdMaterialSolidWorks != materialSolidWorks.id_bank)
            {
                return BadRequest($"O SubBanco com ID '{request.SubBancoId.Value}' não pertence ao Banco com ID '{materialSolidWorks.id_bank}'.");
            }
        }
        else
        {
            if (string.IsNullOrEmpty(request.SubBancoName))
            {
                return BadRequest("Nome do SubBanco é obrigatório se nenhum SubBancoId for fornecido.");
            }
            subMaterialSolidWorks = new SubMaterialSolidWorks
            {
                id_sub = Guid.NewGuid(),
                IdMaterialSolidWorks = materialSolidWorks.id_bank,
                name = request.SubBancoName
            };
            _context.Sub_banco.Add(subMaterialSolidWorks);
        }

        // 4. Criar o Material
        if (string.IsNullOrEmpty(request.MaterialName))
        {
            return BadRequest("Nome do Material é obrigatório.");
        }
        var newMaterial = new Materials
        {
            id_material = Guid.NewGuid(),
            IdSubMaterialSolidWorks = subMaterialSolidWorks.id_sub,
            name = request.MaterialName,
            description = request.Description,
            env_data = request.EnvData,
            app_data = request.AppData,
            name_reduz = request.NameReduz,
            angule = request.Angule,
            escale = request.Escale,
            tipo_selec = request.TipoSelec,
            patch_esp = request.PatchEsp,
            patch_esp_name = request.PatchEspName,
            patch_band = request.PatchBand,
            patch_band_name = request.PatchBandName,
            patch_calc = request.PatchCalc,
            patch_calc_name = request.PatchCalcName,
            mat_id = request.MatId // Certifique-se de que mat_id é tratado conforme sua lógica (gerado ou fornecido)
        };
        _context.Materiais.Add(newMaterial);

        await _context.SaveChangesAsync();

        // Para retornar o objeto completo com as navegações carregadas
        await _context.Entry(newMaterial)
                      .Reference(m => m.SubMaterialSolidWorks).LoadAsync();
        await _context.Entry(newMaterial.SubMaterialSolidWorks)
                      .Reference(smsw => smsw.MaterialSolidWorks).LoadAsync();
        await _context.Entry(newMaterial.SubMaterialSolidWorks.MaterialSolidWorks)
                      .Reference(msw => msw.Bliblioteca).LoadAsync();

        return CreatedAtAction(nameof(GetMaterial), new { id = newMaterial.id_material }, newMaterial);
    }


    /// <summary>
    /// Atualiza um Material existente.
    /// </summary>
    /// <param name="id">O ID do Material a ser atualizado.</param>
    /// <param name="material">Os novos dados do Material.</param>
    /// <returns>NoContent se a atualização for bem-sucedida.</returns>
    [HttpPut("{id}")]
    [SwaggerOperation(Summary = "Atualiza um Material existente", Description = "Atualiza completamente os dados de um Material.")]
    [Consumes("application/json")]
    [ProducesResponseType(204)]
    [ProducesResponseType(400)]
    [ProducesResponseType(404)]
    public async Task<IActionResult> PutMaterial(Guid id, Materials material)
    {
        if (id != material.id_material)
        {
            return BadRequest("O ID na URL não corresponde ao ID do material fornecido.");
        }

        // Validação: id_subbanco é obrigatório (caso seja alterado na atualização)
        if (material.IdSubMaterialSolidWorks == Guid.Empty)
        {
            return BadRequest("O ID do SubBanco (IdSubMaterialSolidWorks) é obrigatório.");
        }

        // Verifica se o SubMaterialSolidWorks pai existe
        var subMaterialSolidWorksExists = await _context.Sub_banco.AnyAsync(s => s.id_sub == material.IdSubMaterialSolidWorks);
        if (!subMaterialSolidWorksExists)
        {
            return NotFound($"O SubBanco com ID '{material.IdSubMaterialSolidWorks}' não foi encontrado.");
        }

        _context.Entry(material).State = EntityState.Modified;

        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateConcurrencyException)
        {
            if (!MaterialExists(id))
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
    /// Exclui um Material.
    /// </summary>
    /// <param name="id">O ID do Material a ser excluído.</param>
    /// <returns>NoContent se a exclusão for bem-sucedida.</returns>
    [HttpDelete("{id}")]
    [SwaggerOperation(Summary = "Exclui um Material", Description = "Remove um Material do sistema pelo seu ID.")]
    [ProducesResponseType(204)]
    [ProducesResponseType(404)]
    public async Task<IActionResult> DeleteMaterial(Guid id)
    {
        var material = await _context.Materiais.FindAsync(id);
        if (material == null)
        {
            return NotFound();
        }

        _context.Materiais.Remove(material);
        await _context.SaveChangesAsync();

        return NoContent();
    }

    private bool MaterialExists(Guid id)
    {
        return _context.Materiais.Any(e => e.id_material == id);
    }
}

// Classe DTO (Data Transfer Object) para a requisição de criação de hierarquia completa
// Crie este arquivo em uma pasta "Models" ou "DTOs" no seu projeto, se ainda não estiver lá.
public class CreateFullMaterialHierarchyRequest
{
    // Para Biblioteca
    [SwaggerSchema("Opcional: ID de uma Biblioteca existente. Se fornecido, BlibliotecaName é ignorado e a biblioteca existente é usada.")]
    public Guid? BlibliotecaId { get; set; } // Opcional: se a biblioteca já existe
    [SwaggerSchema("Nome da nova Biblioteca. Obrigatório se BlibliotecaId não for fornecido.")]
    public string BlibliotecaName { get; set; }

    // Para Banco (MaterialSolidWorks)
    [SwaggerSchema("Opcional: ID de um Banco de dados existente. Se fornecido, BancoName é ignorado e o banco existente é usado.")]
    public Guid? BancoId { get; set; } // Opcional: se o banco já existe
    [SwaggerSchema("Nome do novo Banco de dados. Obrigatório se BancoId não for fornecido.")]
    public string BancoName { get; set; }

    // Para SubBanco (SubMaterialSolidWorks)
    [SwaggerSchema("Opcional: ID de um Sub-Banco existente. Se fornecido, SubBancoName é ignorado e o sub-banco existente é usado.")]
    public Guid? SubBancoId { get; set; } // Opcional: se o subbanco já existe
    [SwaggerSchema("Nome do novo Sub-Banco. Obrigatório se SubBancoId não for fornecido.")]
    public string SubBancoName { get; set; }

    // Para Material (Materials)
    [SwaggerSchema("Nome do Material. Obrigatório.")]
    public string MaterialName { get; set; }
    public string Description { get; set; }
    public string EnvData { get; set; }
    public string AppData { get; set; }
    public string NameReduz { get; set; }
    public string Angule { get; set; }
    public string Escale { get; set; }
    public string TipoSelec { get; set; }
    public string PatchEsp { get; set; }
    public string PatchEspName { get; set; }
    public string PatchBand { get; set; }
    public string PatchBandName { get; set; }
    public string PatchCalc { get; set; }
    public string PatchCalcName { get; set; }
    public int MatId { get; set; } // Se for necessário enviar este ID
}