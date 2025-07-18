using Microsoft.AspNetCore.Mvc.ModelBinding.Validation;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;


namespace WebPAIC_
{
    public class MyDbContext : DbContext
    {
        public MyDbContext(DbContextOptions<MyDbContext> options) : base(options) { }

        public DbSet<Bliblioteca> Bliblioteca_de_materiais { get; set; }
        public DbSet<MaterialSolidWorks> Banco_de_dados { get; set; }
        public DbSet<SubMaterialSolidWorks> Sub_banco { get; set; }
        public DbSet<Materials> Materiais { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // Chaves primárias explícitas
            modelBuilder.Entity<Bliblioteca>().HasKey(b => b.id_lib);
            modelBuilder.Entity<MaterialSolidWorks>().HasKey(m => m.id_bank);
            modelBuilder.Entity<SubMaterialSolidWorks>().HasKey(s => s.id_sub);
            modelBuilder.Entity<Materials>().HasKey(m => m.id_material);

            // Relacionamento Bliblioteca -> MaterialSolidWorks
            modelBuilder.Entity<MaterialSolidWorks>()
                .HasOne(msw => msw.Bliblioteca)
                .WithMany(b => b.MateriaisSolidWorks)
                .HasForeignKey(msw => msw.IdBliblioteca);

            // Relacionamento MaterialSolidWorks -> SubMaterialSolidWorks
            modelBuilder.Entity<SubMaterialSolidWorks>()
                .HasOne(smsw => smsw.MaterialSolidWorks)
                .WithMany(msw => msw.SubMateriaisSolidWorks)
                .HasForeignKey(smsw => smsw.IdMaterialSolidWorks);

            // Relacionamento SubMaterialSolidWorks -> Materials
            modelBuilder.Entity<Materials>()
                .HasOne(m => m.SubMaterialSolidWorks)
                .WithMany(smsw => smsw.Materiais)
                .HasForeignKey(m => m.IdSubMaterialSolidWorks);
        }
    }

    [Table("Bliblioteca_de_materiais", Schema = "solidworks_data")] // <-- Adicione este atributo
    public class Bliblioteca
    {
        [Key]
        public Guid id_lib { get; set; }
        public string name { get; set; }
        [ValidateNever]
        public ICollection<MaterialSolidWorks> MateriaisSolidWorks { get; set; }
    }

    [Table("Banco_de_dados", Schema = "solidworks_data")] // <-- Adicione este atributo
    public class MaterialSolidWorks
    {
        [Key]
        public Guid id_bank { get; set; }
        public Guid IdBliblioteca { get; set; }
        public string name { get; set; }
        public Bliblioteca Bliblioteca { get; set; }
        public ICollection<SubMaterialSolidWorks> SubMateriaisSolidWorks { get; set; }
    }

    [Table("Sub_banco", Schema = "solidworks_data")] // <-- Adicione este atributo
    public class SubMaterialSolidWorks
    {
        [Key]
        public Guid id_sub { get; set; }
        public Guid IdMaterialSolidWorks { get; set; }
        public string name { get; set; }
        public MaterialSolidWorks MaterialSolidWorks { get; set; }
        public ICollection<Materials> Materiais { get; set; }
    }

    [Table("Materiais", Schema = "solidworks_data")] // <-- Adicione este atributo
    public class Materials
    {
        [Key]
        public Guid id_material { get; set; }
        public Guid IdSubMaterialSolidWorks { get; set; }
        public int mat_id { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string env_data { get; set; }
        public string app_data { get; set; }
        public string name_reduz { get; set; }
        public string angule { get; set; }
        public string escale { get; set; }
        public string tipo_selec { get; set; }
        public string patch_esp { get; set; }
        public string patch_esp_name { get; set; }
        public string patch_band { get; set; }
        public string patch_band_name { get; set; }
        public string patch_calc { get; set; }
        public string patch_calc_name { get; set; }
        public SubMaterialSolidWorks SubMaterialSolidWorks { get; set; }
    }
}