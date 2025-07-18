using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using WebPAIC_;
using System.Text.Json.Serialization; // Certifique-se de que está aqui

var builder = WebApplication.CreateBuilder(args);

builder.Host.UseWindowsService();

// Add services to the container.
builder.Services.AddControllers()
    .AddJsonOptions(options => // <-- Bloco de configuração adicionado
    {
        options.JsonSerializerOptions.ReferenceHandler = ReferenceHandler.Preserve; // Importante para GETs com Include
        options.JsonSerializerOptions.MaxDepth = 256;
    });

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(); // Configuração básica do Swagger

builder.Services.AddDbContext<MyDbContext>(options =>
    options.UseNpgsql(builder.Configuration.GetConnectionString("DefaultConnection")));

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();

app.Run();