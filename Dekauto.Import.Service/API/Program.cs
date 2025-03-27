using Dekauto.Import.Service.Domain.Interfaces;
using Dekauto.Import.Service.Domain.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddTransient<IImportService, ImportsService>();

var app = builder.Build();

// Configure the HTTP request pipeline.

// Явно указываем порты (для Docker)
app.Urls.Add("http://*:5503");

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}
else
{
    app.Urls.Add("https://*:5504");
    app.UseHttpsRedirection(); // без https редиректа в dev-версии
}

app.UseAuthorization();

app.MapControllers();

app.Run();