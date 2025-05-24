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

// явно указываем порты (дл€ Docker)
app.Urls.Add("http://*:5503");

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

// ¬ключаем https, если указано в конфиге
if (Boolean.Parse(app.Configuration["UseHttps"] ?? "false"))
{
    app.Urls.Add("https://*:5504");
    app.UseHttpsRedirection();
    //Log.Information("Enabled HTTPS.");
}
else
{
    //Log.Warning("Disabled HTTPS.");
}

if (Boolean.Parse(app.Configuration["UseEndpointAuth"] ?? "true"))
{
    // ¬ключаем межсервисную авторизацию (в том числе через [Authorize])
    //Log.Information("Enabled basic authorization.");
    app.UseAuthorization();
}
else
{
    //Log.Warning("Disabled basic authorization.");
}

app.MapControllers();

app.Run();