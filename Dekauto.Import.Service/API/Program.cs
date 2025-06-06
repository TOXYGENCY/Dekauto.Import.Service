using System.Text;
using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Interfaces;
using Dekauto.Import.Service.Domain.Services;
using Dekauto.Import.Service.Domain.Services.Metric;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.OpenApi.Models;
using Prometheus;
using Serilog;
using Serilog.Events;
using Serilog.Formatting.Compact;
using Serilog.Sinks.Grafana.Loki;

try
{
    var builder = WebApplication.CreateBuilder(args);

    // Add services to the container.

    builder.Services.AddControllers();
    // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
    builder.Services.AddEndpointsApiExplorer();

    //Serilog
    Log.Logger = new LoggerConfiguration()
        .MinimumLevel.Information()
        .MinimumLevel.Override("Microsoft", LogEventLevel.Warning)
        .MinimumLevel.Override("Microsoft.AspNetCore", LogEventLevel.Warning)
        .MinimumLevel.Override("System", LogEventLevel.Warning)
        .Enrich.FromLogContext()
        .Enrich.WithMachineName()
        .Enrich.WithEnvironmentUserName()
        .WriteTo.Console(
            new CompactJsonFormatter()
            //outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] [{SourceContext}] {Message:lj}{NewLine}{Exception}"
            )
        .WriteTo.GrafanaLoki(
                "http://loki:3100",
                labels: new List<LokiLabel>
                {
                    new LokiLabel { Key = "app", Value = "dekauto-import" },
                    new LokiLabel { Key = "app", Value = "dekauto-full" }
                })
        .WriteTo.File("logs/Dekauto-Import-.log",
            outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] [{SourceContext}] {Message:lj}{NewLine}{Exception}",
            rollingInterval: RollingInterval.Day,
            rollOnFileSizeLimit: true,
            fileSizeLimitBytes: 10_485_760,
            retainedFileCountLimit: 31,
            encoding: Encoding.UTF8)
        .CreateLogger();

    builder.Host.UseSerilog();

    builder.Services.AddSwaggerGen(c =>
    {
        c.SwaggerDoc("v1", new OpenApiInfo { Title = "Import Service", Version = "v1" });

        c.AddSecurityDefinition("Basic", new OpenApiSecurityScheme
        {
            Name = "Authorization",
            Type = SecuritySchemeType.Http,
            Scheme = "Basic",
            In = ParameterLocation.Header,
            Description = "Basic Authorization header using the Bearer scheme."
        });

        c.AddSecurityRequirement(new OpenApiSecurityRequirement
    {
        {
            new OpenApiSecurityScheme
            {
                Reference = new OpenApiReference
                {
                    Type = ReferenceType.SecurityScheme,
                    Id = "Basic"
                }
            },
            new string[] {}
        }
    });
    });
    builder.Services.AddTransient<IImportService, ImportsService>();
    builder.Services.AddSingleton<IRequestMetricsService, RequestMetricsService>();
    builder.Services.AddScoped<Mutation>();
    // Âêëþ÷àåì ìåæñåðâèñíóþ àâòîðèçàöèþ ïî êîíôèãó
    if (Boolean.Parse(builder.Configuration["UseEndpointAuth"] ?? "true"))
    {
        builder.Services
        .AddAuthentication("Basic")
        .AddScheme<AuthenticationSchemeOptions, BasicAuthenticationHandler>(
            "Basic",
            options => { });

        // Îáùàÿ ïîëèòèêà (ìîæíî èñïîëüçîâàòü è äëÿ GraphQL è äëÿ îáû÷íûõ endpoints)
        builder.Services.AddAuthorization(options =>
        {
            options.DefaultPolicy = new AuthorizationPolicyBuilder("Basic")
                .RequireAuthenticatedUser()
                .Build();
        });
    }
    else
    {
        // Çàãëóøêà ïîëèòèê äîñòóïà, åñëè àâòîðèçàöèÿ âûêëþ÷åíà
        builder.Services.AddAuthorizationBuilder()
        .SetDefaultPolicy(new AuthorizationPolicyBuilder()
        .RequireAssertion(_ => true) // Âñåãäà ðàçðåøàåì äîñòóï
        .Build());
    }

    // Âêëþ÷àåì GraphQL ïî êîíôèãó
    if (Boolean.Parse(builder.Configuration["UseGraphQL"] ?? "true"))
    {
        builder.Services
        .AddGraphQLServer()
        .AddQueryType<Query>()
        .AddMutationType<Mutation>()
        .AddType<UploadType>();
    }

    builder.WebHost.ConfigureKestrel(options =>
    {
        options.Limits.MaxRequestBodySize = 524_288_000; // 500 MB
    });

    builder.Services.AddCors(options =>
    {
        options.AddPolicy("AllowAll",
            builder => builder
                .AllowAnyOrigin()
                .AllowAnyMethod()
                .AllowAnyHeader()
                .WithExposedHeaders("WWW-Authenticate"));
    });

    var app = builder.Build();

    app.UseCors("AllowAll");

    if (Boolean.Parse(builder.Configuration["UseGraphQL"] ?? "true"))
    {
        app.MapGraphQL();
        Log.Information("Enabled GraphQL.");
    }

    // Âêëþ÷àåì ìåæñåðâèñíóþ àâòîðèçàöèþ (â òîì ÷èñëå ÷åðåç [Authorize])
    if (Boolean.Parse(app.Configuration["UseEndpointAuth"] ?? "true"))
    {
        app.UseAuthentication();
        app.UseAuthorization();
        Log.Information("Enabled basic authorization.");

        if (Boolean.Parse(builder.Configuration["UseGraphQL"] ?? "true"))
        {
            app.MapGraphQL().RequireAuthorization();
            Log.Information("Enabled GraphQL with authorization.");
        }
    }
    else
    {
        Log.Warning("Disabled authorization.");
    }


    // Configure the HTTP request pipeline.

    // ßâíî óêàçûâàåì ïîðòû (äëÿ Docker)
    app.Urls.Add("http://*:5503");

    if (app.Environment.IsDevelopment())
    {
        Log.Warning("Development version of the application is started. Swagger activation...");
        app.UseSwagger();
        app.UseSwaggerUI();
    }

    // Âêëþ÷àåì https, åñëè óêàçàíî â êîíôèãå
    if (Boolean.Parse(app.Configuration["UseHttps"] ?? "false"))
    {
        app.Urls.Add("https://*:5504");
        app.UseHttpsRedirection();
        Log.Information("Enabled HTTPS.");
    }
    else
    {
        Log.Warning("Disabled HTTPS.");
    }

    app.MapControllers();

    app.MapMetrics();
    app.UseMetricsMiddleware(); // Ìåòðèêè

    app.Run();
}
catch (Exception ex)
{
    Log.Fatal(ex, "An unexpected Fatal error has occurred in the application.");
}
finally
{
    Log.CloseAndFlush();
}