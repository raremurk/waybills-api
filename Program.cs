using WaybillsAPI.Context;
using WaybillsAPI.Interfaces;
using WaybillsAPI.Mappings;
using WaybillsAPI.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddDbContext<WaybillsContext>();
builder.Services.AddAutoMapper(typeof(MappingProfile));
builder.Services.AddControllers();
builder.Services.AddHttpClient();

builder.Services.AddSingleton<ISalaryPeriodService, SalaryPeriodService>();
builder.Services.AddSingleton<IExcelWriter, ExcelWriter>();
builder.Services.AddSingleton<IOmnicommFuelService, OmnicommFuelService>();

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddCors();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.MapControllers();

app.UseCors(builder => builder.AllowAnyOrigin()
                              .AllowAnyMethod()
                              .AllowAnyHeader());

app.Run();

