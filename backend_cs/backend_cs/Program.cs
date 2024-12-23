using System.Diagnostics;
using backend_cs.Service;
using DocumentFormat.OpenXml.InkML;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
builder.Services.AddOpenApi();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

// app.UseHttpsRedirection();

// app.Use(async (context, next) =>
// {
//     var logger = context.RequestServices.GetRequiredService<ILogger<Program>>();
//     var before = DateTime.Now;
//     await next(context);
//     var delta = DateTime.Now - before;
//     logger.LogInformation("Ended after {Time}", delta);
// });

app.MapGet("/", () => {
    var stopwatch = Stopwatch.StartNew();
    // var data = ExcelReader.ReadExcel("sample.xlsx");
    var data = ExcelReaderOpenXML.ReadExcel("sample.xlsx");
    stopwatch.Stop();
    Console.Out.WriteLine(stopwatch.Elapsed);
    return data;
});

app.Run();
