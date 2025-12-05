using Contract.Api.Models;
using Contract.Api.Services;


var builder = WebApplication.CreateBuilder(args);

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var env = builder.Environment;

var templatesDir = Path.Combine(env.ContentRootPath, "templates");
var generatedDir = Path.Combine(env.ContentRootPath, "generated");
var templatePath = Path.Combine(templatesDir, "dogovor_spec_pagebreak.docx");

builder.Services.AddSingleton(new AgreementDocumentService(templatePath));

var sofficePath = @"C:\Program Files\LibreOffice\program\soffice.exe";
builder.Services.AddSingleton(new LibreOfficeConverter(sofficePath));
builder.Services.AddSingleton<AgreementDocumentService>(sp =>
    new AgreementDocumentService("templates/agreement_template.docx"));

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

// === ИМЕННО ВОТ ЭТО — наш главный endpoint ===
app.MapPost("/agreements/generate", async (
    AgreementRequest request,
    AgreementDocumentService docService,
    LibreOfficeConverter converter) =>
{
    var docxPath = docService.GenerateDocx(request, generatedDir);
    var pdfPath = converter.ConvertToPdf(docxPath, generatedDir);

    var bytes = await File.ReadAllBytesAsync(pdfPath);
    var fileName = Path.GetFileName(pdfPath);

    return Results.File(bytes, "application/pdf", fileName);
});

app.Run();
