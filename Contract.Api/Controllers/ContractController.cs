using Microsoft.AspNetCore.Mvc;
using Contract.Api.Models;
using Contract.Api.Services;

namespace Contract.Api.Controllers;

[ApiController]
[Route("contracts")]
public class ContractController : ControllerBase
{
    private readonly AgreementDocumentService _service;
    private readonly IWebHostEnvironment _env;

    public ContractController(AgreementDocumentService service, IWebHostEnvironment env)
    {
        _service = service;
        _env = env;
    }

    [HttpPost("create")]
    public IActionResult Create([FromBody] AgreementRequest model)
    {
        if (model?.Items == null || !model.Items.Any())
            return BadRequest("Items list is empty");

        string outputDir = Path.Combine(_env.ContentRootPath, "generated");

        var path = _service.GenerateDocx(model, outputDir);

        var bytes = System.IO.File.ReadAllBytes(path);
        return File(bytes, "application/pdf", Path.GetFileName(path));
    }
}
