using Microsoft.AspNetCore.Mvc;
using Contract.Api.Models;
using Contract.Api.Services;   // ⬅ вот это важно

namespace Contract.Api.Controllers;

[ApiController]
[Route("contracts")]
public class ContractController : ControllerBase
{
    private readonly AgreementDocumentService _service;

    public ContractController(AgreementDocumentService service)
    {
        _service = service;
    }

    [HttpPost("create")]
    public IActionResult Create([FromBody] AgreementRequest model)
    {
        if (model?.Items == null || !model.Items.Any())
            return BadRequest("Items list is empty");

        var path = _service.Generate(model);

        var bytes = System.IO.File.ReadAllBytes(path);
        return File(bytes, "application/pdf", Path.GetFileName(path));
    }
}
