using Microsoft.AspNetCore.Mvc;
using Contract.Api.Models;
using Contract.Api.Services;

namespace Contract.Api.Controllers
{
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
        public IActionResult Create([FromBody] AgreementModel model)
        {
            if (model == null || model.Items == null || model.Items.Count == 0)
                return BadRequest("Invalid model");

            // Генерация PDF
            var filePath = _service.Generate(model); // ← твой сервис уже готов
            var fileBytes = System.IO.File.ReadAllBytes(filePath);

            return File(fileBytes, "application/pdf", Path.GetFileName(filePath));
        }
    }
}