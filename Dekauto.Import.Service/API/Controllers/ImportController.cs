using Dekauto.Import.Service.Domain.Interfaces;
using Dekauto.Import.Service.Domain.Services;
using Microsoft.AspNetCore.Mvc;

namespace Dekauto.Import.Service.API.Controllers
{
    [Route("api/import")]
    [ApiController]
    public class ImportController : ControllerBase
    {
        private readonly IImportService _importService;
        public ImportController(IImportService importService) 
        {
            _importService = importService;
        }

        [HttpPost]
        [Route("students")]
        public async Task<IActionResult> ImportFiles(IFormFile ld, IFormFile? contract) 
        {
            try
            {
                if (ld == null || ld.Length == 0) throw new ArgumentNullException("Файл не найден");
                if (Path.GetExtension(ld.FileName) != ".xlsx") throw new ArgumentNullException(
                    "Неподдерживаемый формат файла. Пожалуйста, загрузите файл в формате .xlsx");
                var students = _importService.GetStudentsLD(ld);
                return Ok(students);
            }
            catch (Exception) 
            {
                return StatusCode(500, "Ошибка на стороне сервера, обратитесь к администратору");
            }
        }
    }
}
