using Dekauto.Import.Service.Domain.Interfaces;
using Dekauto.Import.Service.Domain.Services;
using Microsoft.AspNetCore.Mvc;

namespace Dekauto.Import.Service.API.Controllers
{
    [Route("api/imorts")]
    [ApiController]
    public class ImportController : ControllerBase
    {
        private readonly IImportService _importService;
        public ImportController(IImportService importService) 
        {
            _importService = importService;
        }

        [HttpPost]
        public async Task<IActionResult> ImportLD(IFormFile file) 
        {
            try
            {
                if (file == null || file.Length == 0) throw new ArgumentNullException("Файл не найден");
                if (Path.GetExtension(file.FileName) != ".xlsx") throw new ArgumentNullException(
                    "Неподдерживаемый формат файла. Пожалуйста, загрузите файл в формате .xlsx");
                var students = _importService.GetStudentsLD(file);
                return Ok(students);
            }
            catch (Exception) 
            {
                return StatusCode(500, "Ошибка на стороне сервера, обратитесь к администратору");
            }
        }
    }
}
