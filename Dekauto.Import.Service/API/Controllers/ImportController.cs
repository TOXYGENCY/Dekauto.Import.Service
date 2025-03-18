using Dekauto.Import.Service.Domain.Interfaces;
using Dekauto.Import.Service.Domain.Services;
using Microsoft.AspNetCore.Mvc;

namespace Dekauto.Import.Service.API.Controllers
{
    [Route("api/imports")]
    [ApiController]
    public class ImportController : ControllerBase
    {
        private readonly IImportService _importService;
        public ImportController(IImportService importService) 
        {
            _importService = importService ?? throw new ArgumentNullException(nameof(importService));
        }

        [HttpPost]
        [Route("students")]
        public async Task<IActionResult> ImportStudents(IFormFile ld, IFormFile contract, IFormFile journal) 
        {
            try
            {
                if (ld == null || ld.Length == 0 ||
                    contract == null || contract.Length == 0 ||
                    journal == null || journal.Length == 0) throw new ArgumentNullException("Файл не найден");
                if (Path.GetExtension(ld.FileName) != ".xlsx" || 
                    Path.GetExtension(contract.FileName) != ".xlsx" || 
                    Path.GetExtension(journal.FileName) != ".xlsx") throw new FileLoadException(
                    "Неподдерживаемый формат файла. Пожалуйста, загрузите файл в формате .xlsx");
                var studentsLD = await _importService.GetStudentsLD(ld);
                var studentsOrder = await _importService.GetStudentsContract(contract, (List<Domain.Entities.Student>)studentsLD);
                var students = await _importService.GetStudentsJournal(journal, (List<Domain.Entities.Student>)studentsOrder);

                return Ok(students);
            }
            catch (ArgumentNullException ex)
            {
                return NotFound(ex.Message);
            }
            catch (FileLoadException ex) 
            {
                return BadRequest(ex.Message);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception) 
            {
                return StatusCode(500, "Ошибка на стороне сервера, обратитесь к администратору");
            }
        }
    }
}
