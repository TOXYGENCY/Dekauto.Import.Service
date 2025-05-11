using Dekauto.Import.Service.Domain.Entities.Adapters;
using Dekauto.Import.Service.Domain.Interfaces;
using Dekauto.Import.Service.Domain.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Dekauto.Import.Service.API.Controllers
{
    //Раз мы здесь используем GraphQL, по факту контроллер здесь не нужен
    [Route("api/imports")]
    [ApiController]
    [Authorize]
    public class ImportController : ControllerBase
    {
        private readonly IImportService _importService;
        public ImportController(IImportService importService) 
        {
            _importService = importService ?? throw new ArgumentNullException(nameof(importService));
        }

        [HttpPost]
        [Route("students")]
        public async Task<IActionResult> ImportStudents([FromForm] ImportFilesAdapter files) 
        {
            try
            {
                var ld = files.ld;
                var contract = files.contract;
                var journal = files.journal;

                if (ld == null || ld.Length == 0 ||
                    contract == null || contract.Length == 0 ||
                    journal == null || journal.Length == 0) throw new ArgumentNullException("Файл не найден");
                if (System.IO.Path.GetExtension(ld.FileName) != ".xlsx" ||
                    System.IO.Path.GetExtension(contract.FileName) != ".xlsx" || 
                    System.IO.Path.GetExtension(journal.FileName) != ".xlsx") throw new FileLoadException(
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
