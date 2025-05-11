using Dekauto.Import.Service.Domain.Entities.Adapters;
using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Interfaces;
using HotChocolate;
using Microsoft.AspNetCore.Authorization;


namespace Dekauto.Import.Service.Domain.Services
{
    public class Mutation
    {
        [Authorize]
        public async Task<List<Student>> ImportStudents(
        string ld,       // Base64
        string contract,
        string journal,
        [Service] IImportService importService)
        {
            // Конвертируем GraphQL IFile в IFormFile
            var ldFile = ConvertBase64ToFormFile(ld, "ld.xlsx");
            var contractFile = ConvertBase64ToFormFile(contract, "contract.xlsx");
            var journalFile = ConvertBase64ToFormFile(journal, "journal.xlsx");

            // Последовательная обработка файлов
            var studentsLD = (await importService.GetStudentsLD(ldFile)).ToList();
            var studentsContract = (await importService.GetStudentsContract(contractFile, studentsLD)).ToList();
            var studentsJournal = (await importService.GetStudentsJournal(journalFile, studentsContract)).ToList();

            return studentsJournal;
        }

        private IFormFile ConvertBase64ToFormFile(string base64, string fileName)
        {
            var bytes = Convert.FromBase64String(base64);
            var stream = new MemoryStream(bytes);
            return new FormFile(stream, 0, bytes.Length, fileName, fileName)
            {
                Headers = new HeaderDictionary(),
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            };
        }
    }
}
