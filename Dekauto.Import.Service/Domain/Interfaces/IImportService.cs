using Dekauto.Import.Service.Domain.Entities;

namespace Dekauto.Import.Service.Domain.Interfaces
{
    public interface IImportService
    {
        Task<IEnumerable<Student>> GetStudentsLD(IFormFile ld);
        Task<IEnumerable<Student>> GetStudentsContract(IFormFile contract, List<Student> students);
        Task<IEnumerable<Student>> GetStudentsJournal(IFormFile journal, List<Student> students);

    }
}
