using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Interfaces;
using OfficeOpenXml;

namespace Dekauto.Import.Service.Domain.Services
{
    public class ImportsService : IImportService
    {
        public async Task<IEnumerable<Student>> GetStudentsLD(IFormFile LD)
        {
            var students = new List<Student>();
            using (var stream = new MemoryStream()) 
            {
                await LD.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream)) 
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    var columnCount = worksheet.Dimension.Columns;
                    var rowCount = worksheet.Dimension.Rows;

                    var headers = new List<string>();
                    for (int col = 1; col <= columnCount; col++) 
                    {
                        headers.Add(worksheet.Cells[1, col].Text);
                    }

                    for (int row = 2; row <= rowCount; row++) 
                    {
                        var student = new Student();
                        
                        for (int col = 1; col <= columnCount; col++) 
                        {
                            var header = headers[col-1];
                            var cellValue = worksheet.Cells[row, col].Text.ToLower();

                            switch (header.ToLower()) 
                            {
                                case "фио":
                                    student.Name = cellValue;
                                    break;
                            }
                        }
                        students.Add(student);
                    }
                }
                
            } 
            return students;
        }
    }
}
