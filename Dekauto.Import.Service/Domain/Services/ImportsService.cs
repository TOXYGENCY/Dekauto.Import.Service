using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Interfaces;
using OfficeOpenXml;
using System.Text.RegularExpressions;

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
                            var cellValue = worksheet.Cells[row, col].Value ?? "";

                            switch (header.ToLower()) 
                            {
                                case "фио":
                                    string pattern = @"\S+";
                                    MatchCollection fio = Regex.Matches(cellValue.ToString().ToLower(), pattern);
                                    student.Name = fio[0].Value;
                                    student.Surname = fio[1].Value;
                                    if (fio.Count > 2) student.Pathronymic = fio[2].Value;
                                    else student.Pathronymic = "";
                                    break;
                                case "пол":
                                    if (cellValue.ToString().ToLower() == "мужской") student.Gender = true; 
                                    else student.Gender = false;
                                    break;
                                case "дата рождения":
                                    student.BirthdayDate = (DateTime?)cellValue;
                                    break;
                                case "место рождения":
                                    student.BirthdayPlace = cellValue.ToString();
                                    break;
                                case "телефон":
                                    student.PhoneNumber = cellValue.ToString();
                                    break;
                                case "e-mail":
                                    student.Email = cellValue.ToString();
                                    break;
                                case "серия документа удостоверяющего личность":
                                    student.PassportSerial = cellValue.ToString();
                                    break;
                                case "номер документа удостоверяющего личность":
                                    student.PassportNumber = cellValue.ToString();
                                    break;
                                case "овддокумента удостоверяющего личность":
                                    student.PassportIssuancePlace = cellValue.ToString();
                                    break;
                                case "код подразделения": // Здесь нужны уточнения, как называется поле
                                    student.PassportIssuanceCode = cellValue.ToString();
                                    break;
                                case "дата выдачи паспорта": // Здесь нужны уточнения, как называется поле и существует ли вообще
                                    student.PassportIssuanceDate = (DateTime?)cellValue;
                                    break;
                                case "гражданство":
                                    student.Citizenship = cellValue.ToString();
                                    break;
                                case "предмет1":
                                    student.GiaExam1Score = short.Parse(cellValue.ToString());
                                    break;
                                case "предмет2":
                                    student.GiaExam2Score = short.Parse(cellValue.ToString());
                                    break;
                                case "предмет3":
                                    student.GiaExam3Score = short.Parse(cellValue.ToString());
                                    break;
                                case "сумма баллов за инд.дост.(конкурсные)":
                                    student.BonusScores = short.Parse(cellValue.ToString());
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
