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
                        string courseOfTraining = string.Empty;

                        for (int col = 1; col <= columnCount; col++) 
                        {
                            var header = headers[col-1];
                            var cellValue = worksheet.Cells[row, col].Value ?? "";

                            string indexPattern = @"\b\d{6}\b";
                            string addressTypePattern = @"\b(?:\w+)\s+([г|с])\b(?:\,)";
                            string cityPattern = @"\b(\w+)\s+(?:[г|с]\,)"; // Пока что тригер работает только на города и села, для увеличения вариантов нужно идти в деканат
                            string housePattern = @"\b(?:д\.)\s+(\w+)\b(?:,)?";
                            string streetPattern = @",\s*([^,]*?)\s*,\s+д\.";
                            string housingTypePattern = @"\b([к|с])\.\w*?"; // Нужно уточнение, как обозначается строение
                            string housingPattern = @"[к|с]\.\s+(\w+)";
                            string apartementPattern = @"кв\.\s+(\w+)";
                            string numConcursPattern = @"(\d{2}\.\d{2}\.\d{2})";
                            

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
                                case "адрес по прописке":
                                    student.AddressRegistrationIndex = Regex.Match(cellValue.ToString(), indexPattern).ToString();
                                    student.AddressRegistrationCity = Regex.Match(cellValue.ToString(), cityPattern).Groups[1].ToString();
                                    switch (Regex.Match(cellValue.ToString(), addressTypePattern).Groups[1].ToString())
                                    {
                                        case "г":
                                            student.AddressRegistrationType = "город";
                                            break;
                                        case "с":
                                            student.AddressRegistrationType = "село";
                                            break;
                                    }
                                    student.AddressRegistrationHouse = Regex.Match(cellValue.ToString(), housePattern).Groups[1].ToString();
                                    student.AddressRegistrationStreet = Regex.Match(cellValue.ToString(), streetPattern).Groups[1].ToString().Trim();
                                    if (Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim() == "к") 
                                        student.AddressRegistrationHousingType = "корпус";
                                    else
                                        if (Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim() == "с") 
                                            student.AddressRegistrationHousingType = "строение";
                                    student.AddressRegistrationHousing = Regex.Match(cellValue.ToString(), housingPattern).Groups[1].ToString().Trim();
                                    student.AddressRegistrationApartment = Regex.Match(cellValue.ToString(), apartementPattern).Groups[1].ToString().Trim();
                                    break;

                                case "адрес проживания":
                                    student.AddressResidentialIndex = Regex.Match(cellValue.ToString(), indexPattern).ToString();
                                    student.AddressResidentialCity = Regex.Match(cellValue.ToString(), cityPattern).Groups[1].ToString();
                                    switch(Regex.Match(cellValue.ToString(), addressTypePattern).Groups[1].ToString()) 
                                    {
                                        case "г":
                                            student.AddressResidentialType = "город";
                                            break;
                                        case "с":
                                            student.AddressResidentialType = "село";
                                            break;
                                    }
                                    student.AddressResidentialHouse = Regex.Match(cellValue.ToString(), housePattern).Groups[1].ToString();
                                    student.AddressResidentialStreet = Regex.Match(cellValue.ToString(), streetPattern).Groups[1].ToString().Trim();
                                    if (Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim() == "к")
                                        student.AddressResidentialHousingType = "корпус";
                                    else
                                        if (Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim() == "с")
                                        student.AddressResidentialHousingType = "строение";
                                    student.AddressResidentialHousing = Regex.Match(cellValue.ToString(), housingPattern).Groups[1].ToString().Trim();
                                    student.AddressResidentialApartment = Regex.Match(cellValue.ToString(), apartementPattern).Groups[1].ToString().Trim();

                                    break;
                                case "конкурсная группа":
                                    courseOfTraining = Regex.Match(cellValue.ToString(), numConcursPattern).Groups[1].ToString().Trim();
                                    break;
                                case "направление\\специальность":
                                    student.CourseOfTraining = $"{courseOfTraining} {cellValue.ToString()}";
                                    break;
                                case "серия документа об образовании":
                                    student.EducationReceivedSerial = cellValue.ToString();
                                    break;
                                case "номер документа об образовании":
                                    student.EducationReceivedNum = cellValue.ToString();
                                    break;
                                case "дата выдачи":
                                    student.EducationReceivedDate = (DateTime?)cellValue;
                                    break;
                                case "год завершения":
                                    student.EducationReceivedEndYear = short.Parse(cellValue.ToString());
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
