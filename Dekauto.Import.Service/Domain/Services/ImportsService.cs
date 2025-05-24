using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Interfaces;
using OfficeOpenXml;
using System.Diagnostics.Contracts;
using System.Text.RegularExpressions;

namespace Dekauto.Import.Service.Domain.Services
{
    public class ImportsService : IImportService
    {
        private IConfiguration configuration;
        public ImportsService(IConfiguration configuration)
        {
            this.configuration = configuration;
        }
        public async Task<IEnumerable<Student>> GetStudentsContract(IFormFile contract, List<Student> students)
        {
            using (var stream = new MemoryStream()) 
            {
                await contract.CopyToAsync(stream);
                using (var packege = new ExcelPackage(stream)) 
                {
                    var worksheet = packege.Workbook.Worksheets[0] ?? throw new InvalidOperationException("Загруженный файл не содержит листов");

                    var columnCount = worksheet.Dimension.Columns;
                    var rowCount = worksheet.Dimension.Rows;

                    var headers = new List<string>();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        headers.Add(worksheet.Cells[1, col].Text);
                    }
                    if (students.Count == 0 || students == null) throw new ArgumentNullException("Студенты отсутствуют в таблице личных дел");
                    foreach (var student in students) 
                    {
                        string fio = $"{student.Surname}{student.Name}{student.Patronymic}".ToLower();
                        for (int row = 2; row <= rowCount; row++) 
                        {
                            bool isCurrentStudent = false;
                            string enrollementOrderDatePattern = @"\d{2}\.\d{2}\.\d{4}";
                            string enrollementOrderNumPattern = @"\d*\-\d*\/\d*";
                            for (int col = 1; col <= columnCount; col++)
                            {
                                var header = headers[col - 1];
                                var cellValue = worksheet.Cells[row, col].Value ?? "";

                                if (header.ToLower() == "фио обучающегося" || header.ToLower() == "фио студента")
                                {
                                    string cellfio = cellValue.ToString().ToLower().Replace(" ", "");
                                    if (cellfio == fio)
                                    {
                                        isCurrentStudent = true;
                                        break; 
                                    }
                                }
                            }
                            for (int col = 1; col <= columnCount; col++) 
                            {
                                var header = headers[col - 1];
                                var cellValue = worksheet.Cells[row, col].Value ?? "";
                                switch (header.ToLower()) 
                                {
                                    case "дата":
                                        if (isCurrentStudent == true) 
                                        {
                                            DateTime date = (DateTime)cellValue;
                                            student.EducationRelationDate = DateOnly.FromDateTime(date);
                                            student.EducationStartYear = short.Parse(student.EducationRelationDate.Value.Year.ToString());
                                            student.EducationFinishYear = (short)(student.EducationStartYear + student.EducationTime);
                                        }
                                        break;
                                    case "№ договора": case "номер договора":
                                        if (isCurrentStudent == true)
                                        {
                                            student.EducationRelationNum = cellValue.ToString();
                                        }
                                        break;
                                    case "№ приказа о зачислении": case "номер приказа о зачислении":
                                        if (isCurrentStudent == true)
                                        {
                                            DateTime date = DateTime.Parse(Regex.Match(cellValue.ToString(), enrollementOrderDatePattern).ToString());
                                            student.EnrollementOrderDate = DateOnly.FromDateTime(date);
                                            student.EnrollementOrderNum = Regex.Match(cellValue.ToString(), enrollementOrderNumPattern).ToString();
                                        }
                                        break;
                                }
                            }
                        }
                    }
                }
            }
            return students;
        }

        public async Task<IEnumerable<Student>> GetStudentsJournal(IFormFile journal, List<Student> students)
        {
            using (var stream = new MemoryStream())
            {
                await journal.CopyToAsync(stream);
                using (var packege = new ExcelPackage(stream))
                {
                    var worksheet = packege.Workbook.Worksheets[0] ?? throw new InvalidOperationException("Загруженный файл не содержит листов");

                    var columnCount = worksheet.Dimension.Columns;
                    var rowCount = worksheet.Dimension.Rows;

                    var headers = new List<string>();
                    for (int col = 1; col <= columnCount; col++)
                    {
                        headers.Add(worksheet.Cells[1, col].Text);
                    }

                    foreach (var student in students)
                    {
                        string fio = $"{student.Surname}{student.Name}{student.Patronymic}".ToLower();
                        
                        for (int row = 4; row <= rowCount; row++)
                        {
                            bool isCurrentStudent = false;
                            for (int col = 1; col <= columnCount; col++)
                            {
                                var header = headers[col - 1];
                                var cellValue = worksheet.Cells[row, col].Value ?? "";

                                if (header.ToLower() == "фио студента" || header.ToLower() == "фио обучающегося")
                                {
                                    string cellfio = cellValue.ToString().ToLower().Replace(" ", "");
                                    if (cellfio == fio)
                                    {
                                        isCurrentStudent = true;
                                        break;
                                    }
                                }
                            }
                            for (int col = 1; col <= columnCount; col++)
                            {
                                var header = headers[col - 1];
                                var cellValue = worksheet.Cells[row, col].Value ?? "";
                                switch (header.ToLower()) 
                                {
                                    case "№ студ.билета и зачетной книжки": 
                                    case "№ студ.билета":
                                    case "№ зачетной книжки":
                                    case "№ зачетки":
                                    case "номер зачетки":
                                        if (isCurrentStudent)
                                            student.GradeBook = cellValue.ToString();
                                        break;
                                    case "№ группы":
                                    case "номер группы":
                                        if (isCurrentStudent)
                                            student.GroupName = cellValue.ToString();
                                        break;
                                }
                            }
                        }
                    }
                }
            }
            return students;
        }

        public async Task<IEnumerable<Student>> GetStudentsLD(IFormFile ld)
        {
            var students = new List<Student>();
            using (var stream = new MemoryStream()) 
            {
                await ld.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream)) 
                {
                    var worksheet = package.Workbook.Worksheets[0] ?? throw new InvalidOperationException("Загруженный файл не содержит листов");

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

                        // Фиксированные данные
                        student.Education = "высшее"; // других видов нет
                        student.EducationForm = "очная"; // других форм нет
                        student.Faculty = "Информационных технологий"; // Т.к. делаем прогу для нашего факультета, факультет будет такой. Пока взять альтернативные варианты неоткуда
                        student.Course = "Прикладная информатика в психологии"; // нужна таблица с соответствиями специализаций с направлениями подготовки
                        student.EducationBase = "бюджетная (ФБ)";// других основ нет
                        student.EducationRelationForm = "договор об образовании на обучение";// платных услуг нет
                        student.EducationTime = 4; // Пока ставль 4, т.к. неоткуда брать данные
                        student.LivingInDormitory = false; // Общежитие по умолчанию нет
                        student.MilitaryService = false; // Служба в армии по умолчанию нет
                        student.MaritalStatus = false; // Отношения по умолчанию отсутствуют


                        for (int col = 1; col <= columnCount; col++) 
                        {
                            var header = headers[col-1];
                            var cellValue = worksheet.Cells[row, col].Value ?? "";

                            string indexPattern = @"\b\d{6}\b";
                            string addressTypePattern = @"\b(?:\w+)\s+([г|с|х|д|п])\b(?:\,)";
                            string cityPattern = @"\b(\w+)\s+(?:[г|с|х|д|п]\,)"; // Пока что тригер работает только на города и села, для увеличения вариантов нужно идти в деканат
                            string housePattern = @"\b(?:д\.)\s+(\w+)\b(?:,)?";
                            string streetPattern = @",\s*([^,]*?)\s*,\s+д\.";
                            string housingTypePattern = @"\b([к|стр])\.\w*?"; // Нужно уточнение, как обозначается строение
                            string housingPattern = @"[к|стр]\.\s+(\w+)";
                            string apartementPattern = @"кв\.\s+(\w+)";
                            string numConcursPattern = @"(\d{2}\.\d{2}\.\d{2})";

                            switch (header.ToLower()) 
                            {
                                case "фио":
                                    string pattern = @"\S+";
                                    MatchCollection fio = Regex.Matches(cellValue.ToString().ToLower(), pattern);
                                    string name = $"{fio[1].ToString().Substring(0,1).ToUpper()}{fio[1].ToString().Substring(1)}";
                                    student.Name = name;
                                    string surname = $"{fio[0].ToString().Substring(0, 1).ToUpper()}{fio[0].ToString().Substring(1)}";
                                    student.Surname = surname;
                                    if (fio.Count > 2) 
                                    {
                                        string patronymic = $"{fio[2].ToString().Substring(0, 1).ToUpper()}{fio[2].ToString().Substring(1)}";
                                        student.Patronymic = patronymic;
                                    }
                                    else student.Patronymic = "";
                                    break;
                                case "пол":
                                    if (cellValue.ToString().ToLower() == "мужской") student.Gender = true; 
                                    else student.Gender = false;
                                    break;
                                case "дата рождения":
                                    DateTime birthdayDate = (DateTime)cellValue;
                                    student.BirthdayDate = DateOnly.FromDateTime(birthdayDate);
                                    break;
                                case "место рождения":
                                    student.BirthdayPlace = cellValue.ToString();
                                    break;
                                case "телефон":
                                    student.PhoneNumber = cellValue.ToString();
                                    break;
                                case "e-mail":
                                case "почта":
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
                                case "код подразделения документа удостоверяющего личность":
                                case "код подразделения":// Здесь нужны уточнения, как называется поле
                                    student.PassportIssuanceCode = cellValue.ToString();
                                    break;
                                case "дата выдачи паспорта": // Здесь нужны уточнения, как называется поле и существует ли вообще
                                    DateTime passportDate = (DateTime)cellValue;
                                    student.PassportIssuanceDate = DateOnly.FromDateTime(passportDate);
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
                                case "сумма баллов за инд.дост.":
                                    student.BonusScores = short.Parse(cellValue.ToString());
                                    break;
                                case "адрес по прописке":
                                    student.AddressRegistrationIndex = Regex.Match(cellValue.ToString(), indexPattern).ToString();
                                    student.AddressRegistrationCity = Regex.Match(cellValue.ToString(), cityPattern).Groups[1].ToString();
                                    switch (Regex.Match(cellValue.ToString(), addressTypePattern).Groups[1].ToString())
                                    {
                                        case "г":
                                            student.AddressRegistrationType = "город:";
                                            break;
                                        case "с":
                                            student.AddressRegistrationType = "село:";
                                            break;
                                        case "х":
                                            student.AddressRegistrationType = "хутор:";
                                            break;
                                        case "д":
                                            student.AddressRegistrationType = "деревня:";
                                            break;
                                        case "п":
                                            student.AddressRegistrationType = "посёлок:";
                                            break;
                                    }
                                    student.AddressRegistrationHouse = Regex.Match(cellValue.ToString(), housePattern).Groups[1].ToString();
                                    student.AddressRegistrationStreet = Regex.Match(cellValue.ToString(), streetPattern).Groups[1].ToString().Trim();
                                    if (Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim() == "к") 
                                        student.AddressRegistrationHousingType = "корпус";
                                    else
                                        if (Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim() == "стр") 
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
                                            student.AddressResidentialType = "город:";
                                            break;
                                        case "с":
                                            student.AddressResidentialType = "село:";
                                            break;
                                        case "х":
                                            student.AddressResidentialType = "хутор:";
                                            break;
                                        case "д":
                                            student.AddressResidentialType = "деревня:";
                                            break;
                                        case "п":
                                            student.AddressResidentialType = "посёлок:";
                                            break;
                                    }
                                    student.AddressResidentialHouse = Regex.Match(cellValue.ToString(), housePattern).Groups[1].ToString();
                                    student.AddressResidentialStreet = Regex.Match(cellValue.ToString(), streetPattern).Groups[1].ToString().Trim();
                                    if (Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim() == "к")
                                        student.AddressResidentialHousingType = "корпус";
                                    else
                                        if (Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim() == "стр")
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
                                    DateTime eduReceivedDate = (DateTime)cellValue;
                                    student.EducationReceivedDate = DateOnly.FromDateTime(eduReceivedDate);
                                    break;
                                case "год завершения":
                                    student.EducationReceivedEndYear = short.Parse(cellValue.ToString());
                                    break;
                                case "тип документа об образовании":
                                    if (cellValue.ToString().ToLower() == "аттестат")
                                        student.EducationReceived = "среднее общее образование";
                                    else if (cellValue.ToString().ToLower() == "диплом")
                                        student.EducationReceived = "среднее профессиональное образование";
                                    break;
                                case "образовательное учреждение":
                                    student.OOName = cellValue.ToString();
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
