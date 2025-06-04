using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Interfaces;
using OfficeOpenXml;
using System.Diagnostics.Contracts;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Dekauto.Import.Service.Domain.Services
{
    public class ImportsService : IImportService
    {
        private IConfiguration configuration;
        private readonly ILogger<ImportsService> logger;
        public ImportsService(IConfiguration configuration, ILogger<ImportsService> logger)
        {
            this.configuration = configuration;
            this.logger = logger;
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

                                logger.LogInformation($"Работа с ячейкой: [{col},{row}]; столбец {header}");

                                switch (header.ToLower()) 
                                {
                                    case "дата":
                                        if (isCurrentStudent == true) 
                                        {
                                            if (cellValue is DateTime excelDate) 
                                            {
                                                student.EducationRelationDate = DateOnly.FromDateTime(excelDate);
                                            } 
                                            else
                                            {
                                                string dateStr = cellValue.ToString().Trim();
                                                if (DateTime.TryParseExact(
                                                    dateStr,
                                                    "dd.MM.yyyy",
                                                    CultureInfo.InvariantCulture,
                                                    DateTimeStyles.None,
                                                    out DateTime parsedDate))
                                                {
                                                    student.EducationRelationDate = DateOnly.FromDateTime(parsedDate);
                                                }
                                                else throw new FormatException($"Не удалось распознать дату: {dateStr}");
                                            }
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
                                            string dateStr = Regex.Match(cellValue.ToString().Trim(), enrollementOrderDatePattern).ToString();

                                            if (DateTime.TryParseExact(
                                                   dateStr,
                                                   "dd.MM.yyyy",
                                                   CultureInfo.InvariantCulture,
                                                   DateTimeStyles.None,
                                                   out DateTime parsedDate))
                                            {
                                                student.EnrollementOrderDate = DateOnly.FromDateTime(parsedDate);
                                            }
                                            else throw new FormatException($"Не удалось распознать дату: {dateStr}");

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

                                if (header.ToLower() == "фио студента" || header.ToLower() == "фио обучающегося" || header.ToLower() == "фио")
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

                                logger.LogInformation($"Работа с ячейкой: [{col},{row}]; столбец {header}");

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
                            string addressTypePattern = @"\b(?:\w+\s+(?<abbr>[гсхдп])\b|(?<abbr>[гсхдп])\.?\s+\w+\b)";
                            string cityPattern = @"\b(?:(?<city>[\w\s]+?)\s+[гсхдп]\,?|(?<type>[гсхдп])\.?\s*(?<city>[\w\s]+?))\b";
                            string housePattern = @"(?:\bдом|д)\.?\s*(\w+)\b,?|\bд(\w+)\b";
                            string streetPattern = @",?\s*([^,]+?)\s*,\s*(?:дом|д)\.";
                            string housingTypePattern = @"(?:^|,)\s*(к(?:\.|орпус)?|стр(?:\.|оение)?)(?=\s*\w|$)(?!\w)";
                            string housingPattern = @"(?:^|,)\s*(?<type>к(?:\.|орпус)?|стр(?:\.|оение)?)\s*(?<number>[\w\-]*\d[\w\-]*)\b";
                            string apartementPattern = @"(?:^|\s)(?:кв\.?|квартира\.?)\s*(\w+)\b";
                            string numConcursPattern = @"(\d{2}\.\d{2}\.\d{2})";
                            string regionPattern = @"(?:^|,)\s*(?<region>[\w\s-]+?)\s*(?<type>обл\.?|кр\.?|край|автономная\s+область|авт\.?\s*обл\.?|АО)\b";

                            logger.LogInformation($"Работа с ячейкой: [{col},{row}]; столбец {header}");


                            switch (header.ToLower())
                            {
                                case "фио":
                                    if (cellValue == "") return students;
                                    break;
                            }

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
                                    if (cellValue is DateTime birDate)
                                    {
                                        student.BirthdayDate = DateOnly.FromDateTime(birDate);
                                    }
                                    else
                                    {
                                        string dateStr = cellValue.ToString().Trim();
                                        if (DateTime.TryParseExact(
                                            dateStr,
                                            "dd.MM.yyyy",
                                            CultureInfo.InvariantCulture,
                                            DateTimeStyles.None,
                                            out DateTime parsedDate))
                                        {
                                            student.BirthdayDate = DateOnly.FromDateTime(parsedDate);
                                        }
                                        else throw new FormatException($"Не удалось распознать дату: {dateStr}");
                                    }
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
                                    if (cellValue is DateTime pasDate)
                                    {
                                        student.PassportIssuanceDate = DateOnly.FromDateTime(pasDate);
                                    }
                                    else
                                    {
                                        string dateStr = cellValue.ToString().Trim();
                                        if (DateTime.TryParseExact(
                                            dateStr,
                                            "dd.MM.yyyy",
                                            CultureInfo.InvariantCulture,
                                            DateTimeStyles.None,
                                            out DateTime parsedDate))
                                        {
                                            student.PassportIssuanceDate = DateOnly.FromDateTime(parsedDate);
                                        }
                                        else throw new FormatException($"Не удалось распознать дату: {dateStr}");
                                    }
                                    break;
                                case "гражданство":
                                    student.Citizenship = cellValue.ToString();
                                    break;
                                case "предмет1":
                                    if (cellValue != "")
                                        student.GiaExam1Score = short.Parse(cellValue.ToString());
                                    break;
                                case "предмет2":
                                    if (cellValue != "")
                                        student.GiaExam2Score = short.Parse(cellValue.ToString());
                                    break;
                                case "предмет3":
                                    if (cellValue != "")
                                        student.GiaExam3Score = short.Parse(cellValue.ToString());
                                    break;
                                case "сумма баллов за инд.дост.(конкурсные)":
                                case "сумма баллов за инд.дост.":
                                    if (cellValue != "")
                                        student.BonusScores = short.Parse(cellValue.ToString());
                                    break;
                                case "адрес по прописке":
                                    if (cellValue.ToString() != "")
                                    {
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
                                            case "х":
                                                student.AddressRegistrationType = "хутор";
                                                break;
                                            case "д":
                                                student.AddressRegistrationType = "деревня";
                                                break;
                                            case "п":
                                                student.AddressRegistrationType = "посёлок";
                                                break;
                                        }
                                        student.AddressRegistrationHouse = Regex.Match(cellValue.ToString(), housePattern).Groups[1].ToString();
                                        student.AddressRegistrationStreet = Regex.Match(cellValue.ToString(), streetPattern).Groups[1].ToString().Trim();
                                        string housingMatch = Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim();
                                        if (housingMatch == "к" || housingMatch == "к." || housingMatch == "корпус" || housingMatch == "корпус.")
                                            student.AddressRegistrationHousingType = "корпус";
                                        else
                                            if (housingMatch == "стр" || housingMatch == "стр." || housingMatch == "строение" || housingMatch == "строение.")
                                            student.AddressRegistrationHousingType = "строение";
                                        student.AddressRegistrationHousing = Regex.Match(cellValue.ToString(), housingPattern).Groups[2].ToString().Trim();
                                        student.AddressRegistrationApartment = Regex.Match(cellValue.ToString(), apartementPattern).Groups[1].ToString().Trim();
                                        student.AddressRegistrationOblKrayAvtobl = Regex.Match(cellValue.ToString(), regionPattern).Groups[1].ToString().Trim();

                                    }
                                    break;

                                case "адрес проживания":
                                    if (cellValue.ToString() != "") { 
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
                                            case "х":
                                                student.AddressResidentialType = "хутор";
                                                break;
                                            case "д":
                                                student.AddressResidentialType = "деревня";
                                                break;
                                            case "п":
                                                student.AddressResidentialType = "посёлок";
                                                break;
                                        }
                                        student.AddressResidentialHouse = Regex.Match(cellValue.ToString(), housePattern).Groups[1].ToString();
                                        student.AddressResidentialStreet = Regex.Match(cellValue.ToString(), streetPattern).Groups[1].ToString().Trim();
                                        string housingMatch = Regex.Match(cellValue.ToString(), housingTypePattern).Groups[1].ToString().Trim();
                                        if (housingMatch == "к" || housingMatch == "к." || housingMatch == "корпус" || housingMatch == "корпус.")
                                            student.AddressResidentialHousingType = "корпус";
                                        else
                                            if (housingMatch == "стр" || housingMatch == "стр." || housingMatch == "строение" || housingMatch == "строение.")
                                            student.AddressResidentialHousingType = "строение";
                                        student.AddressResidentialHousing = Regex.Match(cellValue.ToString(), housingPattern).Groups[2].ToString().Trim();
                                        student.AddressResidentialApartment = Regex.Match(cellValue.ToString(), apartementPattern).Groups[1].ToString().Trim();
                                        student.AddressResidentialOblKrayAvtobl = Regex.Match(cellValue.ToString(), regionPattern).Groups[1].ToString().Trim();

                                    }
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
                                    if (cellValue is DateTime exDate) student.EducationReceivedDate = DateOnly.FromDateTime(exDate);
                                    else { 
                                        string dateStr = cellValue.ToString().Trim();
                                        if (DateTime.TryParseExact(
                                            dateStr,
                                            "dd.MM.yyyy",
                                            CultureInfo.InvariantCulture,
                                            DateTimeStyles.None,
                                            out DateTime parsedDate))
                                        {
                                            student.EducationReceivedDate = DateOnly.FromDateTime(parsedDate);
                                        }
                                        else throw new FormatException($"Не удалось распознать дату: {dateStr}");
                                    }
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
                            if ((header.ToLower() == "адрес проживания" ) && (cellValue.ToString() == "")) 
                            {
                                student.AddressResidentialCity = student.AddressRegistrationCity;
                                student.AddressResidentialApartment = student.AddressRegistrationApartment;
                                student.AddressResidentialDistrict = student.AddressRegistrationDistrict;
                                student.AddressResidentialHouse = student.AddressRegistrationHouse;
                                student.AddressResidentialHousing = student.AddressRegistrationHousing;
                                student.AddressResidentialHousingType = student.AddressRegistrationHousingType;
                                student.AddressResidentialIndex = student.AddressRegistrationIndex;
                                student.AddressResidentialOblKrayAvtobl = student.AddressRegistrationOblKrayAvtobl;
                                student.AddressResidentialStreet = student.AddressRegistrationStreet;
                                student.AddressResidentialType = student.AddressRegistrationType;
                            }
                            if (student.EducationReceivedEndYear == null) 
                            {
                                if (student.EducationReceivedDate.HasValue)
                                {
                                    student.EducationReceivedEndYear = (short)student.EducationReceivedDate.Value.Year;
                                }
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
