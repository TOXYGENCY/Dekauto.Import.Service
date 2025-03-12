namespace Dekauto.Import.Service.Domain.Entities
{
    public class Student
    {
        public string? Name { get; set; } // Имя +
        public string? Surname { get; set; } // Фамилия +
        public string? Pathronymic { get; set; } // Отчество +
        public bool? Gender { get; set; } // Пол +
        public DateTime? BirthdayDate { get; set; } // Дата рождения +
        public string? BirthdayPlace { get; set; } // Место рождения +
        public string? PhoneNumber { get; set; } // Номер телефона +
        public string? Email { get; set; } // Email +
        public string? GradeBook { get; set; } // Номер зачетной книжки +
        public string? PassportSerial { get; set; } // Серия паспорта +
        public string? PassportNumber { get; set; } // Номер паспорта +
        public string? PassportIssuancePlace { get; set; } // Место выдачи паспорта +
        public DateTime? PassportIssuanceDate { get; set; } // Дата выдачи паспорта +
        public string? PassportIssuanceCode { get; set; } // Код выдачи паспорта +
        public string? Citizenship { get; set; } // Гражданство +
        public string? AddressRegistrationIndex { get; set; } // Почтовый индекс регистрации +
        public string? AddressRegistrationOblKrayAvtobl { get; set; } // Область/край/автономный округ регистрации +
        public string? AddressRegistrationDistrict { get; set; } // Район регистрации +
        public string? AddressRegistrationType { get; set; }//Тип населенного пункта регистрация +
        public string? AddressRegistrationStreet { get; set; } // Улица регистрации +
        public string? AddressRegistrationCity { get; set; } // Город регистрации +
        public string? AddressRegistrationHouse { get; set; } // Дом регистрации +
        public string? AddressRegistrationHousingType { get; set; } // Тип корпуса регистрации +
        public string? AddressRegistrationHousing { get; set; } // Корпус регистрации +
        public string? AddressRegistrationApartment { get; set; } // Квартира регистрации +
        public string? AddressResidentialIndex { get; set; } // Почтовый индекс проживания +
        public string? AddressResidentialOblKrayAvtobl { get; set; } // Область/край/автономный округ проживания +
        public string? AddressResidentialDistrict { get; set; } // Район проживания +
        public string? AddressResidentialType { get; set; }//Тип населенного пункта проживания +
        public string? AddressResidentialStreet { get; set; } // Улица проживания +
        public string? AddressResidentialCity { get; set; } // Город проживания +
        public string? AddressResidentialHouse { get; set; } // Дом проживания +
        public string? AddressResidentialHousingType { get; set; } // Тип корпуса проживания +
        public string? AddressResidentialHousing { get; set; } // Корпус проживания +
        public string? AddressResidentialApartment { get; set; } // Квартира проживания +
        public bool? LivingInDormitory { get; set; } // Проживание в общежитии +
        public DateTime? EnrollementOrderDate { get; set; } // Дата приказа о зачислении +
        public string? EnrollementOrderNum { get; set; } // Номер приказа о зачислении +
        public string? GroupName { get; set; } // Название группы
        public string? GiaExam1Name { get; set; } // Название экзамена 1 +
        public short? GiaExam1Score { get; set; } // Оценка экзамена 1 +
        public string? GiaExam1Note { get; set; } // Примечание по экзамену 1 +
        public string? GiaExam2Name { get; set; } // Название экзамена 2 +
        public short? GiaExam2Score { get; set; } // Оценка экзамена 2 +
        public string? GiaExam2Note { get; set; } // Примечание по экзамену 2 +
        public string? GiaExam3Name { get; set; } // Название экзамена 3 +
        public short? GiaExam3Score { get; set; } // Оценка экзамена 3 +
        public string? GiaExam3Note { get; set; } // Примечание по экзамену 3 +
        public string? GiaExam4Name { get; set; } // Название экзамена 4
        public short? GiaExam4Score { get; set; } // Оценка экзамена 4
        public string? GiaExam4Note { get; set; } // Примечание по экзамену 4
        public string? EducationReceived { get; set; } // Полученное образование +
        public string? EducationReceivedSerial { get; set; } // Серия полученного образования +
        public string? EducationReceivedNum { get; set; } // Номер полученного образования +
        public DateTime? EducationReceivedDate { get; set; } // Дата полученного образования +
        public string? OOName { get; set; } // Название образовательной организации +
        public string? OOAddress { get; set; } // Адрес образовательной организации +
        public short? EducationReceivedEndYear { get; set; } // Дата окончания ОО +
        public short? BonusScores { get; set; } // Доп баллы, в данном контексте влияет на отличия +
        public bool? MilitaryService { get; set; } // Военная служба +
        public bool? MaritalStatus { get; set; } // Семейное положение +
        public string? Education { get; set; } // Получаемое образование +
        public string? EducationForm { get; set; } // Форма обучения +
        public string? Faculty { get; set; } // Факультет +
        public string? CourseOfTraining { get; set; } // Направление подготовки (специальность) +
        public string? Course { get; set; } // Направленность (специализация) +
        public short? EducationStartYear { get; set; } // Год начала образования +
        public short? EducationFinishYear { get; set; } // Год окончания образования +
        public short? EducationTime { get; set; } // Время обучения +
        public string? EducationBase { get; set; } // Основа обучения + 
        public string? EducationRelationForm { get; set; } // Форма отношений +
        public string? EducationRelationNum { get; set; } // Номер отношений с учебным заведением +
        public DateTime? EducationRelationDate { get; set; } // Дата начала отношений с учебным заведением +
    }
}
