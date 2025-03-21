using Dekauto.Import.Service.API.Controllers;
using Dekauto.Import.Service.Domain.Entities;
using Dekauto.Import.Service.Domain.Entities.Adapters;
using Dekauto.Import.Service.Domain.Interfaces;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;
using Moq;

namespace ImportTest
{
    [TestClass]
    public class TestStudents
    {
        private Mock<IImportService> importService;
        private ImportController controller;

        [TestInitialize]
        public void Setup()
        {
            importService = new Mock<IImportService>();
            controller = new ImportController(importService.Object);
        }
        private IFormFile CreateMockFile(string fileName)
        {
            var fileMock = new Mock<IFormFile>();
            fileMock.Setup(f => f.FileName).Returns(fileName);
            fileMock.Setup(f => f.Length).Returns(1); 
            return fileMock.Object;
        }

        [TestMethod]
        public async Task Success()
        {
            // Arrange
            var files = new ImportFilesAdapter() 
            {
                ld = CreateMockFile("Файл личных дел.xlsx"),
                contract = CreateMockFile("Файл договоров.xlsx"),
                journal = CreateMockFile("Файл журнала.xlsx")
            };
            
            var students = new List<Student>();

            importService.Setup(s => s.GetStudentsLD(It.IsAny<IFormFile>())).ReturnsAsync(students);
            importService.Setup(s => s.GetStudentsContract(It.IsAny<IFormFile>(), students)).ReturnsAsync(students);
            importService.Setup(s => s.GetStudentsJournal(It.IsAny<IFormFile>(), students)).ReturnsAsync(students);

            controller.ControllerContext = new ControllerContext();

            // Act
            var result = await controller.ImportStudents(files);

            // Assert
            Assert.IsNotNull(result);
            var okResult = result as OkObjectResult; 

            Assert.IsNotNull(okResult);
            Assert.AreEqual(StatusCodes.Status200OK, okResult.StatusCode);

            var studentResult = okResult.Value as List<Student>; // Извлечение списка студентов
            Assert.IsNotNull(studentResult);
            Assert.AreEqual(students.Count, studentResult.Count); // Проверка количества студентов
        }
    }
}
