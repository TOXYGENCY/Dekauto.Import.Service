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

        [TestMethod]
        public async Task NotFound()
        {
            // Arrange
            var files = new ImportFilesAdapter()
            {
                ld = CreateMockFile("Файл личных дел.xlsx"),
                contract = CreateMockFile("Файл договоров.xlsx"),
                journal = CreateMockFile("Файл журнала.xlsx")
            };

            var students = new List<Student>();

            importService.Setup(s => s.GetStudentsLD(It.IsAny<IFormFile>())).ThrowsAsync(new ArgumentNullException());
            importService.Setup(s => s.GetStudentsContract(It.IsAny<IFormFile>(), students)).ThrowsAsync(new ArgumentNullException());
            importService.Setup(s => s.GetStudentsJournal(It.IsAny<IFormFile>(), students)).ThrowsAsync(new ArgumentNullException());

            controller.ControllerContext = new ControllerContext();

            // Act
            var result = await controller.ImportStudents(files);

            // Assert
            Assert.IsNotNull(result);
            var notFoundResult = result as NotFoundObjectResult;

            Assert.IsNotNull(notFoundResult);
            Assert.AreEqual(StatusCodes.Status404NotFound, notFoundResult.StatusCode);
        }

        [TestMethod]
        public async Task LoadError()
        {
            // Arrange
            var files = new ImportFilesAdapter()
            {
                ld = CreateMockFile("Файл личных дел.xlsx"),
                contract = CreateMockFile("Файл договоров.xlsx"),
                journal = CreateMockFile("Файл журнала.xlsx")
            };

            var students = new List<Student>();

            importService.Setup(s => s.GetStudentsLD(It.IsAny<IFormFile>())).ThrowsAsync(new FileLoadException());
            importService.Setup(s => s.GetStudentsContract(It.IsAny<IFormFile>(), students)).ThrowsAsync(new FileLoadException());
            importService.Setup(s => s.GetStudentsJournal(It.IsAny<IFormFile>(), students)).ThrowsAsync(new FileLoadException());

            controller.ControllerContext = new ControllerContext();

            // Act
            var result = await controller.ImportStudents(files);

            // Assert
            Assert.IsNotNull(result);
            var fileErrorResult = result as BadRequestObjectResult;

            Assert.IsNotNull(fileErrorResult);
            Assert.AreEqual(StatusCodes.Status400BadRequest, fileErrorResult.StatusCode);
        }

        [TestMethod]
        public async Task BadFileData()
        {
            // Arrange
            var files = new ImportFilesAdapter()
            {
                ld = CreateMockFile("Файл личных дел.xlsx"),
                contract = CreateMockFile("Файл договоров.xlsx"),
                journal = CreateMockFile("Файл журнала.xlsx")
            };

            var students = new List<Student>();

            importService.Setup(s => s.GetStudentsLD(It.IsAny<IFormFile>())).ThrowsAsync(new InvalidOperationException());
            importService.Setup(s => s.GetStudentsContract(It.IsAny<IFormFile>(), students)).ThrowsAsync(new InvalidOperationException());
            importService.Setup(s => s.GetStudentsJournal(It.IsAny<IFormFile>(), students)).ThrowsAsync(new InvalidOperationException());

            controller.ControllerContext = new ControllerContext();

            // Act
            var result = await controller.ImportStudents(files);

            // Assert
            Assert.IsNotNull(result);
            var badFileDataResult = result as BadRequestObjectResult;

            Assert.IsNotNull(badFileDataResult);
            Assert.AreEqual(StatusCodes.Status400BadRequest, badFileDataResult.StatusCode);
        }

        [TestMethod]
        public async Task Fatal()
        {
            // Arrange
            var files = new ImportFilesAdapter()
            {
                ld = CreateMockFile("Файл личных дел.xlsx"),
                contract = CreateMockFile("Файл договоров.xlsx"),
                journal = CreateMockFile("Файл журнала.xlsx")
            };

            var students = new List<Student>();

            importService.Setup(s => s.GetStudentsLD(It.IsAny<IFormFile>())).ThrowsAsync(new Exception());
            importService.Setup(s => s.GetStudentsContract(It.IsAny<IFormFile>(), students)).ThrowsAsync(new Exception());
            importService.Setup(s => s.GetStudentsJournal(It.IsAny<IFormFile>(), students)).ThrowsAsync(new Exception());

            controller.ControllerContext = new ControllerContext();

            // Act
            var result = await controller.ImportStudents(files);

            // Assert
            Assert.IsNotNull(result);
            var fatalResult = result as ObjectResult;

            Assert.IsNotNull(fatalResult);
            Assert.AreEqual(StatusCodes.Status500InternalServerError, fatalResult.StatusCode);
        }
    }
}
