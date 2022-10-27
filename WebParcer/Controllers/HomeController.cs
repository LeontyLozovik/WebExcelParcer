using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using WebParcer.DBContext;
using WebParcer.Models;
using WebParcer.Models.TableModels;
using WebParcer.Services;

namespace WebParcer.Controllers
{
    public class HomeController : Controller        //Главных контроллер который обрабатывает запрсы клиента
    {
        private readonly ILogger<HomeController> _logger;
        private IWebHostEnvironment _webHostEnvironment;
        private ExcelParcingService _excelParcing;
        private ApplicationDBContext _dBContext;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostEnvironment, 
            ExcelParcingService excelParcing, ApplicationDBContext dBContext)
        {
            _logger = logger;
            _webHostEnvironment = hostEnvironment;
            _excelParcing = excelParcing;
            _dBContext = dBContext;
        }

        public IActionResult Index()    //Главная страница
        {
            return View();
        }

        [HttpPost]
        public IActionResult UploadFiles(IFormFile file)        //Страница загрузки файлов
        {
            if (!(file.FileName.Contains(".xls") || file.FileName.Contains(".xlsx")))
            {
                return RedirectToAction("Index");
            }
                
            var nameOfFolfer = "Files";
            string directoryPath = Path.Combine(_webHostEnvironment.WebRootPath, nameOfFolfer);

            if(!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            var filePath = Path.Combine(directoryPath, file.FileName);

            using (FileStream fs = System.IO.File.Create(filePath))     //Сохранение файла по заданому пути
            {
                file.CopyTo(fs);
            }

            return RedirectToAction("Index");           //Возврат на главную страницу
        }

        [HttpGet]
        public IActionResult AllFiles()         //Страница отображения всех фалов на сервере
        {
            var nameOfFolfer = "Files";
            string directoryPath = Path.Combine(_webHostEnvironment.WebRootPath, nameOfFolfer);
            var allFiles = Directory.GetFiles(directoryPath);

            for (int i = 0; i < allFiles.Count(); i++)
            {
                allFiles[i] = allFiles[i].Remove(0, directoryPath.Length + 1);
            }

            ViewBag.AllFiles = allFiles;
            return View();
        }

        [HttpGet]
        public IActionResult FileContent(string filename)       //Отображение контента базы данных
        {
            _excelParcing.Parce(filename);      //Парсинг файла по пути
            
            //Добавление данных из баз данных
            ViewData["Class1"] = _dBContext.Class1s.ToList();
            ViewData["Class2"] = _dBContext.Class2s.ToList();
            ViewData["Class3"] = _dBContext.Class3s.ToList();
            ViewData["Class4"] = _dBContext.Class4s.ToList();
            ViewData["Class5"] = _dBContext.Class5s.ToList();
            ViewData["Class6"] = _dBContext.Class6s.ToList();
            ViewData["Class7"] = _dBContext.Class7s.ToList();
            ViewData["Class8"] = _dBContext.Class8s.ToList();
            ViewData["Class9"] = _dBContext.Class9s.ToList();
            ViewBag.Names = new List<string>() { "Class1", "Class2", "Class3", "Class4", "Class5", "Class6", "Class7", "Class8", "Class9" };
            
            //Дбавление названий таблиц
            ViewData["Class1Tag"] = Class1.Name;
            ViewData["Class2Tag"] = Class2.Name;
            ViewData["Class3Tag"] = Class3.Name;
            ViewData["Class4Tag"] = Class4.Name;
            ViewData["Class5Tag"] = Class5.Name;
            ViewData["Class6Tag"] = Class6.Name;
            ViewData["Class7Tag"] = Class7.Name;
            ViewData["Class8Tag"] = Class8.Name;
            ViewData["Class9Tag"] = Class9.Name;
            return View(new TableModelBase());
        }

    }
}