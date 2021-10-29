using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Tarea_5.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using System.Text.Json;
using OfficeOpenXml;

namespace Tarea_5.Controllers
{
    public class HomeController : Controller
    {
        

     
        public IActionResult Index()
        {

            

            return View();



        }

        [HttpPost]
        public IActionResult Datos()
        {
            EstudiantesModel estudiante = new EstudiantesModel();

            estudiante.Matricula = HttpContext.Request.Form["matricula"];
            estudiante.Nombres = HttpContext.Request.Form["nombres"];
            estudiante.Apellidos = HttpContext.Request.Form["apellidos"];
            estudiante.FechaDeNacimiento = HttpContext.Request.Form["fechaNacimiento"];
            estudiante.Carrera = HttpContext.Request.Form["carrera"];
            estudiante.Direccion = HttpContext.Request.Form["direccion"];
            estudiante.Telefono = HttpContext.Request.Form["telefono"];
            estudiante.Correo = HttpContext.Request.Form["email"];

            string exportType = HttpContext.Request.Form["ExportType"];
            if (exportType == "txt")
            {
                GenerateTxt(estudiante);
            }
            else if (exportType == "excel")
            {
                GenerateExcel(estudiante);
            }
            else if (exportType == "json")
            {
                GenerateJson(estudiante);
            }


            return View(estudiante);

        }



        public IActionResult GenerateTxt(EstudiantesModel estudiante)
        {
            string userPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string body = $"Matricula:  {estudiante.Matricula}, Nombres: {estudiante.Nombres}, Apellido: {estudiante.Apellidos}";
            body += $"Fecha de nacimiento: {estudiante.FechaDeNacimiento}, Carrera: {estudiante.Carrera}";
            body += $"Direccion: {estudiante.Direccion}, Telefono: {estudiante.Telefono}";
            body += $"Correo: {estudiante.Correo}";
            string downloadPath = Path.Combine(userPath, "Downloads\\log.txt");

            using (StreamWriter writer = new StreamWriter(downloadPath, true))
            {
                writer.WriteLine(body);
            }

            return View("Datos",estudiante);

        }

        public IActionResult GenerateExcel(EstudiantesModel estudiante)
        {
            
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            string userPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string downloadPath = Path.Combine(userPath, "Downloads\\log.xlsx");
            
            using (ExcelPackage excel = new ExcelPackage())
            {

            
                    excel.Workbook.Worksheets.Add("Worksheet1");
                    excel.Workbook.Worksheets.Add("Worksheet2");
                    excel.Workbook.Worksheets.Add("Worksheet3");
                
               

               

                var excelWorksheet = excel.Workbook.Worksheets["Worksheet1"];

                List<string[]> headerRow = new List<string[]>()
                {
                    new string[] { "Matricula", "Nombres", "Apellidos", "Direccion", "Telefono", "Correo" }
                };
                // Determine the header range (e.g. A1:D1)
                string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                // Target a worksheet
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];

                // Popular header row data
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);


                worksheet.Cells[headerRange].Style.Font.Bold = true;
                worksheet.Cells[headerRange].Style.Font.Size = 14;
                worksheet.Cells[headerRange].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                worksheet.Cells["A1"].Value = "ejemplo";




               var cellData = new List<object[]>()
               {
                new object[] {estudiante.Matricula,estudiante.Nombres,estudiante.Apellidos,estudiante.Direccion,estudiante.Telefono,estudiante.Correo},
           
               };

                excelWorksheet.Cells[2, 1].LoadFromArrays(cellData);
                    
                FileInfo excelFile = new FileInfo(downloadPath);
                excel.SaveAs(excelFile);


            }

            return View("Datos",estudiante);


        }

        public IActionResult GenerateJson(EstudiantesModel estudiante)
        {
            string userPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string json = JsonSerializer.Serialize(estudiante);

            string downloadPath = Path.Combine(userPath, "Downloads\\log.json");

            using (StreamWriter writer = new StreamWriter(downloadPath, true))
            {
                writer.WriteLine(json);
            }

            return View("Datos", estudiante);

        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
