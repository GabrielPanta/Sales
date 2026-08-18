using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Sales.backend.Models;
using Sales.backend.Services;

namespace Sales.backend.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View(new ExcelMergeViewModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Index(ExcelMergeViewModel model)
        {
            if (!IsValidUpload(model.Trabajadores)) model.Errores.Add("Debe cargar el Excel de trabajadores.");
            if (!IsValidUpload(model.UltimoDiaLaborado)) model.Errores.Add("Debe cargar el Excel del ultimo dia laborado y actividades.");
            if (!IsValidUpload(model.Marcaciones)) model.Errores.Add("Debe cargar el Excel de marcaciones.");

            if (model.Errores.Any()) return View(model);

            try
            {
                var excel = new ExcelMergeService().Merge(model.Trabajadores, model.UltimoDiaLaborado, model.Marcaciones);
                return File(excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "trabajadores_con_estado.xlsx");
            }
            catch (Exception ex)
            {
                model.Errores.Add(ex.Message);
                return View(model);
            }
        }

        private static bool IsValidUpload(HttpPostedFileBase file)
        {
            return file != null && file.ContentLength > 0 && file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}