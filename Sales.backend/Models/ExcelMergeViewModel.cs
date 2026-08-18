using System.Collections.Generic;
using System.Web;

namespace Sales.backend.Models
{
    public class ExcelMergeViewModel
    {
        public HttpPostedFileBase Trabajadores { get; set; }
        public HttpPostedFileBase UltimoDiaLaborado { get; set; }
        public HttpPostedFileBase Marcaciones { get; set; }
        public IList<string> Errores { get; set; }

        public ExcelMergeViewModel()
        {
            Errores = new List<string>();
        }
    }
}
