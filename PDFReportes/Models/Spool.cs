using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFReportes.Models
{
    public class Spool
    {
        public string NumeroControl { get; set; }
        public string OrdenTrabajo { get; set; }
        public string Consecutivo { get; set; }
        public int NumeroPaginas { get; set; }
        public bool AplicaGranel { get; set; }
        public string Dimensional { get; set; }
    }

    public class Reporte
    {        
        public string NumeroControl { get; set; }
        public string NumeroReporte { get; set; }        
    }    

    public class SpoolParticion
    {
        public string NumeroControl { get; set; }
        public int NumParticion { get; set; }
    }

    public class NumeroControlClass
    {
        public string NumeroControl { get; set; }
    }
    public class NumeroReportePullHoliday
    {
        public string NumeroReporte { get; set; }
    }
}
