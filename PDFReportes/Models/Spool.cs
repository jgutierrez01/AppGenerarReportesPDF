using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFReportes.Models
{
    public class Spool
    {
        public string OrdenTrabajo { get; set; }
        public string Consecutivo { get; set; }
        public int NumeroPaginas { get; set; }
    }

    public class NoEncontrado
    {
        public string NumeroControl { get; set; }
        public string Reporte { get; set; }
        public string Error { get; set; }
    }

    public class NumReporte
    {
        public int SpoolID { get; set; }
        public string NumeroControl { get; set; }
        public string NumeroReporte { get; set; }
        public int TipoPruebaID { get; set; }
    }

    public class ReportePND
    {
        public string NumeroControl { get; set; }
        public int TipoPruebaID { get; set; }
        public string NumeroReporte { get; set; }
    }

    public class ReportePWHT
    {
        public string NumeroControl { get; set; }
        public string NumeroReporte { get; set; }
    }

    public class ReportePintura
    {
        public string NumeroControl { get; set; }
        public string NumeroReporte { get; set; }        
    }

    public class SpoolParticion
    {
        public string NumeroControl { get; set; }
        public int NumParticion { get; set; }
    }
}
