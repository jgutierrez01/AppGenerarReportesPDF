using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDFReportes.Models
{
    public class Embarque
    {
        public int EmbarqueID { get; set; }
        public string NumeroEmbarque { get; set; }
        public int OrigenBD { get; set; }
        public Embarque()
        {
            EmbarqueID = 0;
            NumeroEmbarque = "";
            OrigenBD = 0;
        }
    }

    public class Proyecto
    {
        public int ProyectoID { get; set; }
        public string Nombre { get; set; }
        public string RutaTraveler { get; set; }
        public string RutaReportes { get; set; }
        public string FolioDimensional { get; set; }
        public Proyecto()
        {
            ProyectoID = 0;
            Nombre = "";
            RutaTraveler = "";
            RutaReportes = "";
            FolioDimensional = "";
        }
    }
}
