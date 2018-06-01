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
        public Embarque()
        {
            EmbarqueID = 0;
            NumeroEmbarque = "";
        }
    }
}
