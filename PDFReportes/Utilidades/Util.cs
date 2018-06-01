using DataAccess.Sql;
using iTextSharp.text;
using iTextSharp.text.pdf;
using PDFReportes.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace PDFReportes.Utilidades
{
    public class Util
    {                                                      
        public FileStream ArchivoCSV = null;
        private static readonly object _mutex = new object();
        private static Util _Instance;
        public static Util Instance
        {
            get
            {
                lock (_mutex)
                {
                    if(_Instance == null)
                    {
                        _Instance = new Util();
                    }
                }
                return _Instance;
            }
        }
                
        public bool verificaRutaServer()
        {
            bool existeRuta = false;
            try
            {
                string path = ConfigurationManager.AppSettings["RutaReportes"];
                string usuario = ConfigurationManager.AppSettings["Usuario"];
                string pass = ConfigurationManager.AppSettings["Pass"];                
                using (new NetworkConnection(path, new NetworkCredential(usuario, pass)))
                {
                    if (Directory.Exists(path))
                    {
                        existeRuta = true;
                    }
                    else
                    {
                        existeRuta = false;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
            return existeRuta;
        }        
        public List<Embarque> ObtieneEmbarques()
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                DataTable tblEmbarques = sql.EjecutaDataAdapter("PDFReportes_Get_NumeroEmbarque");                
                List<Embarque> listaEmbarque = new List<Embarque>();
                if(tblEmbarques.Rows.Count > 0)                
                    listaEmbarque.Add(new Embarque());
                for(int i = 0; i < tblEmbarques.Rows.Count; i++)
                {
                    listaEmbarque.Add(new Embarque
                    {
                        EmbarqueID = int.Parse(tblEmbarques.Rows[i][0].ToString()),
                        NumeroEmbarque = tblEmbarques.Rows[i][1].ToString()   
                    });
                }
                return listaEmbarque;
            }
            catch (Exception)
            {                
                return null;                
            }
        }        
        public List<Spool> ObtenerNumeroControl(string NumeroEmbarque)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                string[,] parametro = { { "@NumeroEmbarque", NumeroEmbarque } };
                DataTable tblOrdenTrabajo = sql.EjecutaDataAdapter("PDFReportes_Get_OrdenTrabajoXEmbarque", parametro);
                List <Spool> listaNumControl = new List<Spool>();
                if(tblOrdenTrabajo.Rows.Count > 0)
                {
                    for(int i = 0; i < tblOrdenTrabajo.Rows.Count; i++)
                    {
                        listaNumControl.Add(new Spool {
                            OrdenTrabajo = tblOrdenTrabajo.Rows[i][0].ToString(),
                            Consecutivo = tblOrdenTrabajo.Rows[i][1].ToString(),
                            NumeroPaginas = int.Parse(tblOrdenTrabajo.Rows[i][2].ToString())                            
                        });
                    }
                }
                return listaNumControl;
            }
            catch (Exception)
            {
                return null;
            }
        }        
        public List<ReportePND> ObtenerReportesPND_PorNumeroControl(string NumeroControl, int TipoPrueba)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                string[,] parametro = { { "@NumeroControl", NumeroControl }, { "@TipoPrueba", TipoPrueba.ToString() } };
                DataTable tblOrdenTrabajo = sql.EjecutaDataAdapter("PDFReporte_GET_PND_PorNumeroControl", parametro);
                List<ReportePND> listaNumControl = new List<ReportePND>();
                if (tblOrdenTrabajo.Rows.Count > 0)
                {
                    for (int i = 0; i < tblOrdenTrabajo.Rows.Count; i++)
                    {
                        listaNumControl.Add(new ReportePND
                        {                            
                            NumeroControl = tblOrdenTrabajo.Rows[i]["NumeroControl"].ToString(),
                            TipoPruebaID = int.Parse(tblOrdenTrabajo.Rows[i]["TipoPruebaID"].ToString()),
                            NumeroReporte = tblOrdenTrabajo.Rows[i]["NumeroReporte"].ToString()                            
                        });
                    }
                }
                return listaNumControl;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public List<ReportePWHT> ObtenerReportePWHT_PorNumeroControl(string NumeroControl)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                string[,] parametro = { { "@NumeroControl", NumeroControl } };
                DataTable tblOrdenTrabajo = sql.EjecutaDataAdapter("PDFReporte_GET_PWHT_PorNumeroControl", parametro);
                List<ReportePWHT> listaNumControl = new List<ReportePWHT>();
                if (tblOrdenTrabajo.Rows.Count > 0)
                {
                    for (int i = 0; i < tblOrdenTrabajo.Rows.Count; i++)
                    {
                        listaNumControl.Add(new ReportePWHT
                        {
                            NumeroControl = tblOrdenTrabajo.Rows[i]["NumeroControl"].ToString(),                            
                            NumeroReporte = tblOrdenTrabajo.Rows[i]["NumeroReporte"].ToString()
                        });
                    }
                }
                return listaNumControl;
            }
            catch (Exception)
            {
                return null;
            }
        }
        public List<ReportePintura> ObtenerReportePintura_PorNumeroControl(string NumeroControl, int Reporte)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                string[,] parametro = { { "@NumeroControl", NumeroControl }, { "@Reporte", Reporte.ToString() } };
                DataTable tblOrdenTrabajo = sql.EjecutaDataAdapter("PDFReportes_Pintura_GET_ReportesPinturaXNumeroControl", parametro);
                List<ReportePintura> listaNumControl = new List<ReportePintura>();
                if (tblOrdenTrabajo.Rows.Count > 0)
                {
                    for (int i = 0; i < tblOrdenTrabajo.Rows.Count; i++)
                    {
                        listaNumControl.Add(new ReportePintura
                        {
                            NumeroControl = tblOrdenTrabajo.Rows[i]["NumeroControl"].ToString(),
                            NumeroReporte = tblOrdenTrabajo.Rows[i]["NumeroReporte"].ToString()                            
                        });
                    }
                }
                return listaNumControl;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public bool ExisteArchivosTraveler(string OrdenTrabajo)
        {
            bool existe = false;
            try
            {
                string rutaTraveler = ConfigurationManager.AppSettings["RutaTraveler"].ToString();
                string UsuarioTraveler  = ConfigurationManager.AppSettings["UsuarioTraveler"].ToString();
                string PassTraveler = ConfigurationManager.AppSettings["PassTraveler"].ToString();
                string traveler = Path.Combine(rutaTraveler, "ODT " + OrdenTrabajo + ".pdf");
                using (new NetworkConnection(rutaTraveler, new NetworkCredential(UsuarioTraveler, PassTraveler)))
                {
                    if (File.Exists(traveler))
                {
                        existe = true;
                }
                else
                {
                    EscribirEnCSV(OrdenTrabajo, "Traveler", "No se encontro archivo Traveler");
                }
                }
                return existe;
            }
            catch (Exception e)
            {
                return false;
            }
        }
        public int getCantidadPaginas(string OrdenTrabajo)
        {
            string ruta = Path.Combine(ConfigurationManager.AppSettings["RutaTraveler"].ToString(), "ODT " + OrdenTrabajo + ".pdf");
            using (PdfReader reader = new PdfReader(ruta))
            {
                return reader.NumberOfPages;
            }
        }

        //TRAVELERS
        public void TravelerXSpool(string OrdenTrabajo, string NumeroControl, int NumPagina, string NumeroEmbarque)
        {            
            try
            {
                PdfReader reader = null;
                PdfCopy copia = null;
                Document doc = new Document();
                string RutaTraveler = ConfigurationManager.AppSettings["RutaTraveler"].ToString();                
                string rutaDestino = Path.GetTempPath();                
                if (ExisteArchivosTraveler(OrdenTrabajo))
                {
                    NumPagina = getCantidadPaginas(OrdenTrabajo) - NumPagina;
                    reader = new PdfReader(Path.Combine(RutaTraveler, "ODT " + OrdenTrabajo + ".pdf"));
                    if(reader != null)
                    {
                        
                        if(reader.NumberOfPages >= NumPagina)
                        {
                            copia = new PdfCopy(doc, new FileStream(rutaDestino + "\\" + NumeroControl + "_Traveler.pdf", FileMode.Create));
                            doc.Open();
                            copia.AddPage(copia.GetImportedPage(reader, NumPagina));
                            copia.Close();
                            doc.Close();
                        }
                        else
                        {
                            EscribirEnCSV(NumeroControl, "Traveler", "No coinciden los numeros de pagina");
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Traveler", "No se encontro traveler");
                    }
                    if (reader != null)
                        reader.Close();
                }
                else
                {
                    EscribirEnCSV(OrdenTrabajo, "Traveler", "No se encontro archivo principal traveler");
                }                
            }
            catch (Exception e)
            {
                EscribirEnCSV("Traveler", "Traveler", "Error: " + e.Message);                
            }            
        }
        //CERTIFICADOS
        public void CertificadoXSpool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                PdfReader reader = null;
                PdfCopy copia = null;
                Document doc = new Document();
                string rutaCertificado = ConfigurationManager.AppSettings["RutaCertificados"].ToString();
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaCertificado))
                {
                    if(File.Exists(Path.Combine(rutaCertificado, NumeroControl + ".pdf")))
                    {
                        reader = new PdfReader(Path.Combine(rutaCertificado, NumeroControl + ".pdf"));
                        if(reader != null)
                        {                            
                            if(reader.NumberOfPages >= 1)
                            {
                                copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + NumeroControl + "_Certificado.pdf", FileMode.Create));
                                doc.Open();
                                for (int i = 1; i <= reader.NumberOfPages; i++)
                                {
                                    copia.AddPage(copia.GetImportedPage(reader, i));
                                }
                                copia.Close();
                                doc.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Certificado Spool", "Pagina no encontrada");
                            }                            
                        }
                        else
                        {
                            EscribirEnCSV(NumeroControl, "Certificado Spool", "Certificado no encontrado");
                        }
                        reader.Close();
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Certificado Spool", "Certificado no encontrado");
                    }
                    
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Certificado Spool", "Carpeta de certificados no encontrada");
                }
                
            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "CertificadoSpool", "Error Obteniendo Certificado: " + e.Message);
            }
        }
        //PND'S
        public void RTPOST_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePND> Lista = ObtenerReportesPND_PorNumeroControl(NumeroControl, 5);
                string rutaPND = ConfigurationManager.AppSettings["RutaPND"].ToString();
                string rutaRTPOST = rutaPND + "\\RTPostTT";               
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaRTPOST))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaRTPOST, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();                                
                                reader = new PdfReader(Path.Combine(rutaRTPOST, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if(reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_RTPOST.pdf", FileMode.Create));
                                        doc.Open();
                                        for(int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();                                        
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte RTPOST", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte RTPOST", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte RTPOST", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaRTPOST = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_RTPOST.pdf"))
                            {
                                listaRTPOST.Add(rutaTemp + "\\" + item.NumeroReporte + "_RTPOST.pdf");
                            }                           
                        }
                        if(listaRTPOST != null && listaRTPOST.Count > 0)
                        {
                            MergePDF(listaRTPOST, rutaTemp + "\\" + NumeroControl + "_RTPOST.pdf");
                            eliminarArchivosPND(listaRTPOST);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte RTPOST", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte RTPOST", "No se Encontro carpeta de reportes");
                }
                        
            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte RTPOST", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void PTPOST_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePND> Lista = ObtenerReportesPND_PorNumeroControl(NumeroControl, 6);
                string rutaPND = ConfigurationManager.AppSettings["RutaPND"].ToString();
                string rutaPTPOST = rutaPND + "\\PTPostTT";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaPTPOST))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaPTPOST, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaPTPOST, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_PTPOST.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte PTPOST", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte PTPOST", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte PTPOST", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaPTPOST = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_PTPOST.pdf"))
                            {
                                listaPTPOST.Add(rutaTemp + "\\" + item.NumeroReporte + "_PTPOST.pdf");
                            }                            
                        }
                        if(listaPTPOST != null && listaPTPOST.Count > 0)
                        {
                            MergePDF(listaPTPOST, rutaTemp + "\\" + NumeroControl + "_PTPOST.pdf");
                            eliminarArchivosPND(listaPTPOST);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte PTPOST", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte PTPOST", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte PTPOST", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void UTPOST_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePND> Lista = ObtenerReportesPND_PorNumeroControl(NumeroControl, 14);
                string rutaPND = ConfigurationManager.AppSettings["RutaPND"].ToString();
                string rutaUTPOST = rutaPND + "\\UTPostTT";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaUTPOST))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaUTPOST, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaUTPOST, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_UTPOST.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte UTPOST", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte UTPOST", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte UTPOST", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaUTPOST = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_UTPOST.pdf"))
                            {
                                listaUTPOST.Add(rutaTemp + "\\" + item.NumeroReporte + "_UTPOST.pdf");
                            }                            
                        }
                        if(listaUTPOST != null && listaUTPOST.Count > 0)
                        {
                            MergePDF(listaUTPOST, rutaTemp + "\\" + NumeroControl + "_UTPOST.pdf");
                            eliminarArchivosPND(listaUTPOST);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte UTPOST", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte UTPOST", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte UTPOST", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void RT_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePND> Lista = ObtenerReportesPND_PorNumeroControl(NumeroControl, 1);
                string rutaPND = ConfigurationManager.AppSettings["RutaPND"].ToString();
                string rutaRT = rutaPND + "\\RT";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaRT))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaRT, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaRT, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_RT.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte RT", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte RT", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte RT", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaRT = new List<string>();
                        foreach (var item in Lista)
                        {
                            if (File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_RT.pdf"))
                            {
                                listaRT.Add(rutaTemp + "\\" + item.NumeroReporte + "_RT.pdf");
                            }
                        }
                        if (listaRT != null && listaRT.Count > 0)
                        {
                            MergePDF(listaRT, rutaTemp + "\\" + NumeroControl + "_RT.pdf");
                            eliminarArchivosPND(listaRT);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte RT", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte RT", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte RT", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void PT_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePND> Lista = ObtenerReportesPND_PorNumeroControl(NumeroControl, 2);
                string rutaPND = ConfigurationManager.AppSettings["RutaPND"].ToString();
                string rutaPT = rutaPND + "\\PT";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaPT))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaPT, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaPT, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_PT.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte PT", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte PT", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte PT", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaPT = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_PT.pdf"))
                            {
                                listaPT.Add(rutaTemp + "\\" + item.NumeroReporte + "_PT.pdf");
                            }                            
                        }
                        if(listaPT != null && listaPT.Count > 0)
                        {
                            MergePDF(listaPT, rutaTemp + "\\" + NumeroControl + "_PT.pdf");
                            eliminarArchivosPND(listaPT);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte PT", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte PT", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte PT", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void UT_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePND> Lista = ObtenerReportesPND_PorNumeroControl(NumeroControl, 8);
                string rutaPND = ConfigurationManager.AppSettings["RutaPND"].ToString();
                string rutaUT = rutaPND + "\\UT";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaUT))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaUT, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaUT, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_UT.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte UT", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte UT", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte UT", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaUT = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_UT.pdf"))
                            {
                                listaUT.Add(rutaTemp + "\\" + item.NumeroReporte + "_UT.pdf");
                            }                            
                        }
                        if(listaUT != null && listaUT.Count > 0)
                        {
                            MergePDF(listaUT, rutaTemp + "\\" + NumeroControl + "_UT.pdf");
                            eliminarArchivosPND(listaUT);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte UT", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte UT", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte UT", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void PMI_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePND> Lista = ObtenerReportesPND_PorNumeroControl(NumeroControl, 10);
                string rutaPND = ConfigurationManager.AppSettings["RutaPND"].ToString();
                string rutaPMI = rutaPND + "\\PMI";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaPMI))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaPMI, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaPMI, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_PMI.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte PMI", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte PMI", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte PMI", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaPMI = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_PMI.pdf"))
                            {
                                listaPMI.Add(rutaTemp + "\\" + item.NumeroReporte + "_PMI.pdf");
                            }
                        }
                        if (listaPMI != null && listaPMI.Count > 0)
                        {
                            MergePDF(listaPMI, rutaTemp + "\\" + NumeroControl + "_PMI.pdf");
                            eliminarArchivosPND(listaPMI);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte PMI", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte PMI", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte PMI", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        //PWHT
        public void PWHT_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePWHT> Lista = ObtenerReportePWHT_PorNumeroControl(NumeroControl);
                string rutaPND = ConfigurationManager.AppSettings["RutaPND"].ToString();
                string rutaPWHT = rutaPND + "\\PWHT";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaPWHT))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaPWHT, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaPWHT, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_PWHT.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte PWHT", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte PWHT", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte PWHT", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaPWHT = new List<string>();
                        foreach (var item in Lista)
                        {
                           if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_PWHT.pdf"))
                            {
                                listaPWHT.Add(rutaTemp + "\\" + item.NumeroReporte + "_PWHT.pdf");
                            }                            
                        }
                        if(listaPWHT != null && listaPWHT.Count > 0)
                        {
                            MergePDF(listaPWHT, rutaTemp + "\\" + NumeroControl + "_PWHT.pdf");
                            eliminarArchivosPND(listaPWHT);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte PWHT", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte PWHT", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte PWHT", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        //PINTURA
        public void SandBlast_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePintura> Lista = ObtenerReportePintura_PorNumeroControl(NumeroControl, 1);
                string rutaPintura = ConfigurationManager.AppSettings["RutaPintura"].ToString();
                string rutaSanBlast = rutaPintura + "\\Sandblast";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaSanBlast))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaSanBlast, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaSanBlast, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_SandBlast.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte SandBlast", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte SandBlast", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte SandBlast", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaSandBlast = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_SandBlast.pdf"))
                            {
                                listaSandBlast.Add(rutaTemp + "\\" + item.NumeroReporte + "_SandBlast.pdf");
                            }                            
                        }
                        if(listaSandBlast != null && listaSandBlast.Count > 0)
                        {
                            MergePDF(listaSandBlast, rutaTemp + "\\" + NumeroControl + "_SandBlast.pdf");
                            eliminarArchivosPND(listaSandBlast);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte SandBlast", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte SandBlast", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte SandBlast", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void Primario_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePintura> Lista = ObtenerReportePintura_PorNumeroControl(NumeroControl, 2);
                string rutaPintura = ConfigurationManager.AppSettings["RutaPintura"].ToString();
                string rutaPrimario = rutaPintura + "\\Primarios";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaPrimario))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaPrimario, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaPrimario, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_Primario.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte Primarios", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte Primarios", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte Primarios", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaPrimario = new List<string>();
                        foreach (var item in Lista)
                        {
                            if (File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_Primario.pdf"))
                            {
                                listaPrimario.Add(rutaTemp + "\\" + item.NumeroReporte + "_Primario.pdf");
                            }                            
                        }
                        if(listaPrimario != null && listaPrimario.Count > 0)
                        {
                            MergePDF(listaPrimario, rutaTemp + "\\" + NumeroControl + "_Primario.pdf");
                            eliminarArchivosPND(listaPrimario);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte Primarios", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte Primarios", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte Primarios", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void Intermedio_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePintura> Lista = ObtenerReportePintura_PorNumeroControl(NumeroControl, 3);
                string rutaPintura = ConfigurationManager.AppSettings["RutaPintura"].ToString();
                string rutaIntermedios = rutaPintura + "\\Intermedios";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaIntermedios))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaIntermedios, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaIntermedios, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_Intermedio.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte Intermedios", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte Intermedios", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte Intermedios", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaIntermedios = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_Intermedio.pdf"))
                            {
                                listaIntermedios.Add(rutaTemp + "\\" + item.NumeroReporte + "_Intermedio.pdf");
                            }                            
                        }
                        if(listaIntermedios != null && listaIntermedios.Count > 0)
                        {
                            MergePDF(listaIntermedios, rutaTemp + "\\" + NumeroControl + "_Intermedio.pdf");
                            eliminarArchivosPND(listaIntermedios);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte Intermedios", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte Intermedios", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte Intermedios", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void Acabado_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePintura> Lista = ObtenerReportePintura_PorNumeroControl(NumeroControl, 4);
                string rutaPintura = ConfigurationManager.AppSettings["RutaPintura"].ToString();
                string rutaAcabado = rutaPintura + "\\AcabadoVisual";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaAcabado))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaAcabado, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaAcabado, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_Acabado.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte Acabado Visual", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte Acabado Visual", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte Acabado Visual", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaAcabado = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_Acabado.pdf"))
                            {
                                listaAcabado.Add(rutaTemp + "\\" + item.NumeroReporte + "_Acabado.pdf");
                            }                            
                        }
                        if(listaAcabado != null && listaAcabado.Count > 0)
                        {
                            MergePDF(listaAcabado, rutaTemp + "\\" + NumeroControl + "_Acabado.pdf");
                            eliminarArchivosPND(listaAcabado);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte Acabado Visual", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte Acabado Visual", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte Acabado Visual", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void Adherencia_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePintura> Lista = ObtenerReportePintura_PorNumeroControl(NumeroControl, 5);
                string rutaPintura = ConfigurationManager.AppSettings["RutaPintura"].ToString();
                string rutaAdherencia = rutaPintura + "\\Adherencia";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaAdherencia))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaAdherencia, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaAdherencia, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_Adherencia.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte Adherencia", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte Adherencia", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte Adherencia", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaAdherencia = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_Adherencia.pdf"))
                            {
                                listaAdherencia.Add(rutaTemp + "\\" + item.NumeroReporte + "_Adherencia.pdf");
                            }                            
                        }
                        if(listaAdherencia != null && listaAdherencia.Count > 0)
                        {
                            MergePDF(listaAdherencia, rutaTemp + "\\" + NumeroControl + "_Adherencia.pdf");
                            eliminarArchivosPND(listaAdherencia);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte Adherencia", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte Adherencia", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte Adherencia", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        public void PullOf_Holiday_X_Spool(string NumeroControl, string NumeroEmbarque)
        {
            try
            {
                List<ReportePintura> Lista = ObtenerReportePintura_PorNumeroControl(NumeroControl, 6);
                string rutaPintura = ConfigurationManager.AppSettings["RutaPintura"].ToString();
                string rutaPullof = rutaPintura + "\\PullOff";
                string rutaHoliday = rutaPintura + "\\Holiday";
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaPullof) && Directory.Exists(rutaHoliday))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaPullof, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaPullof, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_PullOf.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte PullOf", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte PullOf", "Paginas no encontradas");
                                }
                                reader.Close();
                            }else if (File.Exists(Path.Combine(rutaHoliday, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaHoliday, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_Holiday.pdf", FileMode.Create));
                                        doc.Open();
                                        for (int i = 1; i <= reader.NumberOfPages; i++)
                                        {
                                            copia.AddPage(copia.GetImportedPage(reader, i));
                                        }
                                        copia.Close();
                                        doc.Close();
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroControl, "Reporte Holiday", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte Holiday", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte PullOf", "No se encontro Reporte");
                                EscribirEnCSV(NumeroControl, "Reporte Holiday", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool PullOF
                        List<string> listaPullOf = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_PullOf.pdf"))
                            {
                                listaPullOf.Add(rutaTemp + "\\" + item.NumeroReporte + "_PullOf.pdf");
                            }                            
                        }
                        if(listaPullOf != null && listaPullOf.Count > 0)
                        {
                            MergePDF(listaPullOf, rutaTemp + "\\" + NumeroControl + "_PullOf.pdf");
                            eliminarArchivosPND(listaPullOf);
                        }
                        
                        //Mezclar reportes por spool Holiday
                        List<string> listaHoliday = new List<string>();
                        foreach (var item in Lista)
                        {
                            if(File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_Holiday.pdf"))
                            {
                                listaHoliday.Add(rutaTemp + "\\" + item.NumeroReporte + "_Holiday.pdf");
                            }                            
                        }
                        if(listaHoliday != null && listaHoliday.Count > 0)
                        {
                            MergePDF(listaHoliday, rutaTemp + "\\" + NumeroControl + "_Holiday.pdf");
                            eliminarArchivosPND(listaHoliday);
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte PullOf", "No se encontro ningun Reporte");
                        EscribirEnCSV(NumeroControl, "Reporte Holiday", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte PullOf", "No se Encontro carpeta de reportes");
                    EscribirEnCSV(NumeroControl, "Reporte Holiday", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte PullOf", "Error Obteniendo Numeros de Reporte: " + e.Message);
                EscribirEnCSV(NumeroControl, "Reporte Holiday", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }

        public bool GenerarParticionamiento(List<Spool> Spools, string NumeroEmbarque, string rutaDestino)
        {
            bool estatus = false;
            try
            {
                int limiteMB = int.Parse(ConfigurationManager.AppSettings["LimiteMB"].ToString());
                //double limiteToBytes = (limiteMB * (1024 * 1024.0));
                string rutaTemp = Path.GetTempPath();                
                Decimal suma = 0;                
                int numParticion = 1;
                List<SpoolParticion> ListaTmp = new List<SpoolParticion>();
                List<SpoolParticion> SpoolParticion = new List<SpoolParticion>();
                //Obtiene el tamaño total de los archivos generados
                foreach(var item in Spools)
                {             
                    if(File.Exists(Path.Combine(rutaTemp, item.OrdenTrabajo + "-" + item.Consecutivo + ".pdf")))
                    {
                        suma += Decimal.Parse(ConvertToMegabytes(UInt64.Parse((new FileInfo(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + ".pdf")).Length.ToString())));
                        if (suma < Decimal.Parse(limiteMB.ToString()))
                        {
                            SpoolParticion.Add(new SpoolParticion
                            {
                                NumeroControl = item.OrdenTrabajo + "-" + item.Consecutivo,
                                NumParticion = numParticion,
                            });
                        }
                        else
                        {
                            numParticion++;
                            suma = 0;
                            suma += Decimal.Parse(ConvertToMegabytes(UInt64.Parse((new FileInfo(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + ".pdf")).Length.ToString())));
                            SpoolParticion.Add(new SpoolParticion
                            {
                                NumeroControl = item.OrdenTrabajo + "-" + item.Consecutivo,
                                NumParticion = numParticion,
                            });

                        }
                    }                           
                }
                if(SpoolParticion != null && SpoolParticion.Count > 0)
                {
                    if(SpoolParticion.Count == 1)
                    {
                        if (Decimal.Parse(ConvertToMegabytes(UInt64.Parse((new FileInfo(rutaTemp + "\\" + SpoolParticion[0].NumeroControl + ".pdf")).Length.ToString()))) > 0)
                        {
                            GenerarArchivoGeneralParticionado(SpoolParticion, NumeroEmbarque, rutaDestino);
                            estatus = true;
                        }
                        else
                        {
                            estatus = false;
                        }
                    }
                    else
                    {
                        GenerarArchivoGeneralParticionado(SpoolParticion, NumeroEmbarque, rutaDestino);
                        estatus = true;
                    }                                        
                }
            }
            catch (Exception e)
            {
                EscribirEnCSV("GENERAL", "GENERAL", "Error al particionar archivos: " + e.Message);
                estatus = false;
            }
            return estatus;
        }
        static string ConvertToMegabytes(ulong bytes)
        {
            return ((decimal)bytes / 1024M / 1024M).ToString("F1");
        }


        public void GenerarArchivoGeneralParticionado(List<SpoolParticion> ListaSpool, string NumeroEmbarque, string rutaDestino)
        {
            try
            {
                string rutaTemp = Path.GetTempPath();                
                List<string> rutasArchivos = null;                
                List<string> Paths = new List<string>();
                List<SpoolParticion> L1 = ListaSpool.GroupBy(a => a.NumParticion).Select(b => b.First()).ToList();
                             
                if (L1.Count == 1 && L1[0].NumParticion != 0)
                {
                    rutasArchivos = new List<string>();
                    foreach (var item in ListaSpool)
                    {
                        if (File.Exists(Path.Combine(rutaTemp, item.NumeroControl + ".pdf")))
                            rutasArchivos.Add(Path.Combine(rutaTemp, item.NumeroControl + ".pdf"));
                    }
                    MergePDF(rutasArchivos, rutaDestino + "\\" + NumeroEmbarque + ".pdf");                                 
                }
                else
                {
                    foreach(var item in L1)
                    {
                        rutasArchivos = new List<string>();
                        foreach(var item2 in ListaSpool)
                        {
                            if(item.NumParticion == item2.NumParticion)
                            {
                                if (File.Exists(Path.Combine(rutaTemp, item2.NumeroControl + ".pdf")))
                                    rutasArchivos.Add(Path.Combine(rutaTemp, item2.NumeroControl + ".pdf"));
                            }                           
                        }
                        MergePDF(rutasArchivos, rutaDestino + "\\" + NumeroEmbarque + " Part " + item.NumParticion + ".pdf");
                        rutasArchivos = null;
                    }                    
                }

                // ELIMINA ARCHIVOS TMP DE NUMERO DE CONTROL                
                foreach (var item in ListaSpool)
                {                    
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl + ".pdf"))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl + ".pdf", FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl + ".pdf");
                    }
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV("GENERAL", "GENERAL", "Ocurrio un error al unir archivos PDF: " + e.Message);
            }
        }

        public void eliminarArchivosPND(List<string> lista)
        {
            try
            {
                foreach(var item in lista)
                {
                    if (File.Exists(item))
                    {
                        File.SetAttributes(item, FileAttributes.Normal);
                        File.Delete(item);
                    }
                }
            }
            catch (Exception e)
            {
                throw;
            }
        }
        public void UnirArchivos(List<Spool> ListaSpool)
        {
            try
            {                
                string rutaTemp = Path.GetTempPath();
                string extTraveler = "_Traveler.pdf",
                    extCertificado = "_Certificado.pdf",
                    extRTPOST = "_RTPOST.pdf",
                    extPTPOST = "_PTPOST.pdf",
                    extUTPOST = "_UTPOST.pdf",
                    extRT = "_RT.pdf",
                    extPT = "_PT.pdf",
                    extUT = "_UT.pdf",
                    extPMI = "_PMI.pdf", extPWHT = "_PWHT.pdf", extSandBlast = "_SandBlast.pdf", extPrimario = "_Primario.pdf",
                    extIntermedio = "_Intermedio.pdf", extAcabado = "_Acabado.pdf", extAdherencia = "_Adherencia.pdf", extPullOf = "_PullOf.pdf",
                    extHoliday = "_Holiday.pdf";
                List<string> rutasArchivos = null;
                string NumeroControl = "";
                foreach (var item in ListaSpool)
                {
                    NumeroControl = "";
                    NumeroControl = item.OrdenTrabajo + "-" + item.Consecutivo;
                    rutasArchivos = new List<string>();
                    //TRAVELER
                    if(File.Exists(Path.Combine(rutaTemp, NumeroControl + extTraveler)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extTraveler));
                    //CERTIFICADOS
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extCertificado)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extCertificado));
                    //PND'S
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extRTPOST)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extRTPOST));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extPTPOST)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extPTPOST));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extUTPOST)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extUTPOST));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extRT)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extRT));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extPT)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extPT));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extUT)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extUT));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extPMI)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extPMI));
                    //PWHT
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extPWHT)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extPWHT));
                    //PINTURA
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extSandBlast)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extSandBlast));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extPrimario)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extPrimario));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extIntermedio)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extIntermedio));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extAcabado)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extAcabado));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extAdherencia)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extAdherencia));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extPullOf)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extPullOf));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extHoliday)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extHoliday));

                    if(rutasArchivos != null)
                    {
                        MergePDF(rutasArchivos, rutaTemp + "\\" + NumeroControl + ".pdf");
                    }                    
                    rutasArchivos = null;
                }

                foreach (var item in ListaSpool)
                {
                    // ELIMINA TMP TRAVELER
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extTraveler))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extTraveler, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extTraveler);
                    }
                    // ELIMINA TMP CERTIFICADOS
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extCertificado))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extCertificado, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extCertificado);
                    }
                    // ELIMINA TMP PNDS
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extRTPOST))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extRTPOST, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extRTPOST);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPTPOST))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPTPOST, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPTPOST);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extUTPOST))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extUTPOST, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extUTPOST);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extRT))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extRT, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extRT);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPT))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPT, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPT);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extUT))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extUT, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extUT);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPMI))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPMI, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPMI);
                    }
                    // ELIMINA TMP PWHT
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPWHT))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPWHT, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPWHT);
                    }
                    // ELIMINA TMP PINTURA
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extSandBlast))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extSandBlast, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extSandBlast);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPrimario))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPrimario, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPrimario);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extIntermedio))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extIntermedio, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extIntermedio);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extAcabado))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extAcabado, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extAcabado);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extAdherencia))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extAdherencia, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extAdherencia);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPullOf))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPullOf, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extPullOf);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extHoliday))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extHoliday, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.OrdenTrabajo + "-" + item.Consecutivo + extHoliday);
                    }

                }

            }
            catch (Exception e)
            {
                EscribirEnCSV("GENERAL", "GENERAL", "Ocurrio un error al unir archivos PDF: " + e.Message);
            }
        }
      
        //METODO PARA UNIR PDF
        public bool MergePDF(List<String> InFiles, String OutFile)
        {
            bool ret = false;
            using (FileStream stream = new FileStream(OutFile, FileMode.Create))
            using (Document doc = new Document())
            using (PdfCopy pdf = new PdfCopy(doc, stream))
            {
                doc.Open();

                PdfReader reader = null;
                PdfImportedPage page = null;

                //fixed typo
                InFiles.ForEach(file =>
                {
                    try
                    {                        
                        reader = new PdfReader(file);
                        if (reader != null)
                        {
                            for (int i = 0; i < reader.NumberOfPages; i++)
                            {
                                page = pdf.GetImportedPage(reader, i + 1);
                                pdf.AddPage(page);
                            }
                            pdf.FreeReader(reader);
                            reader.Close();
                        }
                    }
                    catch (Exception e)
                    {
                        
                    }
                                  
                });
                ret = true;   
            }
            return ret;
        }

        //MANEJADOR DE CSV
        public bool crearArchivoCSV(string rutaDestino, string NumeroEmbarque)
        {
            try
            {
                ArchivoCSV = File.Create(rutaDestino + "\\Log_" + NumeroEmbarque + ".csv");
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public bool EscribirEnCSV(string NumeroControl, string Reporte, string Error)
        {
            try
            {
                using (StringWriter write = new StringWriter())
                {
                    write.WriteLine(NumeroControl + "," + Reporte + "," + Error);
                    byte[] info = new UTF8Encoding(true).GetBytes(write.ToString());
                    ArchivoCSV.Write(info, 0, info.Length);
                }
                return true;
            }
            catch (Exception)
            {
                return false;                
            }
        }
        public bool CerrarArchivoCSV()
        {
            try
            {
                ArchivoCSV.Close();
                return false;
            }
            catch (Exception)
            {
                return false;                
            }
        }
                
        //VERIFICAR SI YA EXISTE ARCHIVO PDF
        public string VerificaSiExistePDF(string NumeroEmbarque, string rutaDestino)
        {
            string result = "";
            try
            {                
                string[] archivos = Directory.GetFiles(rutaDestino, NumeroEmbarque + "*.pdf");
                if(archivos.Length > 0)
                {
                    foreach(var i in archivos)
                    {
                        result += Path.GetFileName(i) + Environment.NewLine;                        
                    }                    
                }                                
            }
            catch (Exception e)
            {
                EscribirEnCSV("GENERAL", "GENERAL", "Error verificando si existe archivo");
                result = "";
            }
            return result;
        }
    }
}
    