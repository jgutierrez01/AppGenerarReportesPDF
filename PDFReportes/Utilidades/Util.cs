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
using System.Text;

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
                //using (new NetworkConnection(path, new NetworkCredential(usuario, pass)))
                //{
                    if (Directory.Exists(path))
                    {
                        existeRuta = true;
                    }
                    else
                    {
                        existeRuta = false;
                    }
                //}
            }
            catch (Exception)
            {
                return false;
            }
            return existeRuta;
        }        
        ///////////////////////////////////////// EMPIEZA SEGMENTO DE STORED PROCEDURES ///////////////////////////////////////
        public List<Proyecto> ObtenerProyectos()
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                DataTable tblProyecto = sql.EjecutaDataAdapter("ReportesPDF_GET_Proyectos");
                List<Proyecto> listaProyectos = new List<Proyecto>();
                if (tblProyecto.Rows.Count > 0)
                    listaProyectos.Add(new Proyecto());
                for (int i = 0; i < tblProyecto.Rows.Count; i++)
                {
                    listaProyectos.Add(new Proyecto
                    {
                        ProyectoID = int.Parse(tblProyecto.Rows[i]["ProyectoID"].ToString()),
                        Nombre = tblProyecto.Rows[i]["Nombre"].ToString(),
                        RutaTraveler = tblProyecto.Rows[i]["RutaTraveler"].ToString(),
                        RutaReportes = tblProyecto.Rows[i]["RutaReportes"].ToString(),
                        FolioDimensional = tblProyecto.Rows[i]["FolioDimensional"].ToString()                       
                    });
                }
                return listaProyectos;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public List<Embarque> ObtieneEmbarques(int ProyectoID)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();                
                string[,] parametros = { { "@ProyectoID", ProyectoID.ToString() } };
                DataTable tblEmbarques = sql.EjecutaDataAdapter("ReportesPDF_GET_Embarques", parametros);
                List<Embarque> listaEmbarque = new List<Embarque>();
                if(tblEmbarques.Rows.Count > 0)                
                    listaEmbarque.Add(new Embarque());
                for(int i = 0; i < tblEmbarques.Rows.Count; i++)
                {
                    listaEmbarque.Add(new Embarque
                    {
                        EmbarqueID = int.Parse(tblEmbarques.Rows[i]["EmbarqueID"].ToString()),
                        NumeroEmbarque = tblEmbarques.Rows[i]["NumeroEmbarque"].ToString(),
                        OrigenBD = int.Parse(tblEmbarques.Rows[i]["OrigenBD"].ToString())
                    });
                }
                return listaEmbarque;
            }
            catch (Exception)
            {                
                return null;                
            }
        }
        
        public List<Spool> ObtenerNumeroControl(int ProyectoID, Embarque emb)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                string[,] parametro = {
                    { "@ProyectoID", ProyectoID.ToString() },
                    { "@EmbarqueID", emb.EmbarqueID.ToString() },
                    { "@OrigenBD", emb.OrigenBD.ToString() }
                };
                //DataTable tblOrdenTrabajo = sql.EjecutaDataAdapter("ReportesPDF_GET_HojasTravelers", parametro);
                DataTable tblOrdenTrabajo = sql.EjecutaDataAdapter("ReportesPDF_GET_HojasTravelers_V2", parametro);
                List <Spool> listaNumControl = new List<Spool>();
                if(tblOrdenTrabajo.Rows.Count > 0)
                {
                    for(int i = 0; i < tblOrdenTrabajo.Rows.Count; i++)
                    {
                        listaNumControl.Add(new Spool {
                            NumeroControl = tblOrdenTrabajo.Rows[i]["NumeroControl"].ToString(),
                            OrdenTrabajo = tblOrdenTrabajo.Rows[i]["OrdenTrabajo"].ToString(),
                            Consecutivo = tblOrdenTrabajo.Rows[i]["SpoolID"].ToString(),
                            NumeroPaginas = int.Parse(tblOrdenTrabajo.Rows[i]["Paginas"].ToString()),
                            Dimensional = tblOrdenTrabajo.Rows[i]["Dimensional"].ToString()
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
        public List<Reporte > ObtenerReportesPND_PorNumeroControl(int ProyectoID, string NumeroControl, int TipoPrueba)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                string[,] parametro = { { "@ProyectoID", ProyectoID.ToString() }, { "@NumeroControl", NumeroControl }, { "@TipoPrueba", TipoPrueba.ToString() } };                
                DataTable tblOrdenTrabajo = sql.EjecutaDataAdapter("ReportesPDF_GET_ReportesPND", parametro);
                List<Reporte > listaNumControl = new List<Reporte >();
                if (tblOrdenTrabajo.Rows.Count > 0)
                {
                    for (int i = 0; i < tblOrdenTrabajo.Rows.Count; i++)
                    {
                        listaNumControl.Add(new Reporte 
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
        public List<Reporte > ObtenerReportesTT_PorNumeroControl(int ProyectoID, string NumeroControl, int TipoPrueba)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                string[,] parametro = { { "@ProyectoID", ProyectoID.ToString() }, { "@NumeroControl", NumeroControl }, { "@TipoPrueba", TipoPrueba.ToString() } };                
                DataTable tblOrdenTrabajo = sql.EjecutaDataAdapter("ReportesPDF_GET_ReportesTT", parametro);
                List<Reporte > listaNumControl = new List<Reporte >();
                if (tblOrdenTrabajo.Rows.Count > 0)
                {
                    for (int i = 0; i < tblOrdenTrabajo.Rows.Count; i++)
                    {
                        listaNumControl.Add(new Reporte 
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
        public List<Reporte > ObtenerReportePintura_PorNumeroControl(int ProyectoID, string NumeroControl, int Reporte)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                string[,] parametro = { { "@ProyectoID", ProyectoID.ToString() }, { "@NumeroControl", NumeroControl }, { "@Reporte", Reporte.ToString() } };                
                DataTable tblOrdenTrabajo = sql.EjecutaDataAdapter("ReportesPDF_GET_ReportesPintura", parametro);
                List<Reporte > listaNumControl = new List<Reporte >();
                if (tblOrdenTrabajo.Rows.Count > 0)
                {
                    for (int i = 0; i < tblOrdenTrabajo.Rows.Count; i++)
                    {
                        listaNumControl.Add(new Reporte 
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

        public List<NumeroReportePullHoliday> ObtenerReportePullOf_Holiday(int ProyectoID, DataTable NumeroControl, int Reporte)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                string[,] parametro = {
                    { "@ProyectoID", ProyectoID.ToString() },
                    { "@Reporte", Reporte.ToString() }
                };                
                DataTable tblNumReporte = sql.EjecutaDataAdapter("ReportesPDF_GET_ReportePullOf_Holiday", NumeroControl, "@Spools", parametro);
                List<NumeroReportePullHoliday> lstNumReporte = new List<NumeroReportePullHoliday>();
                if (tblNumReporte.Rows.Count > 0)
                {
                    for (int i = 0; i < tblNumReporte.Rows.Count; i++)
                    {
                        lstNumReporte.Add(new NumeroReportePullHoliday
                        {                            
                            NumeroReporte = tblNumReporte.Rows[i]["NumeroReporte"].ToString()
                        });
                    }
                }
                return lstNumReporte;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public List<Reporte> ObtenerReporteEspesores(int ProyectoID, string NumeroControl)
        {
            try
            {
                ObjetosSQL sql = new ObjetosSQL();
                string[,] parametro = { { "@ProyectoID", ProyectoID.ToString() }, { "@NumeroControl", NumeroControl } };
                DataTable tblReporteEspesores = sql.EjecutaDataAdapter("ReportesPDF_GET_ReporteEspesores", parametro);
                List<Reporte> listaReportes = new List<Reporte>();
                if (tblReporteEspesores.Rows.Count > 0)
                {
                    for (int i = 0; i < tblReporteEspesores.Rows.Count; i++)
                    {
                        listaReportes.Add(new Reporte
                        {
                            NumeroControl = tblReporteEspesores.Rows[i]["NumeroControl"].ToString(),
                            NumeroReporte = tblReporteEspesores.Rows[i]["NumeroReporte"].ToString()
                        });
                    }
                }
                return listaReportes;
            }
            catch (Exception)
            {
                return null;
            }
        }


        ///////////////////////////////////////////TERMINA SEGMENTO DE STORED PROCEDURES /////////////////////////////////////////////////////////

        public bool ExisteArchivosTraveler(string OrdenTrabajo, string rutaTraveler)
        {
            bool existe = false;
            try
            {
                //string rutaTraveler = ConfigurationManager.AppSettings["RutaTraveler"].ToString();
                string UsuarioTraveler  = ConfigurationManager.AppSettings["UsuarioTraveler"].ToString();
                string PassTraveler = ConfigurationManager.AppSettings["PassTraveler"].ToString();
                string traveler = Path.Combine(rutaTraveler, "ODT " + OrdenTrabajo + ".pdf");
                //using (new NetworkConnection(rutaTraveler, new NetworkCredential(UsuarioTraveler, PassTraveler)))
                //{
                if (File.Exists(traveler))
                {
                        existe = true;
                }
                else
                {
                    EscribirEnCSV(OrdenTrabajo, "Traveler", "No se encontro archivo Traveler");
                }
                //}
                return existe;
            }
            catch (Exception e)
            {
                return false;
            }
        }
        public int getCantidadPaginas(string OrdenTrabajo, string RutaTraveler)
        {
            //string ruta = Path.Combine(ConfigurationManager.AppSettings["RutaTraveler"].ToString(), "ODT " + OrdenTrabajo + ".pdf");
            string ruta = Path.Combine(RutaTraveler, "ODT " + OrdenTrabajo + ".pdf");
            using (PdfReader reader = new PdfReader(ruta))
            {
                return reader.NumberOfPages;
            }
        }

        //TRAVELERS        
        public void TravelerXSpool(Spool Spool, string RutaTraveler)
        {            
            try
            {
                PdfReader reader = null;
                PdfCopy copia = null;
                Document doc = new Document();                
                string rutaDestino = Path.GetTempPath();                
                if (ExisteArchivosTraveler(Spool.OrdenTrabajo, RutaTraveler))
                {                    
                    Spool.NumeroPaginas = getCantidadPaginas(Spool.OrdenTrabajo, RutaTraveler) - Spool.NumeroPaginas; //VERIFICAR SI EL ARCHIVO DE TRAVELER NO TRAE MENOS HOJAS Y SI ES NEGATIVO                    
                    if(Spool.NumeroPaginas > 0)
                    {
                        if (!Spool.AplicaGranel)
                        {
                            reader = new PdfReader(Path.Combine(RutaTraveler, "ODT " + Spool.OrdenTrabajo + ".pdf"));
                            if (reader != null)
                            {                                
                                if (reader.NumberOfPages >= Spool.NumeroPaginas)
                                {
                                    copia = new PdfCopy(doc, new FileStream(rutaDestino + "\\" + Spool.NumeroControl + "_Traveler.pdf", FileMode.Create));
                                    doc.Open();
                                    copia.AddPage(copia.GetImportedPage(reader, Spool.NumeroPaginas));
                                    copia.Close();
                                    doc.Close();
                                }
                                else
                                {
                                    EscribirEnCSV(Spool.NumeroControl, "Traveler", "No coinciden los numeros de pagina");
                                }
                            }
                            else
                            {
                                EscribirEnCSV(Spool.NumeroControl, "Traveler", "No se encontro traveler");
                            }
                            if (reader != null)
                                reader.Close();
                        }
                        else
                        {
                            EscribirEnCSV(Spool.NumeroControl, "Traveler", "Es Granel");
                        }                        
                    }
                    else
                    {
                        EscribirEnCSV(Spool.NumeroControl, "Traveler", "Probablemente el archivo este dañado o es granel");
                    }                    
                }
                else
                {
                    EscribirEnCSV(Spool.OrdenTrabajo, "Traveler", "No se encontro archivo principal traveler");
                }                
            }
            catch (Exception e)
            {
                EscribirEnCSV("Traveler", "Traveler", "Error: " + e.Message);                
            }            
        }
        //CERTIFICADOS
        public void CertificadoXSpool(string NumeroControl, string RutaReportes)
        {
            try
            {
                PdfReader reader = null;
                PdfCopy copia = null;
                Document doc = new Document();                
                string carpetaCertificados = ConfigurationManager.AppSettings["CarpetaCertificadoSpool"].ToString();                
                RutaReportes = String.Concat(RutaReportes, carpetaCertificados);
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(RutaReportes))
                {
                    if(File.Exists(Path.Combine(RutaReportes, NumeroControl + ".pdf")))
                    {
                        reader = new PdfReader(Path.Combine(RutaReportes, NumeroControl + ".pdf"));
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
        public void RTPOST_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportesPND_PorNumeroControl(ProyectoID, NumeroControl, 5);
                string carpetaRTPOST = ConfigurationManager.AppSettings["CarpetaRTPost"].ToString();                
                string rutaRTPOST = String.Concat(RutaReportes, carpetaRTPOST);
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
                            if (File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_RTPOST.pdf"))
                            {
                                listaRTPOST.Add(rutaTemp + "\\" + item.NumeroReporte + "_RTPOST.pdf");
                            }
                        }
                        if (listaRTPOST != null && listaRTPOST.Count > 0)
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
        public void PTPOST_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportesPND_PorNumeroControl(ProyectoID, NumeroControl, 6);
                string carpetaPTPOST = ConfigurationManager.AppSettings["CarpetaPTPost"].ToString();
                string rutaPTPOST = String.Concat(RutaReportes, carpetaPTPOST);            
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
        public void UTPOST_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportesPND_PorNumeroControl(ProyectoID, NumeroControl, 14);
                string carpetaUTPOST = ConfigurationManager.AppSettings["CarpetaUTPost"].ToString();
                string rutaUTPOST = String.Concat(RutaReportes, carpetaUTPOST);
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
        
        public void RT_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportesPND_PorNumeroControl(ProyectoID, NumeroControl, 1);
                string carpetaRT = ConfigurationManager.AppSettings["CarpetaRT"].ToString();
                string rutaRT = String.Concat(RutaReportes, carpetaRT);
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
        public void PT_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportesPND_PorNumeroControl(ProyectoID, NumeroControl, 2);
                string carpetaPT = ConfigurationManager.AppSettings["CarpetaPT"].ToString();
                string rutaPT = String.Concat(RutaReportes, carpetaPT);
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
        public void UT_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportesPND_PorNumeroControl(ProyectoID, NumeroControl, 8);
                string carpetaUT = ConfigurationManager.AppSettings["CarpetaUT"].ToString();
                string rutaUT = String.Concat(RutaReportes, carpetaUT);
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
        public void PMI_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportesPND_PorNumeroControl(ProyectoID, NumeroControl, 10);
                string carpetaPMI = ConfigurationManager.AppSettings["CarpetaPMI"].ToString();
                string rutaPMI = String.Concat(RutaReportes, carpetaPMI);
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
        public void Ferrita_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportesPND_PorNumeroControl(ProyectoID, NumeroControl, 15);
                string carpetaFerrita = ConfigurationManager.AppSettings["CarpetaFerrita"].ToString();
                string rutaFerrita = String.Concat(RutaReportes, carpetaFerrita);
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaFerrita))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaFerrita, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaFerrita, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_Ferrita.pdf", FileMode.Create));
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
                                        EscribirEnCSV(NumeroControl, "Reporte Ferritas", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte Ferritas", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte Ferritas", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaFerrita = new List<string>();
                        foreach (var item in Lista)
                        {
                            if (File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_Ferrita.pdf"))
                            {
                                listaFerrita.Add(rutaTemp + "\\" + item.NumeroReporte + "_Ferrita.pdf");
                            }
                        }
                        if (listaFerrita != null && listaFerrita.Count > 0)
                        {
                            MergePDF(listaFerrita, rutaTemp + "\\" + NumeroControl + "_Ferrita.pdf");
                            eliminarArchivosPND(listaFerrita);
                        }
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte Ferritas", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte Ferritas", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte Ferritas", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }
        //REPORTES TT
        public void PWHT_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {                
                List<Reporte > Lista = ObtenerReportesTT_PorNumeroControl(ProyectoID, NumeroControl, 3); //REPORTES TT
                string carpetaPWHT = ConfigurationManager.AppSettings["CarpetaPWHT"].ToString();
                string rutaPWHT = String.Concat(RutaReportes, carpetaPWHT);
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

        public void Preheat_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte> Lista = ObtenerReportesTT_PorNumeroControl(ProyectoID, NumeroControl, 7); //REPORTES TT
                string carpetaPreheat = ConfigurationManager.AppSettings["CarpetaPreheat"].ToString();
                string rutaPreheat = String.Concat(RutaReportes, carpetaPreheat);
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaPreheat))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaPreheat, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaPreheat, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_Preheat.pdf", FileMode.Create));
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
                                        EscribirEnCSV(NumeroControl, "Reporte Preheat", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte Preheat", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte Preheat", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaPreheat = new List<string>();
                        foreach (var item in Lista)
                        {
                            if (File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_Preheat.pdf"))
                            {
                                listaPreheat.Add(rutaTemp + "\\" + item.NumeroReporte + "_Preheat.pdf");
                            }
                        }
                        if (listaPreheat != null && listaPreheat.Count > 0)
                        {
                            MergePDF(listaPreheat, rutaTemp + "\\" + NumeroControl + "_Preheat.pdf");
                            eliminarArchivosPND(listaPreheat);
                        }
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte Preheat", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte Preheat", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte Preheat", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }


        public void HTPOST_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportesTT_PorNumeroControl(ProyectoID, NumeroControl, 16);
                string carpetaHTPOST = ConfigurationManager.AppSettings["CarpetaHTPOST"].ToString();
                string rutaHTPOST = String.Concat(RutaReportes, carpetaHTPOST);
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaHTPOST))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaHTPOST, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaHTPOST, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_HTPOST.pdf", FileMode.Create));
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
                                        EscribirEnCSV(NumeroControl, "Reporte HTPOST", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte HTPOST", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte HTPOST", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaHTPOST = new List<string>();
                        foreach (var item in Lista)
                        {
                            if (File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_HTPOST.pdf"))
                            {
                                listaHTPOST.Add(rutaTemp + "\\" + item.NumeroReporte + "_HTPOST.pdf");
                            }
                        }
                        if (listaHTPOST != null && listaHTPOST.Count > 0)
                        {
                            MergePDF(listaHTPOST, rutaTemp + "\\" + NumeroControl + "_HTPOST.pdf");
                            eliminarArchivosPND(listaHTPOST);
                        }
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte HTPOST", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte HTPOST", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte HTPOST", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }

        public void Durezas_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportesTT_PorNumeroControl(ProyectoID, NumeroControl, 4);
                string carpetaDurezas = ConfigurationManager.AppSettings["CarpetaDurezas"].ToString();
                string rutaDurezas = String.Concat(RutaReportes, carpetaDurezas);
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaDurezas))
                {
                    if (Lista != null && Lista.Count > 0)
                    {
                        foreach (var item in Lista)
                        {
                            if (File.Exists(Path.Combine(rutaDurezas, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaDurezas, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_Durezas.pdf", FileMode.Create));
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
                                        EscribirEnCSV(NumeroControl, "Reporte Durezas", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte Durezas", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte Durezas", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaDurezas = new List<string>();
                        foreach (var item in Lista)
                        {
                            if (File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_Durezas.pdf"))
                            {
                                listaDurezas.Add(rutaTemp + "\\" + item.NumeroReporte + "_Durezas.pdf");
                            }
                        }
                        if (listaDurezas != null && listaDurezas.Count > 0)
                        {
                            MergePDF(listaDurezas, rutaTemp + "\\" + NumeroControl + "_Durezas.pdf");
                            eliminarArchivosPND(listaDurezas);
                        }
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte Durezas", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte Durezas", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte Durezas", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }


        //PINTURA
        public void SandBlast_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportePintura_PorNumeroControl(ProyectoID, NumeroControl, 1);
                string carpetaSandblast = ConfigurationManager.AppSettings["CarpetaSandblast"].ToString();
                string rutaSanBlast = String.Concat(RutaReportes, carpetaSandblast);
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
                            if (File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_SandBlast.pdf"))
                            {
                                listaSandBlast.Add(rutaTemp + "\\" + item.NumeroReporte + "_SandBlast.pdf");
                            }
                        }
                        if (listaSandBlast != null && listaSandBlast.Count > 0)
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
        public void Primario_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportePintura_PorNumeroControl(ProyectoID, NumeroControl, 2);
                string carpetaPrimarios = ConfigurationManager.AppSettings["CarpetaPrimarios"].ToString();
                string rutaPrimario = String.Concat(RutaReportes, carpetaPrimarios);
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
        public void Intermedio_X_Spool(int ProyectoID, string NumeroControl,string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportePintura_PorNumeroControl(ProyectoID, NumeroControl, 3);
                string carpetaIntermedios = ConfigurationManager.AppSettings["CarpetaIntermedios"].ToString();
                string rutaIntermedios = String.Concat(RutaReportes, carpetaIntermedios);
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
        public void Acabado_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportePintura_PorNumeroControl(ProyectoID, NumeroControl, 4);
                string carpetaAcabado = ConfigurationManager.AppSettings["CarpetaAcabado"].ToString();
                string rutaAcabado = String.Concat(RutaReportes, carpetaAcabado);
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
        public void Adherencia_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte > Lista = ObtenerReportePintura_PorNumeroControl(ProyectoID, NumeroControl, 5);
                string carpetaAdherencia = ConfigurationManager.AppSettings["CarpetaAdherencia"].ToString();
                string rutaAdherencia = String.Concat(RutaReportes, carpetaAdherencia);
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
        
        public List<string> ObtenerReportesPullOf_Holiday(int ProyectoID, DataTable NumeroControl, string RutaReportes, int Reporte, string NumeroEmbarque)
        {
            try
            {
                List<string> retorno = new List<string>();
                List<NumeroReportePullHoliday> ListaNumReporte = ObtenerReportePullOf_Holiday(ProyectoID, NumeroControl, Reporte);
                string carpetaPullOf = ConfigurationManager.AppSettings["CarpetaPullOf"].ToString();
                string carpetaHoliday = ConfigurationManager.AppSettings["CarpetaHoliday"].ToString();
                string rutaPullof = String.Concat(RutaReportes, carpetaPullOf);
                string rutaHoliday = String.Concat(RutaReportes, carpetaHoliday);
                string rutaTemp = Path.GetTempPath();
                if(Reporte == 1) //PULLOF
                {
                    if (Directory.Exists(rutaPullof))
                    {
                        if (ListaNumReporte != null && ListaNumReporte.Count > 0)
                        {
                            foreach (var item in ListaNumReporte)
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
                                            EscribirEnCSV(NumeroEmbarque, "Reporte PullOf", "Paginas no encontradas");
                                        }
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroEmbarque, "Reporte PullOf", "Paginas no encontradas");
                                    }
                                    reader.Close();
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroEmbarque, "Reporte PullOf", "No se encontro Reporte");
                                }
                            }                          
                        }
                        else
                        {
                            EscribirEnCSV(NumeroEmbarque, "Reporte PullOf", "No se encontro ningun Reporte");
                        }
                    }
                    else
                    {
                        EscribirEnCSV(NumeroEmbarque, "Reporte PullOf", "No se Encontro carpeta de reportes");
                    }
                }
                else // HOLIDAY
                {
                    if (Directory.Exists(rutaHoliday))
                    {
                        if (ListaNumReporte != null && ListaNumReporte.Count > 0)
                        {
                            foreach (var item in ListaNumReporte)
                            {
                                if (File.Exists(Path.Combine(rutaHoliday, item.NumeroReporte + ".pdf")))
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
                                            EscribirEnCSV(NumeroEmbarque, "Reporte Holiday", "Paginas no encontradas");
                                        }
                                    }
                                    else
                                    {
                                        EscribirEnCSV(NumeroEmbarque, "Reporte Holiday", "Paginas no encontradas");
                                    }
                                    reader.Close();
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroEmbarque, "Reporte Holiday", "No se encontro Reporte");
                                }
                            }                           
                        }
                        else
                        {
                            EscribirEnCSV(NumeroEmbarque, "Reporte Holiday", "No se encontro ningun Reporte");
                        }
                    }
                    else
                    {
                        EscribirEnCSV(NumeroEmbarque, "Reporte Holiday", "No se Encontro carpeta de reportes");
                    }
                }
                if(ListaNumReporte.Count > 0)
                {
                    foreach(var item in ListaNumReporte)
                    {
                        retorno.Add(item.NumeroReporte);
                    }
                }
                return retorno;
            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroEmbarque, "Reporte PullOf-Holiday", "Error Obteniendo Numeros de Reporte: " + e.Message);
                return null;
            }            
        }

        //public void Dimensional_X_Spool(Proyecto Proyecto, string NumeroControl, string RutaReportes)
        //{
        //    try
        //    {
        //        PdfReader reader = null;
        //        PdfCopy copia = null;
        //        Document doc = new Document();
        //        string carpetaDimensional = ConfigurationManager.AppSettings["CarpetaDimensional"].ToString();
        //        RutaReportes = String.Concat(RutaReportes, carpetaDimensional);
        //        string rutaTemp = Path.GetTempPath();
        //        if (Directory.Exists(RutaReportes))
        //        {
        //            if (File.Exists(Path.Combine(RutaReportes, Proyecto.FolioDimensional + NumeroControl + ".pdf")))
        //            {
        //                reader = new PdfReader(Path.Combine(RutaReportes, Proyecto.FolioDimensional + NumeroControl + ".pdf"));
        //                if (reader != null)
        //                {
        //                    if (reader.NumberOfPages >= 1)
        //                    {
        //                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + Proyecto.FolioDimensional + NumeroControl + "_Dimensional.pdf", FileMode.Create));
        //                        doc.Open();
        //                        for (int i = 1; i <= reader.NumberOfPages; i++)
        //                        {
        //                            copia.AddPage(copia.GetImportedPage(reader, i));
        //                        }
        //                        copia.Close();
        //                        doc.Close();
        //                    }
        //                    else
        //                    {
        //                        EscribirEnCSV(NumeroControl, "Dimensional", "Pagina no encontrada");
        //                    }
        //                }
        //                else
        //                {
        //                    EscribirEnCSV(NumeroControl, "Dimensional", "Dimensional no encontrado");
        //                }
        //                reader.Close();
        //            }
        //            else
        //            {
        //                EscribirEnCSV(NumeroControl, "Dimensional", "Dimensional no encontrado");
        //            }

        //        }
        //        else
        //        {
        //            EscribirEnCSV(NumeroControl, "Dimensional", "Carpeta de Dimensional no encontrada");
        //        }

        //    }
        //    catch (Exception e)
        //    {
        //        EscribirEnCSV(NumeroControl, "Dimensional", "Error Obteniendo Dimensional: " + e.Message);
        //    }
        //}

        public void Dimensional_X_Spool(Proyecto Proyecto, string NumeroControl, string RutaReportes, string Dimensional)
        {
            try
            {
                PdfReader reader = null;
                PdfCopy copia = null;
                Document doc = new Document();
                string carpetaDimensional = ConfigurationManager.AppSettings["CarpetaDimensional"].ToString();
                RutaReportes = String.Concat(RutaReportes, carpetaDimensional);
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(RutaReportes))
                {
                    if (File.Exists(Path.Combine(RutaReportes, Proyecto.FolioDimensional + NumeroControl + ".pdf"))) //VERIFICA SI EXISTE CON NUMERO DE CONTROL
                    {
                        reader = new PdfReader(Path.Combine(RutaReportes, Proyecto.FolioDimensional + NumeroControl + ".pdf"));
                        if (reader != null)
                        {
                            if (reader.NumberOfPages >= 1)
                            {
                                copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + Proyecto.FolioDimensional + NumeroControl + "_Dimensional.pdf", FileMode.Create));
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
                                EscribirEnCSV(NumeroControl, "Dimensional", "Pagina no encontrada");
                            }
                        }
                        else
                        {
                            EscribirEnCSV(NumeroControl, "Dimensional", "Dimensional no encontrado");
                        }
                        reader.Close();
                    }
                    else if (File.Exists(Path.Combine(RutaReportes, Dimensional + ".pdf"))) //VERIFICA SI EXISTE CON NUMERO DE CONTROL
                    {
                        reader = new PdfReader(Path.Combine(RutaReportes, Dimensional + ".pdf"));
                        if (reader != null)
                        {
                            if (reader.NumberOfPages >= 1)
                            {
                                copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + Dimensional + "_Dimensional.pdf", FileMode.Create));
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
                                EscribirEnCSV(NumeroControl, "Dimensional", "Pagina no encontrada");
                            }
                        }
                        else
                        {
                            EscribirEnCSV(NumeroControl, "Dimensional", "Dimensional no encontrado");
                        }
                        reader.Close();
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Dimensional", "Dimensional no encontrado");
                    }


                    //else
                    //{
                    //    EscribirEnCSV(NumeroControl, "Dimensional", "Dimensional no encontrado");
                    //}

                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Dimensional", "Carpeta de Dimensional no encontrada");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Dimensional", "Error Obteniendo Dimensional: " + e.Message);
            }
        }

        //ESPESORES
        public void Espesores_X_Spool(int ProyectoID, string NumeroControl, string RutaReportes)
        {
            try
            {
                List<Reporte> ListaReportes = ObtenerReporteEspesores(ProyectoID, NumeroControl);
                string carpetaEspesores = ConfigurationManager.AppSettings["CarpetaEspesores"].ToString();
                string rutaEspesores = String.Concat(RutaReportes, carpetaEspesores);
                string rutaTemp = Path.GetTempPath();
                if (Directory.Exists(rutaEspesores))
                {
                    if (ListaReportes != null && ListaReportes.Count > 0)
                    {
                        foreach (var item in ListaReportes)
                        {
                            if (File.Exists(Path.Combine(rutaEspesores, item.NumeroReporte + ".pdf")))
                            {
                                PdfCopy copia = null;
                                PdfReader reader = null;
                                Document doc = new Document();
                                reader = new PdfReader(Path.Combine(rutaEspesores, item.NumeroReporte + ".pdf"));
                                if (reader != null)
                                {
                                    if (reader.NumberOfPages > 0)
                                    {
                                        copia = new PdfCopy(doc, new FileStream(rutaTemp + "\\" + item.NumeroReporte + "_Espesores.pdf", FileMode.Create));
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
                                        EscribirEnCSV(NumeroControl, "Reporte Espesores", "Paginas no encontradas");
                                    }
                                }
                                else
                                {
                                    EscribirEnCSV(NumeroControl, "Reporte Espesores", "Paginas no encontradas");
                                }
                                reader.Close();
                            }
                            else
                            {
                                EscribirEnCSV(NumeroControl, "Reporte Espesores", "No se encontro Reporte");
                            }
                        }
                        //Mezclar reportes por spool
                        List<string> listaEspesores = new List<string>();
                        foreach (var item in ListaReportes)
                        {
                            if (File.Exists(rutaTemp + "\\" + item.NumeroReporte + "_Espesores.pdf"))
                            {
                                listaEspesores.Add(rutaTemp + "\\" + item.NumeroReporte + "_Espesores.pdf");
                            }
                        }
                        if (listaEspesores != null && listaEspesores.Count > 0)
                        {
                            MergePDF(listaEspesores, rutaTemp + "\\" + NumeroControl + "_Espesores.pdf");
                            eliminarArchivosPND(listaEspesores);
                        }
                    }
                    else
                    {
                        EscribirEnCSV(NumeroControl, "Reporte Espesores", "No se encontro ningun Reporte");
                    }
                }
                else
                {
                    EscribirEnCSV(NumeroControl, "Reporte Espesores", "No se Encontro carpeta de reportes");
                }

            }
            catch (Exception e)
            {
                EscribirEnCSV(NumeroControl, "Reporte Espesores", "Error Obteniendo Numeros de Reporte: " + e.Message);
            }
        }




        //public bool GenerarParticionamiento(List<Spool> Spools, string NumeroEmbarque, string rutaDestino)
        public bool GenerarParticionamiento(List<Spool> Spools, string NumeroEmbarque, string rutaDestino, List<List<string>> NumeroReportes)
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
                        suma += Decimal.Parse(ConvertToMegabytes(UInt64.Parse((new FileInfo(rutaTemp + "\\" + item.NumeroControl +  ".pdf")).Length.ToString())));
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
                            suma += Decimal.Parse(ConvertToMegabytes(UInt64.Parse((new FileInfo(rutaTemp + "\\" + item.NumeroControl +  ".pdf")).Length.ToString())));
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
                            //GenerarArchivoGeneralParticionado(SpoolParticion, NumeroEmbarque, rutaDestino);
                            GenerarArchivoGeneralParticionado(SpoolParticion, NumeroEmbarque, rutaDestino, NumeroReportes);
                            estatus = true;
                        }
                        else
                        {
                            estatus = false;
                        }
                    }
                    else
                    {
                        //GenerarArchivoGeneralParticionado(SpoolParticion, NumeroEmbarque, rutaDestino);
                        GenerarArchivoGeneralParticionado(SpoolParticion, NumeroEmbarque, rutaDestino, NumeroReportes);
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


        public void GenerarArchivoGeneralParticionado(List<SpoolParticion> ListaSpool, string NumeroEmbarque, string rutaDestino, List<List<string>> ListaNumReporte)
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
                    //AGREGO LOS NUMEROS DE REPORTE DE PULLOF Y HOLIDAY
                    if(ListaNumReporte.Count > 0)
                    {
                        for(int i = 0; i < ListaNumReporte.Count; i++)
                        {
                            for(int j = 0; j < ListaNumReporte[i].Count; j++)
                            {
                                if (File.Exists(Path.Combine(rutaTemp, ListaNumReporte[i][j].ToString() + ".pdf")))
                                    rutasArchivos.Add(Path.Combine(rutaTemp, ListaNumReporte[i][j].ToString() + ".pdf"));
                            }
                        }
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
                        //AGREGO LOS NUMEROS DE REPORTE DE PULLOF Y HOLIDAY
                        if (ListaNumReporte.Count > 0)
                        {
                            for (int i = 0; i < ListaNumReporte.Count; i++)
                            {
                                for (int j = 0; j < ListaNumReporte[i].Count; j++)
                                {
                                    if (File.Exists(Path.Combine(rutaTemp, ListaNumReporte[i][j].ToString() + ".pdf")))
                                        rutasArchivos.Add(Path.Combine(rutaTemp, ListaNumReporte[i][j].ToString() + ".pdf"));
                                }
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
                // ELIMINA ARCHIVOS TMP DE NUMEROS DE REPORTE
                if(ListaNumReporte.Count > 0)
                {
                    for(int i = 0; i < ListaNumReporte.Count; i++)
                    {
                        for(int j = 0; j < ListaNumReporte[i].Count; j++)
                        {
                            if(File.Exists(rutaTemp + "\\" + ListaNumReporte[i][j].ToString() + ".pdf"))
                            {
                                File.SetAttributes(rutaTemp + "\\" + ListaNumReporte[i][j].ToString() + ".pdf", FileAttributes.Normal);
                                File.Delete(rutaTemp + "\\" + ListaNumReporte[i][j].ToString() + ".pdf");
                            }
                        }
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

        public void eliminarArchivosPND(string lista)
        {
            try
            {               
                if (File.Exists(lista))
                {
                    File.SetAttributes(lista, FileAttributes.Normal);
                    File.Delete(lista);
                }                
            }
            catch (Exception e)
            {
                throw;
            }
        }

        //public void UnirArchivos(List<Spool> ListaSpool)
        public void UnirArchivos(string FolioDimensional, List<Spool> ListaSpool, List<List<string>> Listas)
        {
            try
            {                
                string rutaTemp = Path.GetTempPath();
                string extTraveler = "_Traveler.pdf",
                    extCertificado = "_Certificado.pdf",
                    extRTPOST = "_RTPOST.pdf",
                    extPTPOST = "_PTPOST.pdf",
                    extUTPOST = "_UTPOST.pdf",
                    extHTPOST = "_HTPOST.pdf",
                    extRT = "_RT.pdf",
                    extPT = "_PT.pdf",
                    extUT = "_UT.pdf",
                    extPMI = "_PMI.pdf", extPWHT = "_PWHT.pdf", extSandBlast = "_SandBlast.pdf", extPrimario = "_Primario.pdf",
                    extIntermedio = "_Intermedio.pdf", extAcabado = "_Acabado.pdf", extAdherencia = "_Adherencia.pdf", extPullOf = "_PullOf.pdf",
                    extHoliday = "_Holiday.pdf", extDimensional = "_Dimensional.pdf", extFerrita = "_Ferrita.pdf", extDurezas = "_Durezas.pdf",
                    extEspesores = "_Espesores.pdf", extPreheat = "_Preheat.pdf";
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
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extFerrita)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extFerrita));

                    //TT
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extPWHT)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extPWHT));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extHTPOST)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extHTPOST));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extDurezas)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extDurezas));
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extPreheat)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extPreheat));
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
                    //DIMENSIONAL
                    if (File.Exists(Path.Combine(rutaTemp, FolioDimensional + NumeroControl + extDimensional)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, FolioDimensional + NumeroControl + extDimensional));
                    if (File.Exists(Path.Combine(rutaTemp, item.Dimensional + extDimensional)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, item.Dimensional + extDimensional));
                    //ESPESORES
                    if (File.Exists(Path.Combine(rutaTemp, NumeroControl + extEspesores)))
                        rutasArchivos.Add(Path.Combine(rutaTemp, NumeroControl + extEspesores));                   

                    if(rutasArchivos != null)
                    {
                        MergePDF(rutasArchivos, rutaTemp + "\\" + NumeroControl + ".pdf");                        
                    }                                        
                    rutasArchivos = null;
                }
                //PINTURA -- PULLOF Y HOLIDAY                
                if (Listas != null && Listas.Count > 0)
                {
                    rutasArchivos = new List<string>();
                    for (int i = 0; i < Listas.Count; i++)
                    {
                        for (int j = 0; j < Listas[i].Count; j++)
                        {
                            if (File.Exists(Path.Combine(rutaTemp, Listas[i][j].ToString() + extPullOf)))
                            {
                                rutasArchivos.Add(Path.Combine(rutaTemp, Listas[i][j].ToString() + extPullOf));
                                MergePDF(rutasArchivos, rutaTemp + "\\" + Listas[i][j].ToString() + ".pdf");
                            }                            
                            if (File.Exists(Path.Combine(rutaTemp, Listas[i][j].ToString() + extHoliday)))
                            {
                                rutasArchivos.Add(Path.Combine(rutaTemp, Listas[i][j].ToString() + extHoliday));
                                MergePDF(rutasArchivos, rutaTemp + "\\" + Listas[i][j].ToString() + ".pdf");
                            }                            
                        }
                    }
                    //if (rutasArchivos != null)
                    //{

                    //    MergePDF(rutasArchivos, rutaTemp + "\\" + NumeroControl + ".pdf");
                    //}
                    rutasArchivos = null;
                }
                

                foreach (var item in ListaSpool)
                {
                    // ELIMINA TMP TRAVELER
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extTraveler))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extTraveler, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extTraveler);
                    }
                    // ELIMINA TMP CERTIFICADOS
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extCertificado))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extCertificado, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extCertificado);
                    }
                    // ELIMINA TMP PNDS
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extRTPOST))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extRTPOST, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extRTPOST);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extPTPOST))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extPTPOST, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extPTPOST);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extUTPOST))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extUTPOST, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extUTPOST);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extRT))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extRT, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extRT);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extPT))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extPT, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extPT);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extUT))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extUT, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extUT);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extPMI))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extPMI, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extPMI);
                    }
                    if(File.Exists(rutaTemp + "\\" + item.NumeroControl + extFerrita))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl + extFerrita, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl + extFerrita);
                    }
                    // ELIMINA TMP TT's
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extPWHT))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extPWHT, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extPWHT);
                    }
                    if(File.Exists(rutaTemp + "\\" + item.NumeroControl + extHTPOST))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl + extHTPOST, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl + extHTPOST);
                    }
                    if(File.Exists(rutaTemp + "\\" + item.NumeroControl + extDurezas))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl + extDurezas, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl + extDurezas);
                    }
                    if(File.Exists(rutaTemp + "\\" + item.NumeroControl + extPreheat))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl + extPreheat, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl + extPreheat);
                    }
                    // ELIMINA TMP PINTURA
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extSandBlast))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extSandBlast, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extSandBlast);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extPrimario))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extPrimario, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extPrimario);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extIntermedio))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extIntermedio, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extIntermedio);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extAcabado))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extAcabado, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extAcabado);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extAdherencia))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extAdherencia, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl +  extAdherencia);
                    }
                    //ELIMINA TMP DIMENSIONAL
                    if(File.Exists(rutaTemp + "\\" + FolioDimensional + item.NumeroControl + extDimensional))
                    {
                        File.SetAttributes(rutaTemp + "\\" + FolioDimensional + item.NumeroControl + extDimensional, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + FolioDimensional + item.NumeroControl + extDimensional);
                    }
                    if (File.Exists(rutaTemp + "\\" + item.Dimensional + extDimensional))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.Dimensional + extDimensional, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.Dimensional + extDimensional);
                    }
                    //ELIMINA TMP espesores
                    if (File.Exists(rutaTemp + "\\" + item.NumeroControl + extEspesores))
                    {
                        File.SetAttributes(rutaTemp + "\\" + item.NumeroControl + extEspesores, FileAttributes.Normal);
                        File.Delete(rutaTemp + "\\" + item.NumeroControl + extEspesores);
                    }


                    //if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extPullOf))
                    //{
                    //    File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extPullOf, FileAttributes.Normal);
                    //    File.Delete(rutaTemp + "\\" + item.NumeroControl +  extPullOf);
                    //}
                    //if (File.Exists(rutaTemp + "\\" + item.NumeroControl +  extHoliday))
                    //{
                    //    File.SetAttributes(rutaTemp + "\\" + item.NumeroControl +  extHoliday, FileAttributes.Normal);
                    //    File.Delete(rutaTemp + "\\" + item.NumeroControl +  extHoliday);
                    //}

                }
                //ELIMINAR ARCHIVOS TEMPORALES DE PULLOF Y HOLIDAY
                if (Listas != null && Listas.Count > 0)
                {
                    for (int i = 0; i < Listas.Count; i++)
                    {
                        for (int j = 0; j < Listas[i].Count; j++)
                        {
                            if(File.Exists(rutaTemp + "\\" + Listas[i][j].ToString() + extPullOf))
                            {
                                File.SetAttributes(rutaTemp + "\\" + Listas[i][j].ToString() + extPullOf, FileAttributes.Normal);
                                File.Delete(rutaTemp + "\\" + Listas[i][j].ToString() + extPullOf);
                            }
                            if (File.Exists(rutaTemp + "\\" + Listas[i][j].ToString() + extHoliday))
                            {
                                File.SetAttributes(rutaTemp + "\\" + Listas[i][j].ToString() + extHoliday, FileAttributes.Normal);
                                File.Delete(rutaTemp + "\\" + Listas[i][j].ToString() + extHoliday);
                            }                           
                        }
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
    