using PDFReportes.Models;
using PDFReportes.Utilidades;
using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;

namespace PDFReportes
{
    public partial class Form1 : Form
    {
        private static string RutaGuardar = "";
        private static string numEmbarqueGlobal = "";
        private Embarque emb = null;
        private Proyecto proyecto = null;
        private ToolTip tooltip = new ToolTip();
        public Form1()
        {         
            InitializeComponent();                               
        }

        private void btnSelectPathSave_Click(object sender, EventArgs e)
        {
            try
            {                
                if (folderGuardar.ShowDialog() == DialogResult.OK)
                {
                    tooltip.RemoveAll();
                    lblRutaGuardar.Text = "";                        
                    RutaGuardar = "";
                    lblRutaGuardar.Text = folderGuardar.SelectedPath;
                    RutaGuardar = folderGuardar.SelectedPath;
                    //Agregar hover                                        
                    tooltip.ShowAlways = true;                    
                    tooltip.SetToolTip(lblRutaGuardar, folderGuardar.SelectedPath);
                }else
                {
                    lblRutaGuardar.Text = "Ninguna carpeta seleccionada";
                    RutaGuardar = "";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error seleccionando ruta, Error: " + ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClicGenerarPDF(object sender, EventArgs e)
        {

            //panelLoading.Visible = true;
            //panelLoading.BringToFront();
            Form1.CheckForIllegalCrossThreadCalls = false;
            backgroundWorkerGenerarPDF.RunWorkerAsync();            
            //Validacion de numero de embarque            
            //Embarque emb = new Embarque();
            



        }

        private void mayusculaNumEmbarque(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Char.ToUpper(e.KeyChar);
        }        
        public void bloquearControles()
        {
            cmbProyecto.Enabled = false;            
            CmbNumEmbarque.Enabled = false;
            btnSelectPathSave.Enabled = false;
            btnGenerar.Enabled = false;                                           
        }
        public void desbloquearControles()
        {
            cmbProyecto.Enabled = true;
            if(cmbProyecto.Text != "")
            {
                CmbNumEmbarque.Enabled = true;
            }            
            btnSelectPathSave.Enabled = true;            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<Proyecto> listaProyecto = Util.Instance.ObtenerProyectos();
            if ( listaProyecto != null && listaProyecto.Count > 0)
            {
                panelLoading.Visible = false;
                //this.CmbNumEmbarque.Enabled = true;
                cmbProyecto.DataSource = listaProyecto;
                //List<Embarque> listaEmbarque = Util.Instance.ObtieneEmbarques();                
                //if (listaEmbarque != null && listaEmbarque.Count > 0)
                //{                    
                //    CmbNumEmbarque.DataSource = Util.Instance.ObtieneEmbarques();
                //}
                //else
                //{
                //    MessageBox.Show("Imposible obtener embarques, verifica la cadena de conexion a BD", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //}
            }
            else
            {
                this.CmbNumEmbarque.Enabled = false;
                MessageBox.Show("Imposible obtener proyectos, verifique la conexión a BD", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }            
        }

        private void backgroundWorkerGenerarPDF_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            InicioPrograma();
        }

        private void backgroundWorkerGenerarPDF_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            panelLoading.Visible = true;
        }

        private void backgroundWorkerGenerarPDF_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("The task has been cancelled");
            }
            else if (e.Error != null)
            {
                MessageBox.Show("Error. Details: " + (e.Error as Exception).ToString());
            }
            else
            {
                //MessageBox.Show("The task has been completed. Results: " + e.Result.ToString());
                //MessageBox.Show("Se termino el proceso");
                panelLoading.Visible = false;
            }

        }

        private void InicioPrograma()
        {
            proyecto = new Proyecto();
            proyecto = (Proyecto)cmbProyecto.SelectedItem;
            List<List<string>> Listas = new List<List<string>>();
            if (cmbProyecto.Text != "")
            {
                if(proyecto != null && proyecto.ProyectoID != 0)
                {
                    emb = new Embarque();
                    emb = (Embarque)CmbNumEmbarque.SelectedItem;
                    if (CmbNumEmbarque.Text != "")
                    {
                        if (emb != null && emb.EmbarqueID != 0)
                        {
                            numEmbarqueGlobal = emb.NumeroEmbarque;
                            if (RutaGuardar != "")
                            {
                                panelLoading.Visible = true;
                                panelLoading.BringToFront();
                                //backgroundWorkerGenerarPDF.RunWorkerAsync();
                                if (Util.Instance.VerificaSiExistePDF(emb.NumeroEmbarque, RutaGuardar) == "")
                                {
                                    List<Spool> ListaSpool = Util.Instance.ObtenerNumeroControl(proyecto.ProyectoID, emb);
                                    Util.Instance.crearArchivoCSV(RutaGuardar, emb.NumeroEmbarque);
                                    Util.Instance.EscribirEnCSV("NumeroControl", "Reporte", "Error");

                                    if (ListaSpool != null)
                                    {
                                        int numSpool = 0;
                                        List<NumeroControlClass> ListaNumeroControl = new List<NumeroControlClass>();
                                        foreach (var item in ListaSpool)
                                        {
                                            //Util.Instance.TravelerXSpool(item.OrdenTrabajo, item.NumeroControl, item.NumeroPaginas, emb.NumeroEmbarque, proyecto.RutaTraveler);
                                            Util.Instance.TravelerXSpool(item, proyecto.RutaTraveler);
                                            Util.Instance.CertificadoXSpool(item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.RTPOST_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.PTPOST_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.UTPOST_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.RT_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.PT_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.UT_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.PMI_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.Ferrita_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.PWHT_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.HTPOST_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.Durezas_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.SandBlast_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.Primario_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.Intermedio_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.Acabado_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.Adherencia_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            Util.Instance.Dimensional_X_Spool(proyecto, item.NumeroControl, proyecto.RutaReportes);
                                            //Util.Instance.PullOf_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            //Util.Instance.Holiday_X_Spool(proyecto.ProyectoID, item.NumeroControl, proyecto.RutaReportes);
                                            ListaNumeroControl.Add(new NumeroControlClass { NumeroControl = item.NumeroControl });
                                            numSpool++;
                                        }
                                        //OBTENGO LOS REPORTES DE PULLOF Y HOLIDAY YA QUE ESTOS CONTIENEN EL MISMO REPORTE PARA LA MAYORIA DE LOS SPOOLS
                                        if(ListaNumeroControl.Count > 0)
                                        {
                                            DataTable tablaNumControl = new DataTable();
                                            tablaNumControl = ToDataTable.Instance.toDataTable(ListaNumeroControl);
                                            List<string> listaTxtPullOf = new List<string>();
                                            List<string> listaTxtHoliday = new List<string>();                                            
                                            listaTxtPullOf = Util.Instance.ObtenerReportesPullOf_Holiday(proyecto.ProyectoID, tablaNumControl, proyecto.RutaReportes, 1, emb.NumeroEmbarque); //PULLOF
                                            listaTxtHoliday = Util.Instance.ObtenerReportesPullOf_Holiday(proyecto.ProyectoID, tablaNumControl, proyecto.RutaReportes, 2, emb.NumeroEmbarque); //HOLIDAY
                                            Listas.Add(listaTxtPullOf);
                                            Listas.Add(listaTxtHoliday);
                                        }

                                        if (ListaSpool.Count == numSpool)
                                        {
                                            //Util.Instance.UnirArchivos(ListaSpool);
                                            Util.Instance.UnirArchivos(proyecto.FolioDimensional, ListaSpool, Listas);
                                            //if (Util.Instance.GenerarParticionamiento(ListaSpool, emb.NumeroEmbarque, RutaGuardar))
                                            if (Util.Instance.GenerarParticionamiento(ListaSpool, emb.NumeroEmbarque, RutaGuardar, Listas))
                                            {
                                                ////CIERRA CSV
                                                Util.Instance.CerrarArchivoCSV();
                                                MessageBox.Show("Archivo Generado Correctamente en la ruta: " + Environment.NewLine + RutaGuardar,
                                                    "Correcto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                //desbloquearControles();
                                            }
                                            else
                                            {
                                                ////CIERRA CSV
                                                Util.Instance.CerrarArchivoCSV();
                                                MessageBox.Show("No se encontró ningún reporte para el embarque: " + emb.NumeroEmbarque, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                                //desbloquearControles();
                                            }
                                        }
                                        else
                                        {
                                            ////CIERRA CSV
                                            Util.Instance.CerrarArchivoCSV();
                                            MessageBox.Show("Ocurrió un error inesperado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            desbloquearControles();
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Ocurrio Un Error Obteniendo Spool", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        desbloquearControles();
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("Ya existe estos archivos creados: " + Environment.NewLine + Util.Instance.VerificaSiExistePDF(emb.NumeroEmbarque, RutaGuardar), "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    desbloquearControles();
                                }
                            }
                            else
                            {
                                MessageBox.Show("Seleccione la ruta donde guardará el archivo generado", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                lblRutaGuardar.Text = "Ninguna carpeta seleccionada";
                                desbloquearControles();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Número de Embarque Incorrecto", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            CmbNumEmbarque.SelectedIndex = 0;
                            desbloquearControles();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Seleccione Número de Embarque", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        if (CmbNumEmbarque.SelectedIndex >= 0)
                            CmbNumEmbarque.SelectedIndex = 0;
                        desbloquearControles();
                    }

                }
                else
                {
                    MessageBox.Show("Proyecto Incorrecto", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    cmbProyecto.SelectedIndex = 0;
                    desbloquearControles();
                }
            }
            else
            {
                MessageBox.Show("Seleccione Proyecto", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (cmbProyecto.SelectedIndex >= 0)
                    cmbProyecto.SelectedIndex = 0;
                desbloquearControles();
            }                                    
        }

        private void MayusculasProyecto(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Char.ToUpper(e.KeyChar);
        }

        private void cmbProyecto_SelectedIndexChanged(object sender, EventArgs e)
        {            
            if((int)cmbProyecto.SelectedValue != 0)
            {
                CmbNumEmbarque.Enabled = true;
                CmbNumEmbarque.DataSource = Util.Instance.ObtieneEmbarques((int)cmbProyecto.SelectedValue);
            }
            else
            {
                CmbNumEmbarque.Enabled = false;
            }
        }
        
    }
}
