using PDFReportes.Models;
using PDFReportes.Utilidades;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PDFReportes
{
    public partial class Form1 : Form
    {
        private static string RutaGuardar = "";
        private static string numEmbarqueGlobal = "";
        private Embarque emb = null;
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
            CmbNumEmbarque.Enabled = false;
            btnSelectPathSave.Enabled = false;
            btnGenerar.Enabled = false;                                           
        }
        public void desbloquearControles()
        {            
            CmbNumEmbarque.Enabled = true;
            btnSelectPathSave.Enabled = true;            
        }

        private void Form1_Load(object sender, EventArgs e)
        {           
            if(Util.Instance.ObtieneEmbarques() != null)
            {
                panelLoading.Visible = false;
                CmbNumEmbarque.DataSource = Util.Instance.ObtieneEmbarques();
            }
            else
            {
                MessageBox.Show("Error de Conexión, verifique las cadenas de conexion en el archivo XML", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                            List<Spool> ListaSpool = Util.Instance.ObtenerNumeroControl(emb.NumeroEmbarque);
                            Util.Instance.crearArchivoCSV(RutaGuardar, emb.NumeroEmbarque);
                            Util.Instance.EscribirEnCSV("NumeroControl", "Reporte", "Error");

                            if (ListaSpool != null)
                            {
                                int numSpool = 0;
                                foreach (var item in ListaSpool)
                                {
                                    Util.Instance.TravelerXSpool(item.OrdenTrabajo, item.OrdenTrabajo + "-" + item.Consecutivo, item.NumeroPaginas, emb.NumeroEmbarque);
                                    Util.Instance.CertificadoXSpool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.RTPOST_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.PTPOST_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.UTPOST_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.RT_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.PT_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.UT_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.PMI_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.PWHT_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.SandBlast_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.Primario_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.Intermedio_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.Acabado_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.Adherencia_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    Util.Instance.PullOf_Holiday_X_Spool(item.OrdenTrabajo + "-" + item.Consecutivo, emb.NumeroEmbarque);
                                    numSpool++;
                                }
                                if (ListaSpool.Count == numSpool)
                                {
                                    Util.Instance.UnirArchivos(ListaSpool);
                                    if (Util.Instance.GenerarParticionamiento(ListaSpool, emb.NumeroEmbarque, RutaGuardar))
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


    }
}
