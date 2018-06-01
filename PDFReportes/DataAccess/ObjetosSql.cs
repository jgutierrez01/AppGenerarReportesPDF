using ConnectorLibrary.Constantes;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccess.Sql
{
    public class ObjetosSQL
    {
        /// <summary>
        /// Cadena Coneccion a la BD
        /// </summary>
        /// <returns></returns>
        protected SqlConnection Conexion()
        {
            return new SqlConnection(ConfigurationManager.ConnectionStrings["Sam2"].ConnectionString);
         
        }
        
        /// <summary>
        /// Retorna un DatatTable con la información de BD
        /// </summary>
        /// <param name="Stord">Stord a ejecutar</param>
        /// <param name="Parametros">Parametros que requiere el stord</param>
        /// <returns>Objeto DatatTable con la coleccion de datos</returns>
        public DataTable Tabla(string Stord, string[,] Parametros = null)
        {

            DataTable dt = new DataTable();
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.Connection.Open();
                        da.Fill(dt);
                        cmd.Connection.Close();
                    }
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }
                return dt;
            }

        }
        /// <summary>
        /// Retorna un coleccion de Tablas con la información de BD
        /// </summary>
        /// <param name="Stord">nombre del stord a ejecutar</param>
        /// <param name="Parametros">Parametros que requiere el stord</param>
        /// <returns>Coleccion de tablas</returns>
        public DataSet Coleccion(string Stord, string[,] Parametros = null)
        {
            DataSet ds = new DataSet();
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.Connection.Open();
                        da.Fill(ds);
                        cmd.Connection.Close();
                    }
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }
                return ds;
            }


        }
        /// <summary>
        /// Retorna un coleccion de Tablas con la información de BD
        /// </summary>
        /// <param name="Stord">Nombre del Stord a ejecutar</param>
        /// <param name="TablaSube">objeto DataTable que envía al stord</param>
        /// <param name="NombreTabla">nombre del parametro de tabla</param>
        /// <param name="Parametros">Parametros que requiere el stord</param>
        /// <returns>Coleccion de tablas</returns>
        public DataSet Coleccion(string Stord, DataTable TablaSube, String NombreTabla, string[,] Parametros = null)
        {
            DataSet ds = new DataSet();
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.Parameters.Add(new SqlParameter(NombreTabla, TablaSube));
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.Connection.Open();
                        da.Fill(ds);
                        cmd.Connection.Close();
                    }
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }
                return ds;
            }


        }
        /// <summary>
        /// Ejecuta un stord en la BD
        /// </summary>
        /// <param name="Stord">Nombre del Stord a ejecutar</param>
        /// <param name="Parametros">Parametros que requiere el stord</param>
        /// <returns>True si se ejecuto el stord, Exception: error</returns>
        public bool Ejecuta(string Stord, string[,] Parametros = null)
        {
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    cmd.ExecuteNonQuery();
                    cmd.Connection.Close();
                    return true;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }
            }
        }

        /// <summary>
        /// Ejecuta un stord en la BD
        /// </summary>
        /// <param name="Stord">Nombre del Stord a ejecutar</param>
        /// <param name="Parametros">Parametros que requiere el stord</param>
        /// <returns>regresa el numero de filas afectadas.</returns>
        public int EjecutaStoreExecuteReader(string Stord, string[,] Parametros = null)
        {
            int rowsAfected = 0;
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    rowsAfected = cmd.ExecuteNonQuery();
                    cmd.Connection.Close();
                    return rowsAfected;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    return rowsAfected;
                }
            }
        }
        /// <summary>
        /// Ejecuta un stord en la BD
        /// </summary>
        /// <param name="Stord">Nombre del Stord a ejecutar</param>
        /// <param name="TablaSube">objeto DataTable que envía al stord</param>
        /// <param name="NombreTabla">nombre del parametro de tabla</param>
        /// <param name="Parametros">Parametros que requiere el stord</param>
        /// <returns>Objeto DatatTable con la coleccion de datos</returns>
        public bool Ejecuta(string Stord, DataTable TablaSube, String NombreTabla, string[,] Parametros = null)
        {
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.Parameters.Add(new SqlParameter(NombreTabla, TablaSube));
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    cmd.ExecuteNonQuery();
                    cmd.Connection.Close();
                    return true;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }
            }
        }

        public int EjecutaInsertUpdate(string Stord, DataTable TablaSube, String NombreTabla, string[,] Parametros = null)
        {
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                int lastRow = 0;
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.Parameters.Add(new SqlParameter(NombreTabla, TablaSube));
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    lastRow = int.Parse(cmd.ExecuteScalar().ToString());
                    cmd.Connection.Close();
                    return lastRow;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }
            }
        }

        public int EjecutaInsertUpdate(string Stord, DataTable TablaSube, String NombreTabla, DataTable TablaSube2, String NombreTabla2, string[,] Parametros = null)
        {
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                int lastRow = 0;
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.Parameters.Add(new SqlParameter(NombreTabla, TablaSube));
                cmd.Parameters.Add(new SqlParameter(NombreTabla2, TablaSube2));
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    lastRow = int.Parse(cmd.ExecuteScalar().ToString());
                    cmd.Connection.Close();
                    return lastRow;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }
            }
        }

        public DataTable Tabla(string Stord, DataTable TablaSube, String NombreTabla, string[,] Parametros = null)
        {
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                DataTable dt = new DataTable();
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.Parameters.Add(new SqlParameter(NombreTabla, TablaSube));
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.Connection.Open();
                        da.Fill(dt);
                        cmd.Connection.Close();
                    }
                    return dt;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }
            }
        }

        /// <summary>
        /// Ejecuta un stord en la BD
        /// </summary>
        /// <param name="Stord">Nombre del Stord a ejecutar</param>
        /// <param name="TablaSube">objeto DataTable que envía al stord</param>
        /// <param name="NombreTabla">nombre del parametro de tabla</param>
        /// <param name="Parametros">Parametros que requiere el stord</param>
        /// <returns>Regresa un objeto Datatable</returns>
        /// 
        public DataTable EjecutaDataAdapter(string Stord, DataTable TablaSube, String NombreTabla, string[,] Parametros = null)
        {
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                DataTable dt = new DataTable();
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.Parameters.Add(new SqlParameter(NombreTabla, TablaSube));
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.Connection.Open();
                        da.Fill(dt);
                        cmd.Connection.Close();
                    }


                    return dt;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();

                    DataTable dtError = new DataTable("error");
                    DataColumn dataColumn = null;
                    for (int i = 0; i < 2; i++)
                    {
                        dataColumn = new DataColumn(i.ToString());
                        dtError.Columns.Add(dataColumn);
                    }

                    DataRow row = dtError.NewRow();
                    row["0"] = e.Message;
                    row["1"] = e.Message;
                    dtError.Rows.Add(row);
                    return dtError;
                }
            }
        }
        /// <summary>
        /// Ejecuta un stord en la BD
        /// </summary>
        /// <param name="Stord">Nombre del Stord a ejecutar</param>
        /// <param name="TablaSube">objeto DataTable que envía al stord</param>
        /// <param name="NombreTabla">nombre del parametro de tabla</param>
        /// <param name="Parametros">Parametros que requiere el stord</param>
        /// <returns>Regresa un objeto Datatable</returns>
        /// 
        public DataTable EjecutaDataAdapter(string Stord, string[,] Parametros = null)
        {
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                DataTable dt = new DataTable();
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());

                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                    {
                        cmd.CommandTimeout = 0;
                        cmd.Connection.Open();
                        da.Fill(dt);
                        cmd.Connection.Close();
                    }


                    return dt;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();

                    DataTable dtError = new DataTable("error");
                    DataColumn dataColumn = null;
                    for (int i = 0; i < 2; i++)
                    {
                        dataColumn = new DataColumn(i.ToString());
                        dtError.Columns.Add(dataColumn);
                    }

                    DataRow row = dtError.NewRow();
                    row["0"] = "error";
                    row["1"] = "error";
                    dtError.Rows.Add(row);
                    return dtError;
                }
            }            
        }
        /// <summary>
        /// Ejecuta un stord en la BD
        /// </summary>
        /// <param name="Stord">Nombre del Stord a ejecutar</param>
        /// <param name="TablaSube">objeto DataTable que envía al stord</param>
        /// <param name="NombreTabla">nombre del parametro de tabla</param>
        /// <param name="Parametros">Parametros que requiere el stord</param>
        /// <returns>Objeto DatatTable con la coleccion de datos</returns>
        public bool Ejecuta(string Stord, DataTable TablaSube, String NombreTabla, DataTable TablaSube1, String NombreTabla1, string[,] Parametros = null)
        {
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.Parameters.Add(new SqlParameter(NombreTabla1, TablaSube1));
                cmd.Parameters.Add(new SqlParameter(NombreTabla, TablaSube));
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    cmd.ExecuteNonQuery();
                    cmd.Connection.Close();
                    return true;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }

            }
        }

        /// <summary>
        /// Ejecuta un stord en la BD
        /// </summary>
        /// <param name="Stord">Nombre del Stord a ejecutar</param>
        /// <param name="TablaSube">objeto DataTable que envía al stord</param>
        /// <param name="NombreTabla">nombre del parametro de tabla</param>
        /// <param name="Parametros">Parametros que requiere el stord</param>
        /// <returns>Objeto DatatTable con la coleccion de datos</returns>
        public bool Ejecuta(string Stord, DataTable TablaSube, String NombreTabla, DataTable TablaSube1, String NombreTabla1, DataTable TablaSube2, String NombreTabla2, string[,] Parametros = null)
        {

            DataTable dt = new DataTable();
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.Parameters.Add(new SqlParameter(NombreTabla2, TablaSube2));
                cmd.Parameters.Add(new SqlParameter(NombreTabla1, TablaSube1));
                cmd.Parameters.Add(new SqlParameter(NombreTabla, TablaSube));
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    cmd.ExecuteNonQuery();
                    cmd.Connection.Close();
                    return true;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }
            }
        }

        public int Ejecuta(string Stord, DataTable TablaSube1, String NombreTabla1, DataTable TablaSube2, String NombreTabla2, DataTable TablaSube3, String NombreTabla3, DataTable TablaSube4, String NombreTabla4, DataTable TablaSube6, String NombreTabla6, DataTable TablaSube5, String NombreTabla5, string[,] Parametros = null)
        {

            DataTable dt = new DataTable();
            int lastRow = 0;
            using (SqlCommand cmd = new SqlCommand(Stord, Conexion()))
            {
                if (Parametros != null)
                    for (int i = Numeros.CERO; i < Parametros.Length / Numeros.DOS; i++)
                        cmd.Parameters.AddWithValue(Parametros[i, Numeros.CERO].ToString(), Parametros[i, Numeros.UNO].ToString());
                cmd.Parameters.Add(new SqlParameter(NombreTabla1, TablaSube1));
                cmd.Parameters.Add(new SqlParameter(NombreTabla2, TablaSube2));
                cmd.Parameters.Add(new SqlParameter(NombreTabla3, TablaSube3));
                cmd.Parameters.Add(new SqlParameter(NombreTabla4, TablaSube4));
                cmd.Parameters.Add(new SqlParameter(NombreTabla6, TablaSube6));
                cmd.Parameters.Add(new SqlParameter(NombreTabla5, TablaSube5));
                cmd.CommandType = CommandType.StoredProcedure;
                try
                {
                    cmd.CommandTimeout = 0;
                    cmd.Connection.Open();
                    lastRow = int.Parse(cmd.ExecuteScalar().ToString());
                    cmd.Connection.Close();
                    return lastRow;
                }
                catch (Exception e)
                {
                    cmd.Connection.Close();
                    throw new Exception(e.Message);
                }



            }
        }
    }
}
