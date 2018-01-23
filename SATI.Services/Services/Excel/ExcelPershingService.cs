using System;
using System.Collections.Generic;
using SATI.Common.Entities.Excel;
using System.Data;
using System.Data.OleDb;
using SATI.Services.Helper;
using SATI.Common.Common;
using SATI.Common.Entities;

namespace SATI.Services.Services.Excel
{
    public class ExcelPershingService
    {
        private PershingCommon _pershingCommon = new PershingCommon();
        private int TotalRegistrosExcel = 0;

        #region Rescatar los datos desde el excel
        public ActivityPershing ObtenerDatosExcel(string urlArchivo)
        {
            ActivityPershing activityResult = new ActivityPershing();
            activityResult.Error = new PershingError();
            List<ExcelPershing> listadoExcelPershing = new List<ExcelPershing>();
            try
            {
                listadoExcelPershing = ProcesarExcelPershing(urlArchivo);
                if (listadoExcelPershing.Count > 0)
                {
                    activityResult.FechaCarga = listadoExcelPershing[0].fecha_proceso;
                    activityResult.Instrumento = "Renta Variable";
                    activityResult.valorDolar = _pershingCommon.ObtenerValorDolarByFecha(activityResult.FechaCarga);
                    activityResult.NumeroFolio = listadoExcelPershing[0].folio;
                    activityResult.TotalRegistros = listadoExcelPershing.Count;
                    activityResult.TotalComCLP = "0";
                    activityResult.TotalComUS = 0;
                    activityResult.RegistroAnulados = int.Parse(_pershingCommon.ObtenerCantidadPershingByFecha(activityResult.FechaCarga));
                    activityResult.Duplicado = _pershingCommon.ExisteFechaProcesoPershing(activityResult.FechaCarga);
                    activityResult.listadoExcelPershing = listadoExcelPershing;
                }
                else
                {
                    activityResult.Error.Mensaje = $"Se produjo un error al intentar procesar el archivo seleccionado.";
                    activityResult.Error.descripcion = "el archivo seleccionado no contiene datos para procesar.";
                }
            }
            catch (Exception e)
            {
                activityResult.Error.Mensaje = $"Se produjo un error al intentar procesar el archivo seleccionado.";
                activityResult.Error.descripcion = e.Message;
            }

            return activityResult;
        }
        #endregion
        #region Procesar Excel (Guardar Datos Excel en Base de Datos)
        private List<ExcelPershing> ProcesarExcelPershing(string urlArchivo)
        {
            string ConecctionExcel = ExcelHelper.StringConnectionExcel().Replace("FILE_EXCEL", urlArchivo);
            DataTable dtResultado = new DataTable();
            try
            {
                OleDbConnection Connection = new OleDbConnection(ConecctionExcel);
                Connection.Open();
                DataTable dt = Connection.GetSchema("Tables");
                OleDbDataAdapter dtAdapter = new OleDbDataAdapter($"Select * From [{dt.Rows[0]["TABLE_NAME"].ToString()}{ExcelHelper.RangeColumsExcel()}]", ConecctionExcel);
                dtAdapter.Fill(dtResultado);
                Connection.Dispose();
                TotalRegistrosExcel = dtResultado.Rows.Count;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                throw new ArgumentException("El archivo que intenta subir no es un archivo excel de activity pershing.");
            }

            return ObtenerListadoPershingFromExcel(dtResultado);
        }
        private List<ExcelPershing> ObtenerListadoPershingFromExcel(DataTable dtExcel)
        {
            List<ExcelPershing> ListadoPershing = new List<ExcelPershing>();
            dtExcel = ValidacionesExcelPershing(dtExcel);
            string NuevoFolio = DateTime.Now.ToString().Replace("-", "").Replace("/", "").Replace(":", "").Replace(" ", "");
            try
            {
                foreach (DataRow row in dtExcel.Rows)
                {
                    if (row["ACCOUNT NUMBER"].ToString().IndexOf("NN9") == 0)
                    {
                        if (!_pershingCommon.VerificarCuentasTanner(row["ACCOUNT NUMBER"].ToString()))
                        {

                            if (_pershingCommon.ExisteClientePershing(row["ACCOUNT NUMBER"].ToString()))
                            {
                                double precio = (double.Parse(row["PRICE"].ToString()) < 0) ? double.Parse(row["PRICE"].ToString()) * -1 : double.Parse(row["PRICE"].ToString());
                                double cantidad = (double.Parse(row["QUANTITY"].ToString()) < 0) ? double.Parse(row["QUANTITY"].ToString()) * -1 : double.Parse(row["QUANTITY"].ToString());

                                if (cantidad > 0 && precio > 0)
                                {
                                    ExcelPershing excelPershing = new ExcelPershing();
                                    excelPershing.fecha_proceso = row["PROCESS DATE"].ToString();
                                    excelPershing.num_cuenta = row["ACCOUNT NUMBER"].ToString();
                                    excelPershing.nombre_corto = row["SHORT NAME"].ToString();
                                    excelPershing.fecha_trans = row["TRADE DATE"].ToString() == "" ? new DateTime(1990, 1, 1).ToString() : row["TRADE DATE"].ToString();
                                    excelPershing.fecha_strlmnt = row["STLMNT DATE"].ToString() == "" ? new DateTime(1990, 1, 1).ToString() : row["STLMNT DATE"].ToString();
                                    excelPershing.tipo_ope = row["BUY/SELL"].ToString();
                                    excelPershing.cantidad = row["QUANTITY"].ToString();
                                    excelPershing.simbolo = row["SYMBOL"].ToString();
                                    excelPershing.cusip = row["CUSIP"].ToString();
                                    excelPershing.desc_sec = row["SEC# DESCRIPTION"].ToString();
                                    excelPershing.precio = row["PRICE"].ToString();
                                    excelPershing.cod_trans = row["TRANS# CODE"].ToString();
                                    excelPershing.interes = row["INTEREST"].ToString();
                                    excelPershing.monto_neto_base = row["NET AMT#_(BASE CCY)"].ToString();
                                    excelPershing.monto_neto_trans = row["TRANS# NET AMT#_(TRANS# CCY)"].ToString();
                                    excelPershing.categoria = row["CATEG#"].ToString();
                                    excelPershing.comision = row["COMM#"].ToString();
                                    excelPershing.impuestos = row["FEES"].ToString();
                                    excelPershing.currency = "USD";
                                    excelPershing.bruto = (double.Parse(row["PRICE"].ToString()) * double.Parse(row["QUANTITY"].ToString())).ToString();
                                    excelPershing.monto_principal = row["PRINCIPAL"].ToString();
                                    excelPershing.isin = row["ISIN"].ToString();
                                    excelPershing.ip_number = row["IP NUMBER"].ToString();
                                    excelPershing.ip_name = row["IP NAME"].ToString();
                                    excelPershing.acct_mnemo = row["ACCT# MNEMONIC"].ToString();
                                    excelPershing.mstr_mnemo = row["MSTR# MNEMONIC"].ToString();
                                    excelPershing.corr_ref = row["CORR REF NUMBER"].ToString();
                                    excelPershing.solicited = row["SOLICITED"].ToString();
                                    excelPershing.org_input_src = row["ORG# INPUT SRC#"].ToString();
                                    excelPershing.clase_activo = row["ASSET CLASS"].ToString();
                                    excelPershing.sec_payment = row["SEC# PAYMENT"].ToString();
                                    excelPershing.lot_selection = row["LOT SELECTION"].ToString();
                                    excelPershing.tipo_cuenta = row["ACCT# TYPE"].ToString();
                                    excelPershing.exec_broker = "Pershing";
                                    excelPershing.nombre_completo = "EE.UU.";
                                    excelPershing.monto_neto = row["NET AMT#_(BASE CCY)"].ToString();
                                    excelPershing.act_descrip = row["ACTIVITY DESCRIPTION"].ToString();
                                    excelPershing.folio = NuevoFolio;


                                    ListadoPershing.Add(excelPershing);
                                }
                            } else
                            {
                                throw new ArgumentException($"La Cuenta : '{row["ACCOUNT NUMBER"].ToString()}' no existe en los registros de Pershing");
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw new ArgumentException(e.Message);
            }

            return ListadoPershing;
        }
        private DataTable ValidacionesExcelPershing(DataTable dtPershing)
        {
            DataTable dtResult = new DataTable();
            int contador = 0;

            //PROCESO DE MODIFICACIONES VENTAS(SX)/COMPRAS(BX)
            //VENTAS Y COMPRAS
            foreach (DataRow row in dtPershing.Rows)
            {
                contador += 1;

                if (row["PRICE"].ToString().Equals("") && row["PRICE"].ToString().Equals("COMM#"))
                {
                    continue;
                }
                else
                {
                    if (contador.Equals(dtPershing.Rows.Count))
                    {
                        break;
                    }
                }

                if ((row["PROCESS DATE"].ToString().Equals(dtPershing.Rows[contador]["PROCESS DATE"].ToString())) &&
                   (row["ACCOUNT NUMBER"].ToString().Equals(dtPershing.Rows[contador]["ACCOUNT NUMBER"].ToString())) &&
                   (row["CUSIP"].ToString().Equals(dtPershing.Rows[contador]["CUSIP"].ToString())) &&
                   (Math.Abs(double.Parse(row["PRICE"].ToString()))) == (Math.Abs(double.Parse(dtPershing.Rows[contador]["PRICE"].ToString()))) &&
                   (Math.Abs(double.Parse(row["COMM#"].ToString()))) == (Math.Abs(double.Parse(dtPershing.Rows[contador]["COMM#"].ToString())))
                   )
                {
                    if (row["BUY/SELL"].ToString().Equals("SC"))
                    {
                        row["BUY/SELL"] = "S";
                        row.AcceptChanges();
                    }
                    else
                    {
                        if (dtPershing.Rows[contador]["BUY/SELL"].Equals("SC"))
                        {
                            dtPershing.Rows[contador]["BUY/SELL"] = "S";
                            dtPershing.AcceptChanges();
                        }
                        else
                        {
                            if (row["BUY/SELL"].ToString().Equals("BC"))
                            {
                                row["BUY/SELL"] = "B";
                                row.AcceptChanges();
                            }
                            else
                            {
                                if (dtPershing.Rows[contador]["BUY/SELL"].Equals("BC"))
                                {
                                    dtPershing.Rows[contador]["BUY/SELL"] = "B";
                                    dtPershing.Rows[contador].AcceptChanges();
                                }
                                else
                                {
                                    if (row["BUY/SELL"].ToString().Equals("BX"))
                                    {
                                        row.Delete();
                                        row.AcceptChanges();
                                    }
                                    else
                                    {
                                        if (dtPershing.Rows[contador]["BUY/SELL"].Equals("BX"))
                                        {
                                            dtPershing.Rows[contador].Delete();
                                            dtPershing.Rows[contador].AcceptChanges();
                                        }
                                        else
                                        {
                                            if (row["BUY/SELL"].ToString().Equals("SX"))
                                            {
                                                row.Delete();
                                                row.AcceptChanges();
                                            }
                                            else
                                            {
                                                if (dtPershing.Rows[contador]["BUY/SELL"].Equals("SX"))
                                                {
                                                    dtPershing.Rows[contador].Delete();
                                                    dtPershing.Rows[contador].AcceptChanges();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return dtPershing;
        }
        #endregion

    }
}
