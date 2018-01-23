using System;
using System.Collections.Generic;
using SATI.Common.Entities.Excel;
using System.Data.SqlClient;
using SATI.Common.Entities.Pershing;
using SATI.Common.Entities;
using System.Data;


namespace SATI.Common.Common
{
    public class PershingCommon
    {
        public bool ExisteClientePershing(string AcountNumber)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            bool resultado = false;

                string Query = "";
                Query = "SELECT * FROM PUBLICA.DBO.pershing_cli WHERE REPLACE(cta_pershing,'-', '') = @cta_pershing";

                try
                {
                    SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                    command.Parameters.AddWithValue("@cta_pershing", AcountNumber);

                    dtReader = command.ExecuteReader();
                    while (dtReader.Read())
                    {
                        resultado = true;
                    }
                }
                catch (SqlException ex)
                {
                    throw new ArgumentException($"Error al Consultar Modulos Aplicacion Detalle : {ex.Message}");
                }catch(Exception e)
                {
                Console.WriteLine(e.Message);
                }               
                finally
                {
                    dtReader.Close();
                    conectar.Desconectar();
                }

                return resultado;
            }
            
        

        public void GuardarExcelPershingBD(List<ExcelPershing> listadoPershing)
        {
            Connection conectar = new Connection();

            string Query = "";
            Query = "INSERT INTO PUBLICA.dbo.crg_pershing_excel ";
            Query += "(psh_carga, psh_fecha_proceso, psh_num_cuenta, psh_nombre_corto, ";
            Query += "psh_fecha_trans,psh_fecha_strlmnt, psh_tipo_ope, psh_cantidad, psh_simbolo, ";
            Query += "psh_cusip, psh_desc_sec, psh_precio, psh_cod_trans, psh_interes, ";
            Query += "psh_monto_neto_base, psh_monto_neto_trans, psh_categoria, psh_comision, ";
            Query += "psh_impuestos, psh_currency, psh_monto_principal, psh_isin, psh_ip_number,";
            Query += "psh_ip_name, psh_acct_mnemo, psh_mstr_mnemo, psh_corr_ref, psh_solicited,";
            Query += "psh_org_input_src, psh_clase_activo, psh_sec_payment, psh_lot_selection,";
            Query += "psh_tipo_cuenta, psh_exec_broker, psh_nombre_completo, psh_monto_neto,";
            Query += "psh_act_descrip, psh_stt, psh_comm_tanner, psh_gastos_tanner, psh_comments, psh_bruto) VALUES ";

            Query += "(@psh_carga, @psh_fecha_proceso, @psh_num_cuenta, @psh_nombre_corto, ";
            Query += "@psh_fecha_trans, @psh_fecha_strlmnt, @psh_tipo_ope, @psh_cantidad, @psh_simbolo, ";
            Query += "@psh_cusip, @psh_desc_sec, @psh_precio, @psh_cod_trans, @psh_interes, ";
            Query += "@psh_monto_neto_base, @psh_monto_neto_trans, @psh_categoria, @psh_comision, ";
            Query += "@psh_impuestos, @psh_currency, @psh_monto_principal, @psh_isin, @psh_ip_number,";
            Query += "@psh_ip_name, @psh_acct_mnemo, @psh_mstr_mnemo, @psh_corr_ref, @psh_solicited,";
            Query += "@psh_org_input_src, @psh_clase_activo, @psh_sec_payment, @psh_lot_selection,";
            Query += "@psh_tipo_cuenta, @psh_exec_broker, @psh_nombre_completo, @psh_monto_neto,";
            Query += "@psh_act_descrip, @psh_stt, @psh_comm_tanner, @psh_gastos_tanner, @psh_comments, @psh_bruto)";
            try
            {
                foreach (ExcelPershing pershing in listadoPershing)
                {
                    SqlCommand command = new SqlCommand(Query, conectar.Conectar());

                    command.Parameters.AddWithValue("@psh_carga", pershing.folio); /*N*/
                    command.Parameters.AddWithValue("@psh_fecha_proceso", DateTime.Parse(pershing.fecha_proceso)); /*0*/
                    command.Parameters.AddWithValue("@psh_num_cuenta", pershing.num_cuenta.Trim()); /*1*/
                    command.Parameters.AddWithValue("@psh_nombre_corto", pershing.nombre_corto.Trim()); /*2*/
                    command.Parameters.AddWithValue("@psh_fecha_trans", DateTime.Parse(pershing.fecha_trans)); /*3*/
                    command.Parameters.AddWithValue("@psh_fecha_strlmnt", DateTime.Parse(pershing.fecha_strlmnt)); /*4*/
                    command.Parameters.AddWithValue("@psh_tipo_ope", pershing.tipo_ope.Trim()); /*5*/
                    command.Parameters.AddWithValue("@psh_cantidad", Decimal.Parse(pershing.cantidad)); /*6*/                   
                    command.Parameters.AddWithValue("@psh_simbolo", pershing.simbolo.Trim()); /*7*/
                    command.Parameters.AddWithValue("@psh_cusip", pershing.cusip.Trim()); /*8*/
                    command.Parameters.AddWithValue("@psh_desc_sec", pershing.desc_sec.Trim()); /*9*/
                    command.Parameters.AddWithValue("@psh_precio", Decimal.Parse(pershing.precio)); /*10*/
                    command.Parameters.AddWithValue("@psh_cod_trans", (double.Parse(pershing.cantidad)/ double.Parse(pershing.monto_principal)).ToString()); /*11*/
                    command.Parameters.AddWithValue("@psh_interes", Decimal.Parse(pershing.interes)); /*12*/
                    command.Parameters.AddWithValue("@psh_monto_neto_base", Decimal.Parse(pershing.monto_neto_base)); /*13*/
                    command.Parameters.AddWithValue("@psh_monto_neto_trans", Decimal.Parse(pershing.monto_neto_trans)); /*14*/
                    command.Parameters.AddWithValue("@psh_categoria", pershing.categoria.Trim()); /*15*/
                    command.Parameters.AddWithValue("@psh_comision", Decimal.Parse(pershing.comision)); /*16*/
                    command.Parameters.AddWithValue("@psh_impuestos", Decimal.Parse(pershing.impuestos)); /*17*/
                    command.Parameters.AddWithValue("@psh_currency", pershing.currency.Trim()); /*18*/
                    command.Parameters.AddWithValue("@psh_monto_principal", Decimal.Parse(pershing.monto_principal)); /*19*/
                    command.Parameters.AddWithValue("@psh_isin", pershing.isin.Trim()); /*20*/
                    command.Parameters.AddWithValue("@psh_ip_number", pershing.ip_number.Trim()); /*21*/
                    command.Parameters.AddWithValue("@psh_ip_name", pershing.ip_name.Trim()); /*22*/
                    command.Parameters.AddWithValue("@psh_acct_mnemo", pershing.acct_mnemo.Trim());/*23*/
                    command.Parameters.AddWithValue("@psh_mstr_mnemo", pershing.mstr_mnemo.Trim());/*24*/
                    command.Parameters.AddWithValue("@psh_corr_ref", pershing.corr_ref.Trim());/*25*/
                    command.Parameters.AddWithValue("@psh_solicited", pershing.solicited.Trim());/*26*/
                    command.Parameters.AddWithValue("@psh_org_input_src", pershing.org_input_src.Trim());/*27*/
                    command.Parameters.AddWithValue("@psh_clase_activo", pershing.clase_activo.Trim());/*28*/
                    command.Parameters.AddWithValue("@psh_sec_payment", pershing.sec_payment.Trim());/*29*/
                    command.Parameters.AddWithValue("@psh_lot_selection", pershing.lot_selection.Trim());/*30*/
                    command.Parameters.AddWithValue("@psh_tipo_cuenta", pershing.tipo_cuenta.Trim());/*31*/
                    command.Parameters.AddWithValue("@psh_exec_broker", pershing.exec_broker.Trim());/*32*/
                    command.Parameters.AddWithValue("@psh_nombre_completo", pershing.nombre_completo);/*33*/
                    command.Parameters.AddWithValue("@psh_monto_neto", Decimal.Parse(pershing.monto_neto));/*34*/
                    command.Parameters.AddWithValue("@psh_act_descrip", pershing.act_descrip.Trim());/*35*/
                    command.Parameters.AddWithValue("@psh_stt",int.Parse("1"));
                    command.Parameters.AddWithValue("@psh_comm_tanner",Decimal.Parse("0.00"));
                    command.Parameters.AddWithValue("@psh_gastos_tanner", Decimal.Parse("0.00"));
                    command.Parameters.AddWithValue("@psh_comments","");
                    command.Parameters.AddWithValue("@psh_bruto", Decimal.Parse(pershing.bruto));

                    int FilasAfectadas = command.ExecuteNonQuery();
                }         
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Insertar registros Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {           
                conectar.Desconectar();
            }

        }

        public bool ExisteFechaProcesoPershing(string fecha)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
          
            string Query = "";
            Query = "SELECT COUNT(*) FROM PUBLICA.dbo.crg_pershing_d WHERE psh_fecha_proceso = @FECHA";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@FECHA", DateTime.Parse(fecha));

                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    return true;
                }
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return false;
        }

        public bool EliminaFechasRepetidas(string fecha)
        {
            Connection conectar = new Connection();

            string Query = "";
            Query = " UPDATE PUBLICA.dbo.crg_pershing_d SET psh_estado_item = 0 WHERE psh_fecha_proceso = @FECHA";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@FECHA", DateTime.Parse(fecha));

               command.ExecuteNonQuery();
               return true;
            }
            catch (SqlException ex)
            {
                conectar.Conectar().BeginTransaction().Rollback();
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool LimpiaTablasExcel(string fecha)
        {
            Connection conectar = new Connection();

            string Query = "";
            Query = "DELETE FROM PUBLICA.dbo.crg_pershing_excel WHERE psh_fecha_proceso = @FECHA; ";
            Query += "DELETE FROM PUBLICA.dbo.crg_pershing_origin WHERE psh_fecha_proceso = @FECHA; ";
            Query += "DELETE FROM GESTION_AGENTES.dbo.PERSHING_RF_GES WHERE FECHA = @FECHA; ";
            Query += "DELETE FROM GESTION_AGENTES.dbo.PERSHING_RV_GES WHERE FECHA = @FECHA; ";
            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@FECHA", DateTime.Parse(fecha));

                command.ExecuteNonQuery();
                return true;
            }
            catch (SqlException ex)
            {
                conectar.Conectar().BeginTransaction().Rollback();
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public string ObtenerCantidadPershingByFecha(string fecha)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            string contador = string.Empty;

            string Query = "";
            Query = "SELECT CONVERT(VARCHAR(5),COUNT(*)) AS CONTADOR FROM PUBLICA.dbo.crg_pershing_d WHERE psh_fecha_proceso = @FECHA";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@FECHA", DateTime.Parse(fecha));

                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    contador = dtReader.GetString(0);
                }
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return contador;
        }

        #region Cargadores segun Productos 
        public bool CargaRentaVariablePershing(string NumeroFolio)
        {
            Connection conectar = new Connection();

            string Query = " ";
            Query = "INSERT INTO PUBLICA.dbo.crg_pershing_origin (psh_num_cuenta, psh_fecha_proceso, ";
            Query += "psh_tipo_ope, psh_cantidad, psh_simbolo, psh_cusip, psh_precio, psh_interes, psh_monto_neto_base, ";
            Query += "psh_comision, psh_impuestos, psh_currency, psh_exec_broker, psh_monto_neto, ";
            Query += "psh_nombre_corto, psh_fecha_strlmnt, psh_monto_neto_trans, psh_tipo_instrumento, psh_stt, ";
            Query += "psh_carga, psh_isin, psh_comm_tanner, psh_gastos_tanner, psh_comments, psh_bruto,psh_desc_sec,psh_fecha_trans) ";

            Query += "SELECT a.psh_num_cuenta, MAX(a.psh_fecha_proceso) AS psh_fecha_proceso, a.psh_tipo_ope,";
            Query += "SUM(a.psh_cantidad) AS psh_cantidad, a.psh_simbolo, a.psh_cusip, CAST(SUM(CASE WHEN a.psh_monto_principal <=0 THEN a.psh_monto_principal * -1 ELSE a.psh_monto_principal END)/SUM(CASE WHEN a.psh_cantidad <= 0 THEN a.psh_cantidad * -1 ELSE a.psh_cantidad END) AS DECIMAL(18,9)) AS psh_precio,";
            Query += "SUM(a.psh_interes) AS psh_interes, SUM(a.psh_monto_neto_base) AS psh_monto_neto_base,";
            Query += "SUM(a.psh_comision) AS psh_comision, SUM(a.psh_impuestos) AS psh_impuestos, a.psh_currency,";
            Query += "a.psh_exec_broker, SUM(a.psh_monto_neto) AS psh_monto_neto, a.psh_nombre_corto, ";
            Query += "MAX(a.psh_fecha_strlmnt) AS psh_fecha_strlmnt, SUM(a.psh_monto_neto_trans) AS psh_monto_neto_trans, ";
            Query += "'Renta Variable', 0, a.psh_carga, a.psh_isin, '0.00', '0.00', '', SUM(a.psh_bruto) AS psh_bruto,a.psh_act_descrip,a.psh_fecha_trans ";
            Query += "FROM PUBLICA.dbo.crg_pershing_excel a ";
            Query += "WHERE a.psh_carga = @NUMEROFOLIO ";
            Query += "AND LEN(a.psh_simbolo) < 5 ";
            Query += "AND a.psh_stt = 1 ";
            Query += "AND UPPER(a.psh_clase_activo) not like '%LOAD FUND%' ";
            Query += "GROUP BY a.psh_num_cuenta, a.psh_fecha_proceso, a.psh_tipo_ope, a.psh_simbolo, a.psh_cusip, ";
            Query += "a.psh_currency, a.psh_exec_broker, a.psh_nombre_corto, a.psh_carga, a.psh_isin,a.psh_act_descrip,a.psh_desc_sec,a.psh_fecha_trans  ";
                  
            try
            {
              
                    SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                    command.Parameters.AddWithValue("@NUMEROFOLIO", NumeroFolio);          
                    command.ExecuteNonQuery();
                    return true;
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Insertar registros Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool CargaFondosMutuosPershing(string NumeroFolio)
        {
            Connection conectar = new Connection();

            string Query = " ";
            Query = "INSERT INTO PUBLICA.dbo.crg_pershing_origin (psh_num_cuenta, psh_fecha_proceso, ";
            Query += "psh_tipo_ope, psh_cantidad, psh_simbolo, psh_cusip, psh_precio, psh_interes, psh_monto_neto_base, ";
            Query += "psh_comision, psh_impuestos, psh_currency, psh_exec_broker, psh_monto_neto, ";
            Query += "psh_nombre_corto, psh_fecha_strlmnt, psh_monto_neto_trans, psh_tipo_instrumento, psh_stt, ";
            Query += "psh_carga, psh_isin, psh_comm_tanner, psh_gastos_tanner, psh_comments, psh_bruto,psh_desc_sec,psh_fecha_trans) ";

            Query += "SELECT a.psh_num_cuenta, MAX(a.psh_fecha_proceso) AS psh_fecha_proceso, a.psh_tipo_ope,";
            Query += "SUM(a.psh_cantidad) AS psh_cantidad,'FM-'+SUBSTRING(a.psh_desc_sec,0,5) as psh_simbolo, a.psh_cusip, CAST(SUM(CASE WHEN a.psh_monto_principal <=0 THEN a.psh_monto_principal * -1 ELSE a.psh_monto_principal END)/SUM(CASE WHEN a.psh_cantidad <= 0 THEN a.psh_cantidad * -1 ELSE a.psh_cantidad END) AS DECIMAL(18,9)) AS psh_precio,";
            Query += "SUM(a.psh_interes) AS psh_interes, SUM(a.psh_monto_neto_base) AS psh_monto_neto_base,";
            Query += "SUM(a.psh_comision) AS psh_comision, SUM(a.psh_impuestos) AS psh_impuestos, a.psh_currency,";
            Query += "a.psh_exec_broker, SUM(a.psh_monto_neto) AS psh_monto_neto, a.psh_nombre_corto, ";
            Query += "MAX(a.psh_fecha_strlmnt) AS psh_fecha_strlmnt, SUM(a.psh_monto_neto_trans) AS psh_monto_neto_trans, ";
            Query += "'Fondos Mutuos', 0, a.psh_carga, a.psh_isin, '0.00', '0.00', '', SUM(a.psh_bruto) AS psh_bruto,a.psh_act_descrip,a.psh_fecha_trans ";
            Query += "FROM PUBLICA.dbo.crg_pershing_excel a ";
            Query += "WHERE a.psh_carga = @NUMEROFOLIO ";
            Query += "AND LEN(a.psh_simbolo) <= 5 ";
            Query += "AND UPPER(a.psh_clase_activo) like '%LOAD FUND%' ";
            Query += "AND a.psh_stt = 1 ";
            Query += "GROUP BY a.psh_num_cuenta, a.psh_fecha_proceso, a.psh_tipo_ope, a.psh_simbolo, a.psh_cusip, ";
            Query += "a.psh_currency, a.psh_exec_broker, a.psh_nombre_corto, a.psh_carga, a.psh_isin,a.psh_act_descrip,a.psh_fecha_trans,a.psh_desc_sec ";

            try
            {

                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@NUMEROFOLIO", NumeroFolio);
                command.ExecuteNonQuery();
                return true;
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Insertar registros Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool CargaRentaFijaPershing(string NumeroFolio)
        {
            Connection conectar = new Connection();

            string Query = " ";
            Query = "INSERT INTO PUBLICA.dbo.crg_pershing_origin (psh_num_cuenta, psh_fecha_proceso, ";
            Query += "psh_tipo_ope, psh_cantidad, psh_simbolo, psh_cusip, psh_precio, psh_interes, psh_monto_neto_base, ";
            Query += "psh_comision, psh_impuestos, psh_currency, psh_exec_broker, psh_monto_neto, ";
            Query += "psh_nombre_corto, psh_fecha_strlmnt, psh_monto_neto_trans, psh_tipo_instrumento, psh_stt, psh_carga, psh_isin, psh_comm_tanner, psh_gastos_tanner, psh_comments, psh_bruto,psh_desc_sec,psh_fecha_trans) ";

            Query += "SELECT a.psh_num_cuenta, a.psh_fecha_proceso, a.psh_tipo_ope, a.psh_cantidad,";
            Query += "a.psh_simbolo, a.psh_cusip, a.psh_precio, a.psh_interes, a.psh_monto_neto_base,";
            Query += "a.psh_comision, a.psh_impuestos, a.psh_currency, a.psh_exec_broker, ";
            Query += "a.psh_monto_neto, a.psh_nombre_corto, a.psh_fecha_strlmnt, ";
            Query += "a.psh_monto_neto_trans, 'Renta Fija', 0, a.psh_carga, a.psh_isin, '0.00', '0.00', '', a.psh_bruto,a.psh_act_descrip,a.psh_fecha_trans ";
            Query += "FROM PUBLICA.dbo.crg_pershing_excel a ";
            Query += "WHERE a.psh_carga = @NUMEROFOLIO ";
            Query += "AND LEN(a.psh_simbolo) >= 5 ";
            Query += "AND a.psh_stt = 1 ";
        
            try
            {

                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@NUMEROFOLIO", NumeroFolio);
                command.ExecuteNonQuery();
                return true;
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Insertar registros Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                conectar.Desconectar();
            }
        }
        #endregion
        public List<ExcelPershing> Carga_Origen_excel(string numeroFolio,string instrumento)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            List<ExcelPershing> ListadoOrigenResult = new List<ExcelPershing>();

            string Query = "";
            Query = "SELECT psh_num_cuenta, psh_fecha_proceso, psh_tipo_ope, ";
            Query += "psh_cantidad, psh_simbolo, psh_cusip, psh_precio, ";
            Query += "psh_interes, psh_monto_neto_base, psh_comision, psh_impuestos, ";
            Query += "psh_currency, psh_exec_broker, psh_monto_neto, psh_nombre_corto, ";
            Query += "psh_fecha_strlmnt, psh_monto_neto_trans, psh_tipo_instrumento, psh_id, psh_isin, psh_bruto,psh_desc_sec,psh_fecha_trans ";
            Query += "FROM PUBLICA.dbo.crg_pershing_origin ";
            Query += "WHERE psh_carga = @NUMEROFOLIO AND psh_tipo_instrumento = @INSTRUMENTO AND psh_stt = 0 ";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@NUMEROFOLIO", numeroFolio);
                command.Parameters.AddWithValue("@INSTRUMENTO", (instrumento == "" ? "Renta Variable" : instrumento));

                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    ExcelPershing excelPershing = new ExcelPershing();
                    excelPershing.num_cuenta = dtReader.GetString(0);
                    excelPershing.fecha_proceso = dtReader.GetDateTime(1).ToString();
                    excelPershing.tipo_ope = dtReader.GetString(2) == "B" ? "COMPRA":"VENTA";
                    excelPershing.cantidad = dtReader.GetDecimal(3) > 0 ? dtReader.GetDecimal(3).ToString() : (dtReader.GetDecimal(3) * -1).ToString();
                    excelPershing.simbolo = dtReader.GetString(4);
                    excelPershing.cusip = dtReader.GetString(5);
                    excelPershing.precio = dtReader.GetDecimal(6).ToString();
                    excelPershing.interes = dtReader.GetDecimal(7).ToString();
                    excelPershing.monto_neto_base = (dtReader.GetDecimal(6) * dtReader.GetDecimal(3)).ToString();
                    excelPershing.comision = dtReader.GetDecimal(9).ToString();
                    excelPershing.impuestos = dtReader.GetDecimal(10).ToString();
                    excelPershing.currency = dtReader.GetString(11);
                    excelPershing.exec_broker = dtReader.GetString(12).ToString();
                    excelPershing.monto_neto = dtReader.GetDecimal(13) > 0 ? dtReader.GetDecimal(13).ToString() : (dtReader.GetDecimal(13) * -1).ToString();
                    excelPershing.nombre_corto = dtReader.GetString(14);
                    excelPershing.fecha_strlmnt = dtReader.GetDateTime(15).ToString();
                    excelPershing.monto_neto_trans = dtReader.GetDecimal(16).ToString();
                    excelPershing.tipo_cuenta = dtReader.GetString(17);
                    excelPershing.Id = dtReader.GetInt32(18).ToString();
                    excelPershing.isin = dtReader.GetString(19);
                    excelPershing.bruto = dtReader.GetDecimal(20).ToString();
                    excelPershing.desc_sec = dtReader.GetString(21);
                    excelPershing.fecha_trans = dtReader.GetDateTime(22).Date.ToString();

                    ListadoOrigenResult.Add(excelPershing);
                }
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            catch(Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return ListadoOrigenResult;
        }

        public double ObtenerValorDolarByFecha(string fecha)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            double valor = 0;

            string Query = "SELECT VALOR AS VALOR FROM GESTION_AGENTES.dbo.valor_dolar WHERE FECHA = @FECHA ";
        

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@FECHA", DateTime.Parse(fecha));
              

                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    valor = dtReader.GetDouble(0);
                }
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return valor;
        }

        public long ObtenerUltimoFolioAsignado()
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            long folio = 0;

            string Query = "SELECT ISNULL(MAX(fol_folio), 0) FROM PUBLICA.dbo.crg_pershing_folio";


            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());


                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    folio = dtReader.GetInt64(0);
                }
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            catch(Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return folio;
        }

        public bool ValidarRegistroFondosMutuos(string nombreCorto,string cusip,string numeroCuenta,string tradeDate)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            bool Validacion = false;

            string Query = "SELECT * FROM PUBLICA.DBO.crg_pershing_historico_ValFM ";
                   Query += "WHERE ACCOUNT = @CUENTA ";
                   Query += "AND SHORT_NAME = @SHORTNAME ";
                   Query += "AND PROCESS_DATE = @TRADEDATE ";
                   Query += "AND CUSIP = @CUSIP ";
                   Query += "AND SUBSTRING(ACCOUNT,0,6) = 'NN900' ";
                   Query += "AND LEN(SYMBOL) = 0 ";
                   Query += "AND FLAG = 0";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@SHORTNAME", nombreCorto);
                command.Parameters.AddWithValue("@CUSIP",cusip);
                command.Parameters.AddWithValue("@CUENTA",numeroCuenta);
                command.Parameters.AddWithValue("@TRADEDATE", tradeDate);

                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    Validacion = true;
                }
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return Validacion;
        }

        public bool ActualizaLogValidaFMExcel(string simbol,string cusip,string numeroCuenta)
        {
            Connection conectar = new Connection();

            string Query = "UPDATE PUBLICA.dbo.crg_pershing_excel Set psh_simbolo= @SIMBOL WHERE PSH_CUSIP = @CUSIP And PSH_NUM_CUENTA = @CUENTA";


            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@CUSIP", cusip);
                command.Parameters.AddWithValue("@CUENTA", numeroCuenta);
                command.Parameters.AddWithValue("@SIMBOL", simbol);
                command.ExecuteNonQuery();
                return true;
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool ActualizaLogValidaFMExcelHistorico(string simbol, string cusip, string numeroCuenta)
        {
            Connection conectar = new Connection();

            string Query = "UPDATE PUBLICA.DBO.crg_pershing_historico_ValFM SET FLAG=1, SYMBOL =@SIMBOL WHERE CUSIP = @CUSIP And ACCOUNT = @CUENTA";


            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@CUSIP", cusip);
                command.Parameters.AddWithValue("@CUENTA", numeroCuenta);
                command.Parameters.AddWithValue("@SIMBOL", simbol);
                command.ExecuteNonQuery();
                return true;
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool InsertaLogValidaFMExcelHistorico(string processDate,string acount,string shortname,string tradeDate,string stlmntDate,string simbol,string cusip,string fecha,string user)
        {
            Connection conectar = new Connection();

            string Query = "INSERT PUBLICA.DBO.crg_pershing_historicoLog_ValFM ([PROCESS_DATE] ,[ACCOUNT] ,[SHORT_NAME]  ,[TRADE_DATE],[STLMNT_DATE] , [SYMBOL] ,[CUSIP]  ,[DATE] ,[USER])";
                   Query = "VALUES(@PROCESSDATE,@ACOUNT,@SHORTNAME,@TRADEDATE,@STLMNTDATE,@SYMBOL,@CUSIP,@FECHA,@USER)";


            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@PROCESSDATE", processDate);
                command.Parameters.AddWithValue("@ACOUNT", acount);
                command.Parameters.AddWithValue("@SHORTNAME", shortname);
                command.Parameters.AddWithValue("@TRADEDATE", tradeDate);
                command.Parameters.AddWithValue("@STLMNTDATE", stlmntDate);
                command.Parameters.AddWithValue("@SYMBOL", simbol);
                command.Parameters.AddWithValue("@CUSIP", cusip);
                command.Parameters.AddWithValue("@FECHA", fecha);
                command.Parameters.AddWithValue("@USER", user);
                command.ExecuteNonQuery();
                return true;
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Fecha Proceso Pershing Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public ClientePershing ObtenerClientePershingByCta(string ctaPershing)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            ClientePershing cliente = new ClientePershing();

            //string Query = "SELECT [nombre cliente],Rut_cliente,[Cod#],REPLACE(REPLACE(Agente,[Cod#],''),'-','') FROM publica.dbo.pershing_cli WHERE REPLACE(cta_pershing,'-', '') = @CTAPERSHING  ";
            string Query = "SELECT COALESCE(rtrim(ltrim(b.nombre_cli)), rtrim(ltrim(a.[nombre cliente]))) collate latin1_general_cs_ai AS CLIENTE, ";
            Query += "COALESCE(rtrim(ltrim(b.RUT_CLI)), rtrim(ltrim(a.Rut_cliente))) collate latin1_general_cs_ai  AS RUT_CLI, ";
            Query += "COALESCE(rtrim(ltrim(c.COD_AGENTE)), rtrim(ltrim(a.[Cod#]))) AS COD_EJECUTIVO,  ";
            Query += "COALESCE(rtrim(ltrim(c.NOMBRE_AGENTE)), rtrim(ltrim(a.Agente))) collate latin1_general_cs_ai AS EJECUTIVO ";
            Query += "FROM publica.dbo.pershing_cli a ";
            Query += "LEFT JOIN SEBRA.dbo.TBPLFICL b ON ";
            Query += "LTRIM(RTRIM(CONVERT(VARCHAR,CONVERT(INT,SUBSTRING(RUT_CLIENTE,1,LEN(RUT_CLIENTE)-1)))+'-'+SUBSTRING(RUT_CLIENTE,LEN(RUT_CLIENTE),LEN(RUT_CLIENTE)))) = LTRIM(RTRIM(RUT_CLI)) AND sec_rut_cli = 0 ";
            Query += "LEFT JOIN MAESTROS_BCS.dbo.AGENTES_MAESTRO c ON CONVERT(int, LTRIM(RTRIM(a.[Cod#]))) = c.COD_AGENTE ";
            Query += "WHERE REPLACE(a.cta_pershing,'-', '') = @CTAPERSHING";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar("APOLO"));
                command.Parameters.AddWithValue("@CTAPERSHING", ctaPershing);
               

                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    cliente.Nombre = dtReader.GetString(0);
                    //int ConvertRut = int.Parse(dtReader.GetString(1));
                    cliente.rut = dtReader.GetString(1).Substring(0, dtReader.GetString(1).Length - 2);
                    cliente.Dv = dtReader.GetString(1).Substring(dtReader.GetString(1).Length - 1);
                    cliente.CodigoAgente = dtReader.GetString(2);
                    cliente.NombreAgente = dtReader.GetString(3);
                }
            }
            catch (SqlException ex)
            {
                Console.Write(ex.Message);
            }
            catch(Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return cliente;
        }

        public DateTime ObtenerFechaProcesoByFolio(string numFolio)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            DateTime fecha = new DateTime();

            string Query = " SELECT TOP 1 psh_fecha_proceso FROM PUBLICA.dbo.crg_pershing_origin where psh_carga = @NUMFOLIO ";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@NUMFOLIO", numFolio);


                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    fecha = dtReader.GetDateTime(0);
                }
            }
            catch (SqlException ex)
            {
                Console.Write(ex.Message);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return fecha;
        }

        public bool Inserta_Validador_FM(ActivityPershing activity,string NombreUsuario)
        {
            Connection conectar = new Connection();

            string Query = "INSERT INTO [Publica].[dbo].[crg_pershing_historico_ValFM] ";
            Query += "([PROCESS_DATE] ,[ACCOUNT] ,[SHORT_NAME]  ,[TRADE_DATE] ";
            Query += ",[STLMNT_DATE] ,[BUY_SELL] ,[QUANTITY] ,[SYMBOL] ,[CUSIP] ";
            Query += ",[PRICE] ,[ISIN] ,[FULL_NAME] ,[SEC_DESCRIPTION] ,[FLAG] ,[DATE] ,[USER]) ";

            Query += "VALUES (@PROCESS_DATE,@ACCOUNT,@SHORT_NAME,@TRADE_DATE,@STLMNT_DATE,@BUY_SELL,@QUANTITY,@SYMBOL,";
            Query += "@CUSIP,@PRICE,@ISIN,@FULL_NAME,@SEC_DESCRIPTION,@FLAG,@DATE,@USER)";

            try
            {
                foreach (ExcelPershing excelPershing in activity.listadoExcelOmitidos)
                {
                    SqlCommand command = new SqlCommand(Query, conectar.Conectar());

                    command.Parameters.AddWithValue("@PROCESS_DATE", excelPershing.fecha_proceso);
                    command.Parameters.AddWithValue("@ACCOUNT", excelPershing.num_cuenta);
                    command.Parameters.AddWithValue("@SHORT_NAME", excelPershing.nombre_corto);
                    command.Parameters.AddWithValue("@TRADE_DATE", string.IsNullOrEmpty(excelPershing.fecha_trans)? excelPershing.fecha_strlmnt : excelPershing.fecha_trans);
                    command.Parameters.AddWithValue("@STLMNT_DATE", excelPershing.fecha_strlmnt);
                    command.Parameters.AddWithValue("@BUY_SELL", excelPershing.tipo_ope);
                    command.Parameters.AddWithValue("@QUANTITY", Decimal.Parse(excelPershing.cantidad));
                    command.Parameters.AddWithValue("@SYMBOL", excelPershing.simbolo);
                    command.Parameters.AddWithValue("@CUSIP", excelPershing.cusip);
                    command.Parameters.AddWithValue("@PRICE", Decimal.Parse(excelPershing.precio));
                    command.Parameters.AddWithValue("@ISIN", excelPershing.isin);
                    command.Parameters.AddWithValue("@FULL_NAME", excelPershing.Cliente.Nombre);
                    command.Parameters.AddWithValue("@SEC_DESCRIPTION", excelPershing.desc_sec);
                    command.Parameters.AddWithValue("@FLAG", 0);
                    command.Parameters.AddWithValue("@DATE", DateTime.Now);
                    command.Parameters.AddWithValue("@USER", NombreUsuario);
                    command.ExecuteNonQuery();
                }
               
                return true;
            }
            catch (SqlException ex)
            {
                return false;
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool ActualizaEstadoOriginFondosMutuos(ActivityPershing activity)
        {
            Connection conectar = new Connection();

            string Query = "UPDATE PUBLICA.dbo.crg_pershing_origin SET psh_stt = 1 ";
            Query += "where  psh_cusip = @CUSIP ";
            Query += "and psh_fecha_proceso = @FECHAPROCESO ";
                   
            try
            {
                foreach (ExcelPershing excelPershing in activity.listadoExcelOmitidos)
                {
                    SqlCommand command = new SqlCommand(Query, conectar.Conectar());

                    command.Parameters.AddWithValue("@FECHAPROCESO", excelPershing.fecha_proceso);
                    command.Parameters.AddWithValue("@CUSIP", excelPershing.cusip);                
                    command.ExecuteNonQuery();
                }

                return true;
            }
            catch (SqlException ex)
            {
                return false;
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool InsertarFoliosPershing(List<int> Folios,string FechaCarga,double ValorDolar)
        {
            Connection conectar = new Connection();

            string Query = "INSERT INTO PUBLICA.dbo.crg_pershing_folio (fol_fecha, fol_folio, fol_val_dolar) VALUES ";
            Query += "(@FECHACARGA,@FOLIO,@VALORDOLAR)";

            try
            {
                for(int i = 0;i<Folios.Count;i++)
                {
                    SqlCommand command = new SqlCommand(Query, conectar.Conectar());

                    command.Parameters.AddWithValue("@FECHACARGA",DateTime.Parse(FechaCarga));
                    command.Parameters.AddWithValue("@FOLIO",Folios[i]);
                    command.Parameters.AddWithValue("@VALORDOLAR",ValorDolar);
                    command.ExecuteNonQuery();
                }

                return true;
            }
            catch (SqlException ex)
            {
                conectar.Conectar().BeginTransaction().Rollback();
                return false;
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool InsertarEncabezadoFoliosPershing(List<int> Folios, string FechaCarga)
        {
            Connection conectar = new Connection();

            string Query = "INSERT INTO PUBLICA.dbo.crg_pershing_e (prs_e_folio, prs_e_fcarga) VALUES ";
            Query += "(@FOLIO,@FECHACARGA)";

            try
            {
                for (int i = 0; i < Folios.Count; i++)
                {
                    SqlCommand command = new SqlCommand(Query, conectar.Conectar());

                    command.Parameters.AddWithValue("@FECHACARGA", DateTime.Parse(FechaCarga));
                    command.Parameters.AddWithValue("@FOLIO", Folios[i]);
                    command.ExecuteNonQuery();
                }

                return true;
            }
            catch (SqlException ex)
            {
                conectar.Conectar().BeginTransaction().Rollback();
                return false;
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool InsertarPershingLibroOperaciones(ActivityPershing Activity,string usuario)
        {
            Connection conectar = new Connection();

            string Query = "INSERT INTO PUBLICA.dbo.crg_pershing_d (psh_folio, ";
            Query += "psh_fecha_proceso, psh_fecha_limite, psh_tipo_instrumento, ";
            Query += "psh_tipo_ope, psh_cantidad, psh_simbolo, psh_precio, psh_interes, ";
            Query += "psh_monto_neto_base, psh_comision, psh_fees, psh_currency, ";
            Query += "psh_exec_broker, psh_rut_cli, psh_dv, ";
            Query += "psh_nombre_cliente, psh_cod_agente, psh_nombre_agente, ";
            Query += "psh_usuario, psh_fecha_c, psh_folio_origen, psh_monto_neto_trans, psh_num_cuenta, psh_isin, psh_estado_item) VALUES ";

            Query += "(@FOLIO,@FECHAPROCESO,@FECHALIMITE,@INSTRUMENTO,@TIPOOPERACION,@CANTIDAD,@SIMBOLO,@PRECIO,@INTERES,";
            Query += "@MONTONETOBASE,@COMISION,@FEES,@CURRENCY,@EXECBROKER,@RUTCLI,@DV,@NOMBRECLI,@CODAGENTE,@NOMAGENTE,";
            Query += "@USUARIO,@FECHACREACION,@FOLIOORIGEN,@MONTONETROTRANS,@NUMCUENTA,@ISIN,@ESTADOITEMS)";
            try
            {
                foreach (ExcelPershing pershing in Activity.listadoExcelPershing)
                {
                    SqlCommand command = new SqlCommand(Query, conectar.Conectar());

                    command.Parameters.AddWithValue("@FOLIO",pershing.folio);
                    command.Parameters.AddWithValue("@FECHAPROCESO",DateTime.Parse(pershing.fecha_proceso));
                    command.Parameters.AddWithValue("@FECHALIMITE", DateTime.Parse(pershing.fecha_strlmnt));
                    command.Parameters.AddWithValue("@INSTRUMENTO",pershing.tipo_cuenta);
                    command.Parameters.AddWithValue("@TIPOOPERACION",pershing.tipo_ope);
                    command.Parameters.AddWithValue("@CANTIDAD",Decimal.Parse(pershing.cantidad));
                    command.Parameters.AddWithValue("@SIMBOLO",pershing.simbolo);
                    command.Parameters.AddWithValue("@PRECIO",Decimal.Parse(pershing.precio));
                    command.Parameters.AddWithValue("@INTERES", Decimal.Parse(pershing.interes));
                    command.Parameters.AddWithValue("@MONTONETOBASE", Decimal.Parse(pershing.monto_neto_base));
                    command.Parameters.AddWithValue("@COMISION", Decimal.Parse(pershing.comision));
                    command.Parameters.AddWithValue("@FEES",Decimal.Parse(pershing.impuestos));
                    command.Parameters.AddWithValue("@CURRENCY",pershing.currency);
                    command.Parameters.AddWithValue("@EXECBROKER",pershing.exec_broker);
                    command.Parameters.AddWithValue("@RUTCLI",pershing.Cliente.rut);
                    command.Parameters.AddWithValue("@DV",pershing.Cliente.Dv);
                    command.Parameters.AddWithValue("@NOMBRECLI",pershing.Cliente.Nombre);
                    command.Parameters.AddWithValue("@CODAGENTE",pershing.Cliente.CodigoAgente);
                    command.Parameters.AddWithValue("@NOMAGENTE",pershing.Cliente.NombreAgente);
                    command.Parameters.AddWithValue("@USUARIO", usuario);
                    command.Parameters.AddWithValue("@FECHACREACION",DateTime.Now);
                    command.Parameters.AddWithValue("@FOLIOORIGEN",pershing.Id);
                    command.Parameters.AddWithValue("@MONTONETROTRANS", Decimal.Parse(pershing.monto_neto_trans));
                    command.Parameters.AddWithValue("@NUMCUENTA",pershing.num_cuenta);
                    command.Parameters.AddWithValue("@ISIN",pershing.isin);
                    command.Parameters.AddWithValue("@ESTADOITEMS",1);               
                    command.ExecuteNonQuery();
                }

                return true;
            }
            catch (SqlException ex)
            {
                conectar.Conectar().BeginTransaction().Rollback();
                return false;
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool VerificarCuentasTanner(string NumCuenta)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            bool resultado = false;

            string Query = " SELECT * FROM PUBLICA.DBO.pershing_cli_ctas_tanner WHERE REPLACE(cta_pershing,'-', '') = @NUMCUENTA";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@NUMCUENTA", NumCuenta);


                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    resultado = true;
                }
            }
            catch (SqlException ex)
            {
                Console.Write(ex.Message);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return resultado;
        }

        public DataTable ObtenerLibroOperacionesPershing(string desde,string hasta)
        {
            Connection conectar = new Connection();
            DataTable resultado = new DataTable();
            SqlDataAdapter Dt = new SqlDataAdapter();

            string Query = "SELECT CONVERT(varchar(10), psh_fecha_proceso, 103) AS psh_fecha_proceso, CONVERT(varchar(10), psh_fecha_proceso, 103) AS psh_fecha_comm, psh_folio, psh_tipo_ope, psh_tipo_instrumento, ";
            Query += "psh_simbolo,psh_exec_broker,psh_cantidad, psh_currency, psh_precio, ";
            Query += "(CASE WHEN psh_interes < 0 then psh_interes * -1 ELSE psh_interes END) as psh_interes_dev, ";
            Query += "CAST(CAST(psh_cantidad AS decimal(16, 4))*CAST(psh_precio AS decimal(16, 4)) AS decimal(16, 4)) AS psh_monto_bruto, ";
            Query += "(CASE WHEN psh_comision < 0 then (psh_comision*-1) ELSE psh_comision END) as psh_comision, psh_fees, 0 as psh_sec_fees, ";
            Query += "psh_comm_tanner, psh_gastos_tanner, (CASE WHEN psh_monto_neto_trans < 0 THEN psh_monto_neto_trans * -1 ELSE psh_monto_neto_trans END) AS psh_monto_neto_transa, ";
            Query += "psh_nombre_agente, RTRIM(LTRIM(CONVERT(varchar, replace(psh_rut_cli, '-', ''))+'-'+psh_dv)) as psh_rut, ";
            Query += "psh_nombre_cliente, ";
            Query += "CASE WHEN psh_estado_item = 0 THEN 'Anulada' ";
            Query += "WHEN psh_estado_item = 1 THEN 'Disponible para contrato' ";
            Query += "WHEN psh_estado_item = 2 THEN 'Finalizada' ";
            Query += "END as estado ";
            Query += "FROM PUBLICA.dbo.crg_pershing_d ";
            Query += "WHERE psh_fecha_proceso BETWEEN @DESDE  AND @HASTA ";
            Query += "ORDER BY psh_folio ASC ";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@DESDE", DateTime.Parse(desde));
                command.Parameters.AddWithValue("@HASTA", DateTime.Parse(hasta));

                Dt = new SqlDataAdapter(command);
                Dt.Fill(resultado);    
            }
            catch (SqlException ex)
            {
                Console.Write(ex.Message);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                conectar.Desconectar();
                Dt = null;
            }

            return resultado;
        }

        public DataTable ObtenerRegistroRecepcionesPershing(string desde, string hasta)
        {
            Connection conectar = new Connection();
            DataTable resultado = new DataTable();
            SqlDataAdapter Dt = new SqlDataAdapter();

            string Query = "SELECT RTRIM(LTRIM(CONVERT(varchar, replace(psh_rut_cli, '-', ''))+'-'+psh_dv)) as psh_rut, ";
            Query += "psh_nombre_cliente, psh_nombre_agente, psh_folio, CONVERT(varchar(10), psh_fecha_proceso, 103) AS psh_fecha_proceso, ";
            Query += "psh_tipo_ope, psh_tipo_instrumento, psh_simbolo, psh_exec_broker, 'EE.UU.' as psh_mercado, psh_cantidad, psh_currency, ";
            Query += "psh_precio, CONVERT(varchar(10), psh_fecha_limite, 103) AS psh_fecha_limite, ";
            Query += "CASE WHEN psh_estado_item = 0 THEN 'Anulada' ";
            Query += "WHEN psh_estado_item = 1 THEN 'Disponible para contrato' ";
            Query += "WHEN psh_estado_item = 2 THEN 'Finalizada' ";
            Query += "END as estado ";
            Query += "FROM PUBLICA.dbo.crg_pershing_d ";
            Query += "WHERE psh_fecha_proceso BETWEEN @DESDE AND @HASTA ";
            Query += " ORDER BY psh_folio ASC ";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@DESDE", DateTime.Parse(desde));
                command.Parameters.AddWithValue("@HASTA", DateTime.Parse(hasta));

                Dt = new SqlDataAdapter(command);
                Dt.Fill(resultado);
            }
            catch (SqlException ex)
            {
                Console.Write(ex.Message);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                conectar.Desconectar();
                Dt = null;
            }

            return resultado;
        }

        #region Generacion de Contratos

        public List<DateTime> ListadoFechasContratosAGenerar()
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            List<DateTime> ListadoFechas = new List<DateTime>();

            string Query = "SELECT DISTINCT psh_fecha_proceso as _fecha FROM PUBLICA.dbo.crg_pershing_d WHERE psh_estado_item = 1";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                dtReader = command.ExecuteReader();

                while (dtReader.Read())
                {
                    ListadoFechas.Add(dtReader.GetDateTime(0));
                }
            }
            catch (SqlException ex)
            {
                Console.Write(ex.Message);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return ListadoFechas;
        } 

        public List<Contrato> ListadoOperacionesPendientes(DateTime fechaProceso)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            List<Contrato> ListadoContratos = new List<Contrato>();

            string Query = "SELECT psh_fecha_proceso, psh_folio, psh_simbolo, psh_currency, ";
            Query += "(CASE WHEN LEN(psh_simbolo) <=4 THEN 'Accion' ELSE 'Bono' END) AS psh_tipo_inst, psh_isin, psh_cantidad as psh_nominales, ";
            Query += "psh_precio, psh_tipo_ope, psh_comision, psh_num_cuenta, ";
            Query += "(CONVERT(nvarchar(12), psh_rut_cli)+'-'+psh_dv) as psh_rut, psh_nombre_cliente,psh_id ";
            Query += "FROM PUBLICA.dbo.crg_pershing_d WHERE psh_estado_item = 1 ";
            Query += "AND psh_fecha_proceso = @FECHAPROCESO ";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@FECHAPROCESO", fechaProceso);
                dtReader = command.ExecuteReader();               

                while (dtReader.Read())
                {
                    Contrato contrato = new Contrato();
                    contrato.FechaProceso = dtReader.GetDateTime(0);
                    contrato.Folio = dtReader.GetInt64(1);
                    contrato.Simbol = dtReader.GetString(2);
                    contrato.Currency = dtReader.GetString(3);
                    contrato.TipoInstrumento = dtReader.GetString(4);
                    contrato.Isin = dtReader.GetString(5);
                    contrato.Nominales = dtReader.GetDecimal(6);
                    contrato.Precio = dtReader.GetDecimal(7);
                    contrato.TipoOperacion = dtReader.GetString(8);
                    contrato.Comision = dtReader.GetDecimal(9);
                    contrato.NumeroCuenta = dtReader.GetString(10);
                    contrato.RutCliente = dtReader.GetString(11);
                    contrato.NombreCliente = dtReader.GetString(12);
                    contrato.Id = dtReader.GetInt32(13);
                    ListadoContratos.Add(contrato);
                }
            }
            catch (SqlException ex)
            {
                Console.Write(ex.Message);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return ListadoContratos;
        }

        public List<Contrato> ListadoContratosPendientes(DateTime fechaProceso)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            List<Contrato> ListadoContratos = new List<Contrato>();

            string Query = "SELECT DISTINCT psh_num_cuenta,psh_nombre_cliente,CONVERT(VARCHAR(10),psh_rut_cli) +'-' +psh_dv as rut ";
            Query += "FROM PUBLICA.dbo.crg_pershing_d where psh_estado_item = 1 ";
            Query += "and psh_fecha_proceso = @FECHAPROCESO ";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@FECHAPROCESO",fechaProceso);
                dtReader = command.ExecuteReader();

                while (dtReader.Read())
                {
                    Contrato contrato = new Contrato();           
                    contrato.NumeroCuenta = dtReader.GetString(0);
                    contrato.RutCliente = dtReader.GetString(2);
                    contrato.NombreCliente = dtReader.GetString(1);                  
                    ListadoContratos.Add(contrato);
                }
            }
            catch (SqlException ex)
            {
                Console.Write(ex.Message);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return ListadoContratos;
        }

        #endregion

        public bool ActualizarMarcaContrato(long folio)
        {
            Connection conectar = new Connection();

            string Query = "UPDATE PUBLICA.dbo.crg_pershing_d SET psh_estado_item = 2 WHERE psh_folio = @FOLIO ";
            try
            {             
                    SqlCommand command = new SqlCommand(Query, conectar.Conectar());

                    command.Parameters.AddWithValue("@FOLIO", folio);                 
                    command.ExecuteNonQuery();

                return true;
            }
            catch (SqlException ex)
            {
                return false;
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public bool ActualizarMarcaContratosOrigen(List<ExcelPershing> ListadoContratos)
        {
            Connection conectar = new Connection();

            string Query = "UPDATE PUBLICA.dbo.crg_pershing_origin SET psh_stt = 1 WHERE psh_id = @FOLIO";
            try
            {            
                foreach (ExcelPershing _pershing in ListadoContratos)
                {
                    SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                    command.Parameters.AddWithValue("@FOLIO", _pershing.Id);
                    command.ExecuteNonQuery();
                }
                          
                return true;
            }
            catch (SqlException ex)
            {
                return false;
            }
            finally
            {
                conectar.Desconectar();
            }
        }

        public int[] OperacionesPendientesParaProcesar(DateTime fecha)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            int[] OperacionesPendientes = new int[3];

            string Query = "SELECT  'Renta Variable' as Descripcion,COUNT(psh_tipo_instrumento) as Contador FROM PUBLICA.dbo.crg_pershing_origin where psh_tipo_instrumento = 'Renta Variable' and psh_stt=0 and psh_fecha_proceso = @FECHAPROCESO UNION ";
                   Query += "SELECT  'Renta Fija' as Descripcion,COUNT(psh_tipo_instrumento) as Contador FROM PUBLICA.dbo.crg_pershing_origin where psh_tipo_instrumento = 'Renta Fija' and psh_stt=0 and psh_fecha_proceso = @FECHAPROCESO UNION ";
                   Query += "SELECT  'Fondos Mutuos' as Descripcion,COUNT(psh_tipo_instrumento) as Contador FROM PUBLICA.dbo.crg_pershing_origin where psh_tipo_instrumento = 'Fondos Mutuos' and psh_stt=0 and psh_fecha_proceso = @FECHAPROCESO ";
       
            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@FECHAPROCESO", fecha);
                dtReader = command.ExecuteReader();

                int indice = 0;

                while (dtReader.Read())
                {
                    OperacionesPendientes[indice] = dtReader.GetInt32(1);
                    indice++;
                }
            }
            catch (SqlException ex)
            {
                Console.Write(ex.Message);
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return OperacionesPendientes;
        }

    }
}
