using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using SATI.Common.Common;
using SATI.Common.Entities;
using SATI.Common.Entities.Excel;
using SATI.Common.Entities.Pershing;
namespace SATI.Services.Services.Pershing
{
    public class PershingService
    {
        private PershingCommon _pershingCommon = new PershingCommon();
        private static List<int> FoliosAsignados;

        #region Procesar Activity Pershing
        public ActivityPershing ProcesarActivityPershing(ActivityPershing activityResumen)
        {
            bool respuestaLimpia = true;
            try
            {
                if (activityResumen.Duplicado)
                {
                    respuestaLimpia = LimpiarTablasPershing(activityResumen.FechaCarga);
                }

                _pershingCommon.GuardarExcelPershingBD(activityResumen.listadoExcelPershing);
                _pershingCommon.CargaRentaVariablePershing(activityResumen.NumeroFolio);
                _pershingCommon.CargaFondosMutuosPershing(activityResumen.NumeroFolio);
                _pershingCommon.CargaRentaFijaPershing(activityResumen.NumeroFolio);

                activityResumen = ObtenerActivity(activityResumen.NumeroFolio, activityResumen.Instrumento);
            }
            catch (Exception e)
            {
                PershingError error = new PershingError();
                error.Titulo = "Error Procesar Pershing";
                error.Mensaje = "Se Produjo un Error al Intentar procesar el archivo de activity pershing";
                error.descripcion = e.Message;
                activityResumen.Error = error;
            }

            return activityResumen;
        }
        #endregion
        #region Obtener Activity Pershing y Listados del Mismo
        public ActivityPershing ObtenerActivity(string NumeroFolio, string Instrumento)
        {
            ActivityPershing activity = new ActivityPershing();
            try
            {
                activity.FechaCarga = _pershingCommon.ObtenerFechaProcesoByFolio(NumeroFolio).Date.ToString();
                activity.valorDolar = _pershingCommon.ObtenerValorDolarByFecha(activity.FechaCarga);
                activity.listadoExcelPershing = _pershingCommon.Carga_Origen_excel(NumeroFolio, Instrumento);
                activity.NumeroFolio = NumeroFolio;
                activity.Instrumento = Instrumento;
                activity.RegistroAnulados = 0;
                activity.TotalComCLP = "0";
                activity.TotalComUS = 0;
                activity.TotalRegistros = activity.listadoExcelPershing.Count;
                activity.Duplicado = _pershingCommon.ExisteFechaProcesoPershing(activity.FechaCarga);

                if (activity.listadoExcelPershing.Count > 0)
                {
                    activity.TotalRegistros = activity.listadoExcelPershing.Count;
                    activity.RegistroAnulados = int.Parse(_pershingCommon.ObtenerCantidadPershingByFecha(activity.FechaCarga));
                }
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
            }

            return ObtenerTotalesPershing(activity);
        }

        public ActivityPershing ObtenerTotalesPershing(ActivityPershing activity)
        {
            List<ExcelPershing> ListadoPershingOmitidos = new List<ExcelPershing>();
            List<ExcelPershing> NuevoListadoPershing = new List<ExcelPershing>();

            foreach (ExcelPershing pershing in activity.listadoExcelPershing)
            {
                pershing.Cliente = _pershingCommon.ObtenerClientePershingByCta(pershing.num_cuenta);

                if (pershing.simbolo != "" && pershing.tipo_ope != "")
                {
                                    
                    if (pershing.tipo_ope.Trim() == "S")
                    {
                        pershing.tipo_ope = "Venta";
                    }
                    else if (pershing.tipo_ope.Trim() == "B")
                    {
                        pershing.tipo_ope = "Compra";
                    }
                  
                    bool IsFondosMutuos = _pershingCommon.ValidarRegistroFondosMutuos(pershing.nombre_corto, pershing.cusip, pershing.num_cuenta, pershing.fecha_trans);

                    if (!IsFondosMutuos && pershing.simbolo.IndexOf("FM-") == 0)
                        ListadoPershingOmitidos.Add(pershing);
                    else
                        NuevoListadoPershing.Add(pershing);
                }
                else
                {                 
                        ListadoPershingOmitidos.Add(pershing);               
                }
            }

            activity.listadoExcelPershing = NuevoListadoPershing;
            activity.listadoExcelOmitidos = ListadoPershingOmitidos;

            foreach (ExcelPershing excelpershing in NuevoListadoPershing)
            {
                activity.TotalComUS += double.Parse(excelpershing.comision);
            }

            activity.TotalComCLP = string.Format("{0:n0}", int.Parse(Math.Round(activity.TotalComUS * activity.valorDolar).ToString())).Replace(",", ".");
            activity.TotalRegistros = NuevoListadoPershing.Count;
            activity.RegistroAnulados = ListadoPershingOmitidos.Count;
            return activity;
        }

        private List<ExcelPershing> ObtenerListadoPershingSeleccionados(List<ExcelPershing> listadoPershing, string[] pershingCod)
        {
            List<ExcelPershing> ListadoSeleccionados = new List<ExcelPershing>();
            foreach (ExcelPershing pershing in listadoPershing)
            {
                if (pershingCod.Where(p => p == pershing.Id).ToList().Count > 0)
                {
                    ListadoSeleccionados.Add(pershing);
                }
            }

            return ListadoSeleccionados;
        }
        #endregion
        #region Limpiadores Tablas Pershing
        private bool LimpiarTablasPershing(string fecha)
        {
            bool resultado = _pershingCommon.EliminaFechasRepetidas(fecha);
            if (resultado)
            {
                resultado = _pershingCommon.LimpiaTablasExcel(fecha);
            }

            return resultado;
        }

        #endregion
        #region Cargar Pershing Seleccionados en la Tabla de la Vista
        public ActivityPershing CargarPershingSeleccionados(ActivityPershing activity, string[] pershingCod,string usuario)
        {
            FoliosAsignados = new List<int>();
            ActivityPershing activityDetalle = activity;           
            int MaxFolio = 0;
            int MinFolio = 999999;
                      
            if (pershingCod != null)
            {
                int NuevoFolio = int.Parse(_pershingCommon.ObtenerUltimoFolioAsignado().ToString()) + 1;
                DateTime FechaCarga = DateTime.Now;

                List<ExcelPershing> listadoSeleccionados = ObtenerListadoPershingSeleccionados(activity.listadoExcelPershing, pershingCod);

                foreach (ExcelPershing pershing in listadoSeleccionados)
                {
                                                 
                    if (NuevoFolio < MinFolio)
                    {
                        MinFolio = NuevoFolio;
                    }
                    if (NuevoFolio > MaxFolio)
                    {
                        MaxFolio = NuevoFolio;
                    }

                    pershing.folio = NuevoFolio.ToString();
                    FoliosAsignados.Add(NuevoFolio);
                    NuevoFolio += 1;                               
                }

                InsertarFondosMutuosPendientes(activity.NumeroFolio, usuario);

                activity.listadoExcelPershing = listadoSeleccionados;
                activity.NumeroFolio = string.Format("{0} - {1}", MinFolio, MaxFolio);
                activity.TotalComUS = 0;
                activity = ObtenerTotalesPershing(activity);
            }
            else
            {
                PershingError error = new PershingError();
                error.Titulo = "Error al cargar pershing";
                error.Mensaje = "No se encontraron pershing seleccionados.";
                error.descripcion = "debe seleccionar por lo menos 1 pershing para poder seguir con el proceso";
                activity.Error = error;
            }
                      
            return activity;
        }

      
        #endregion
        #region Methodos Activity Fondos Mutuos
        public void ActualizarHistoricoFM(string simbol, string numeroCuenta, string cusip, string processDate, string shortName, string TradeDate, string StlmntDate, string fecha, string usuario)
        {
            string valorMod = string.Empty;
            valorMod = simbol;
            int result = 0;

            if (int.TryParse(valorMod, out result))
            {
                return;
            }
            else
            {
                _pershingCommon.ActualizaLogValidaFMExcel(simbol, cusip, numeroCuenta);
                _pershingCommon.ActualizaLogValidaFMExcelHistorico(simbol, cusip, numeroCuenta);
                _pershingCommon.InsertaLogValidaFMExcelHistorico(processDate, numeroCuenta, shortName, TradeDate, StlmntDate, simbol, cusip, fecha, usuario);
            }
        }

        public void InsertarFondosMutuosPendientes(string numFolio, string usuario)
        {
            ActivityPershing ActivityFMPendientes = ObtenerActivity(numFolio, "Fondos Mutuos");
            bool Resultado = _pershingCommon.Inserta_Validador_FM(ActivityFMPendientes, usuario);

            if (Resultado)
                _pershingCommon.ActualizaEstadoOriginFondosMutuos(ActivityFMPendientes);
        }
        #endregion   

        public CargaPershingResult CargarLibroOperaciones(ActivityPershing activity,string usuario)
        {
            CargaPershingResult ResultPershing = new CargaPershingResult();
            try
            {
                _pershingCommon.InsertarFoliosPershing(FoliosAsignados, activity.FechaCarga, activity.valorDolar);
                _pershingCommon.InsertarEncabezadoFoliosPershing(FoliosAsignados, activity.FechaCarga);
                _pershingCommon.InsertarPershingLibroOperaciones(activity, usuario);
                _pershingCommon.ActualizarMarcaContratosOrigen(activity.listadoExcelPershing);

                ResultPershing.Mensaje = $"Las Operaciones de {activity.Instrumento} Fueron Cargadas Exitosamente. ";
                ResultPershing.Instrumento = activity.Instrumento;
                int[] OperacionesPendientes = _pershingCommon.OperacionesPendientesParaProcesar(DateTime.Parse(activity.FechaCarga));
                ResultPershing.FFMMPendiente = OperacionesPendientes[2];
                ResultPershing.RentaFijaPendiente = OperacionesPendientes[1];
                ResultPershing.RentaVariablePendiente = OperacionesPendientes[0];
            }
            catch(Exception e)
            {
                Console.Write(e.Message);
            }

            return ResultPershing;
        }
    }
}

