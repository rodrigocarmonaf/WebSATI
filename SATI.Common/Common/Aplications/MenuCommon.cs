using System;
using System.Collections.Generic;
using SATI.Common.Entities;
using System.Data.SqlClient;
using SATI.Common.Helper;

namespace SATI.Common.Aplications
{
    public class MenuCommon
    {
        public List<Menu> ListadoMenuAplicacionsUsuario(string Id,string Modulo)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            List<Menu> resultado = new List<Menu>();
            string Query = "";
            Query = "SELECT DISTINCT a.apl_nombre,a.apl_enlace FROM  [USUARIO].[dbo].[t_aplicacion] a ";
            Query += "LEFT JOIN [USUARIO].[dbo].[t_modulo] c ON a.apl_mdl_id = c.mdl_id ";
            Query += "LEFT JOIN [USUARIO].[dbo].[t_permiso] b ON a.apl_id = b.prm_appid ";
            Query += "WHERE b.prm_sysid = 18 and b.prm_usr = 4 and b.prm_stt = 1 and c.mdl_nombre = @Modulo ";
            Query += "ORDER BY a.apl_nombre DESC ";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@Id", Id);
                command.Parameters.AddWithValue("@Modulo", Modulo);

                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    Menu menu = new Menu();
                    menu.Nombre = dtReader.GetString(0);
                    menu.Enlace = dtReader.GetString(1);
                    menu.items = ListadoItemsMenuAplicacionsUsuario(Id, menu.Nombre, Modulo);
                    resultado.Add(menu);
                }
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Modulos Aplicacion Detalle : {ex.Message}");
            }
            finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }

            return resultado;
        }

        private List<ItemsMenu> ListadoItemsMenuAplicacionsUsuario(string Id, string Aplicacion, string Modulo)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            List<ItemsMenu> resultado = new List<ItemsMenu>();
            string Query = "";
            Query = "SELECT DISTINCT a.scc_nombre,a.scc_enlace FROM [USUARIO].[dbo].[t_seccion] a ";
            Query += "LEFT JOIN [USUARIO].[dbo].[t_aplicacion] c ON a.scc_apl_id = c.apl_id ";
            Query += "LEFT JOIN [USUARIO].[dbo].[t_modulo] d ON c.apl_mdl_id = d.mdl_id ";
            Query += "LEFT JOIN [USUARIO].[dbo].[t_permiso] b ON  a.scc_id = b.prm_secid ";
            Query += "WHERE b.prm_sysid = 18 and b.prm_usr = @Id and b.prm_stt = 1 and c.apl_nombre = @Aplicacion and d.mdl_nombre = @Modulo ";

            try
            {
                SqlCommand command = new SqlCommand(Query,conectar.Conectar());
                command.Parameters.AddWithValue("@Id", Id);
                command.Parameters.AddWithValue("@Modulo", Modulo);
                command.Parameters.AddWithValue("@Aplicacion", Aplicacion);

                dtReader = command.ExecuteReader();
                while (dtReader.Read())
                {
                    ItemsMenu menu = new ItemsMenu();
                    menu.Nombre = dtReader.GetString(0);
                    menu.Enlace = dtReader.GetString(1);
                    resultado.Add(menu);
                }
            }
            catch (SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Modulos Aplicacion Detalle : {ex.Message}");
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                dtReader.Close();
                conectar.Conectar().Close();
            }

            return resultado;
        }

    }
}
