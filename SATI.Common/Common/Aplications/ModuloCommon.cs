using System;
using System.Collections.Generic;
using SATI.Common.Entities;
using System.Data.SqlClient;
using SATI.Common.Helper;

namespace SATI.Common.Aplications
{
    public class ModuloCommon
    {
        public List<Modulos> ListadoModulosUsuario(string Id,string moduloActivo)
        {
            Connection conectar = new Connection();
            SqlDataReader dtReader = null;
            List<Modulos> resultado = new List<Modulos>();
            string Query = "";
            Query = "SELECT DISTINCT a.mdl_id,a.mdl_titulo,a.mdl_enlace ";
            Query += "FROM [USUARIO].[dbo].[t_modulo] a ";
            Query += "LEFT JOIN [USUARIO].[dbo].[t_permiso] b ON b.prm_mdlid  = a.mdl_id ";
            Query += "WHERE b.prm_sysid = 18 and b.prm_usr = @Id and b.prm_stt = 1 ";

            try
            {
                SqlCommand command = new SqlCommand(Query, conectar.Conectar());
                command.Parameters.AddWithValue("@Id", Id);

                dtReader = command.ExecuteReader();
                while(dtReader.Read())
                {
                    Modulos modulo = new Modulos();
                    modulo.Id = dtReader.GetInt32(0);
                    modulo.Titulo = dtReader.GetString(1);

                    string[] AcctionController = ModuloHelper.ObtenerActionAndController(dtReader.GetString(2));
                    modulo.Controller = AcctionController[0];
                    modulo.Action = AcctionController[1];

                    if(dtReader.GetString(1).Equals(moduloActivo))
                    {
                        modulo.Activo = "active";
                    }

                    resultado.Add(modulo);
                }
            }catch(SqlException ex)
            {
                throw new ArgumentException($"Error al Consultar Modulos Aplicacion Detalle : {ex.Message}");
            }finally
            {
                dtReader.Close();
                conectar.Desconectar();
            }
                                    
            return resultado;
        }
    }
}
