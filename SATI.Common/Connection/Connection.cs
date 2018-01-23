using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace SATI.Common
{
    public class Connection
    {
        private SqlConnection _connection = new SqlConnection();


        public SqlConnection Conectar(string Base = "APOLO-DES")  
        {
            try
            {
                if(_connection.State.Equals(ConnectionState.Open))
                {
                    return _connection;
                }else
                {
                    string ConexionString = ConfigurationManager.ConnectionStrings["sql:apolo"].ToString().Replace("@DATABASE", Base);
                    _connection.ConnectionString = ConexionString;
                    _connection.Open();
                    return _connection;
                }
            }catch(SqlException ex)
            {
                throw new ArgumentException($"Error al Desconectar base de datos Detalle : {ex.Message}");
            }
        }

        public void Desconectar()
        {
            try
            {
                if (_connection.State.Equals(ConnectionState.Open))
                {
                    _connection.Close();
                }
            }catch(Exception ex)
            {
                throw new ArgumentException($"Error al Desconectar base de datos Detalle : {ex.Message}");
            }
        }

    }
}
