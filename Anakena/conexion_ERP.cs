using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common;
using System.Windows.Forms;

namespace Anakena
{
    class conexionERP
    {
        private SqlConnection Con = new SqlConnection("Data Source=SERVIDORERP;Initial Catalog=anakena;User id=sa;Password=Pall2015;Integrated Security = False"); // Obj Conexion

        public void  Conexion()

        {
           
        }

        public void Abrir() // Metodo para Abrir la Conexion
        {
           Con.Open();
        }

        public void Cerrar() // Metodo para Cerrar la Conexion
        {
            Con.Close();
        }
        public SqlConnection  getConexion()// Metodo devuelve un objeto Conexion
        {
            return Con;
        }
    }
}
