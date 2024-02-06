using System;
using System.Data;
using System.Data.OleDb;

class Program
{
    static void Main(string[] args)
    {
        // Ruta del archivo Excel
        string archivoExcel = @"C:\Users\kevin\OneDrive\Escritorio\Boca.xlsx";

        // Cadena de conexión 
        string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={archivoExcel};Extended Properties='Excel 12.0 XML';";

        // Consulta para insertar datos
        string insertQuery = "INSERT INTO [Hoja1$] (Nombre, Apellido, Dorsal) VALUES (@nombre, @apellido, @dorsal)";

        // Datos para insertar
        string jugador1 = "Lionel";
        string apellido1 = "Messi";
        int dorsal2 = 19; 

        string jugador2 = "Neymar";
        string apellido2 = "Neymar";
        int dorsal = 10;

        // Ejecutar inserciones
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
            {
                // Parámetros
                command.Parameters.AddWithValue("@nombre", jugador1);
                command.Parameters.AddWithValue("@apellido", apellido1);
                command.Parameters.AddWithValue("@dorsal", dorsal2);
                
                

                // Abrir conexión y ejecutar la inserción
                connection.Open();
                command.ExecuteNonQuery();
                //limpio parametros
                command.Parameters.Clear();
                //inserto datos del jugador nro
                
                command.Parameters.AddWithValue("@jugador2", jugador2);
                command.Parameters.AddWithValue("@apellido2", apellido2);
                command.Parameters.AddWithValue("@dorsal", dorsal);
                command.ExecuteNonQuery();

            }
        }
    }
}




