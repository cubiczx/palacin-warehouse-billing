namespace WarehouseBilling;

using System;
using System.IO;
using System.Data;
using System.Globalization;
using System.Collections.Generic;
using Microsoft.Data.SqlClient;
using ExcelDataReader;
using Microsoft.Extensions.Configuration;

public class BillingEngine
{
    public static void Main(string[] args)
    {
        // Fijamos la cultura española para que :C2 muestre € en vez de ¤
        CultureInfo.CurrentCulture = new CultureInfo("es-ES");
        CultureInfo.CurrentUICulture = new CultureInfo("es-ES");

        // Requisito de codificación para ExcelDataReader en .NET 6+
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        var config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        string connectionString = config.GetConnectionString("Almacen")
            ?? throw new InvalidOperationException("No se encontró la cadena de conexión 'Almacen' en appsettings.json");

        // Resolución de la ruta del Excel:
        // 1. Prioridad: Argumento de consola (ej: dotnet run "C:/Datos/Movimientos_Agosto.xlsx")
        // 2. Segunda opción: Configuración definida en appsettings.json
        // 3. Valor por defecto: "Movimientos.xlsx"
        string excelFilePath = args.Length > 0 
            ? args[0] 
            : config["ExcelFilePath"] ?? "Movimientos.xlsx";

        Console.WriteLine("1. Importar Excel | 2. Calcular Facturación");
        var opcion = Console.ReadLine()?.Trim();

        if (opcion == "1") 
        {
            ImportarExcel(connectionString, excelFilePath); 
        } 
        else if (opcion == "2")
        {
            using var conn = new SqlConnection(connectionString);
            conn.Open();

            // 1. Mostrar clientes disponibles en la BD de forma más clara
            List<string> clientesDisponibles = new List<string>();
            using (var cmd = new SqlCommand("SELECT CodCliente, Nombre FROM Clientes", conn))
            using (var rdr = cmd.ExecuteReader())
            {
                Console.WriteLine("\n--- Clientes Registrados ---");
                while (rdr.Read())
                {
                    string c = rdr.GetString(0);
                    string n = rdr.GetString(1);
                    clientesDisponibles.Add(c);
                    Console.WriteLine($"- Código: [{c}]  Nombre: {n}");
                }
            }

            if (clientesDisponibles.Count == 0)
            {
                Console.WriteLine("No hay clientes registrados. Por favor, importe un archivo Excel primero (Opción 1).");
                return;
            }

            // 2. Pedir y validar código de cliente
            Console.Write("\nIntroduce el código del cliente (indique el valor entre corchetes, ej. C1): ");
            string codCliente = Console.ReadLine()?.ToUpper().Trim() ?? "";

            if (!clientesDisponibles.Contains(codCliente))
            {
                Console.WriteLine($"Error: El código '{codCliente}' no existe en el sistema.");
                return;
            }

            // 3. Consultar el rango real de años disponibles en los movimientos del cliente
            int minYear = 0;
            int maxYear = 0;
            
            using (var cmdYear = new SqlCommand("SELECT MIN(YEAR(Fecha)), MAX(YEAR(Fecha)) FROM Movimientos WHERE CodCliente = @Cliente", conn))
            {
                cmdYear.Parameters.AddWithValue("@Cliente", codCliente);
                using var rdrYear = cmdYear.ExecuteReader();
                if (rdrYear.Read() && !rdrYear.IsDBNull(0))
                {
                    minYear = rdrYear.GetInt32(0);
                    maxYear = rdrYear.GetInt32(1);
                    Console.WriteLine($"[Info] Rango de años con movimientos para este cliente: {minYear} - {maxYear}.");
                }
                else
                {
                    Console.WriteLine("[Aviso] Este cliente no tiene movimientos registrados en la base de datos.");
                    return;
                }
            }

            // 4. Pedir año y validar estrictamente contra el rango devuelto por la base de datos
            Console.Write($"Introduce el año (entre {minYear} y {maxYear}): ");
            if (!int.TryParse(Console.ReadLine(), out int year) || year < minYear || year > maxYear)
            {
                Console.WriteLine($"Error: El año introducido no es válido o está fuera del rango disponible ({minYear}-{maxYear}).");
                return;
            }

            Console.Write("Introduce el mes (1-12): ");
            if (!int.TryParse(Console.ReadLine(), out int month) || month < 1 || month > 12)
            {
                Console.WriteLine("Error: El mes debe estar comprendido entre 1 y 12.");
                return;
            }

            CalcularFacturacion(connectionString, codCliente, year, month); 
        }
        else
        {
            Console.WriteLine("Error: Opción no válida. Debe introducir '1' para importar o '2' para calcular.");
        }
    }

    private static void ImportarExcel(string connString, string filePath) 
    {
        // =========================================================================================
        // ESCENARIO B (Implementado): El Excel recibido es un fichero acumulativo ("foto completa" histórica).
        // Contiene tanto los meses anteriores ya procesados (ej. meses 1 a 7) como nuevos meses futuros 
        // (ej. agosto, septiembre). 
        // Para evitar duplicidades de datos y desincronización en los cálculos de saldo de estanterías,
        // aplicamos un vaciado completo de la tabla Movimientos (TRUNCATE) antes de verter los nuevos datos 
        // mediante SqlBulkCopy.
        // 
        // *Nota alternativa (Escenario A - Ficheros incrementales):* Si el proveedor enviase ficheros separados 
        // y puramente incrementales (aportando solo los datos del nuevo mes sin repetir el histórico), no se 
        // requeriría el TRUNCATE previo. En ese caso, bastaría con utilizar exclusivamente el SqlBulkCopy 
        // para realizar una inserción masiva aditiva (append) de forma directa.
        // =========================================================================================

        using var conn = new SqlConnection(connString);
        conn.Open();

        Console.WriteLine($"[Info] Vaciando tabla 'Movimientos' para procesar el archivo acumulativo: {filePath}...");
        using (var cmdClear = new SqlCommand("TRUNCATE TABLE Movimientos", conn))
        {
            cmdClear.ExecuteNonQuery();
        }

        DataTable dt = new();
        dt.Columns.Add("CodCliente", typeof(string));
        dt.Columns.Add("Estanteria", typeof(string));
        dt.Columns.Add("Cd", typeof(int));
        dt.Columns.Add("Fecha", typeof(DateTime));

        using var bulkCopy = new SqlBulkCopy(conn);
        bulkCopy.DestinationTableName = "Movimientos";

        // Mapeo explícito entre las columnas de la memoria y la base de datos
        bulkCopy.ColumnMappings.Add("CodCliente", "CodCliente");
        bulkCopy.ColumnMappings.Add("Estanteria", "Estanteria");
        bulkCopy.ColumnMappings.Add("Cd", "Cd");
        bulkCopy.ColumnMappings.Add("Fecha", "Fecha");

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error crítico: No se encuentra el archivo Excel en la ruta especificada: '{filePath}'.");
            return;
        }

        Console.WriteLine("Leyendo archivo Excel y exportando movimientos a la base de datos...");

        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        // Omitir cabeceras de la fila 1
        reader.Read(); 

        int rowCount = 0;
        while (reader.Read())
        {
            // Posiciones en el Excel: CodCliente(0), Estanteria(4), Cd(5), Fecha(6)
            string codCliente = reader.GetValue(0)?.ToString() ?? string.Empty;
            string estanteria = reader.GetValue(4)?.ToString() ?? string.Empty;
            
            if (string.IsNullOrEmpty(codCliente) || string.IsNullOrEmpty(estanteria)) 
                continue;

            int cd = Convert.ToInt32(reader.GetValue(5));
            
            // Tratamiento robusto para la fecha
            DateTime fecha;
            var fechaCelda = reader.GetValue(6);
            if (fechaCelda is DateTime dtExcel) 
            {
                fecha = dtExcel;
            } 
            else 
            {
                string fechaStr = fechaCelda?.ToString() ?? string.Empty;
                
                if (string.IsNullOrWhiteSpace(fechaStr))
                    continue;

                fecha = DateTime.ParseExact(fechaStr.Trim(), "M/d/yyyy", CultureInfo.InvariantCulture);
            }

            dt.Rows.Add(codCliente, estanteria, cd, fecha);
            rowCount++;

            // Batching: Inserción por bloques de 50.000 para optimizar memoria RAM
            if (rowCount % 50000 == 0)
            {
                bulkCopy.WriteToServer(dt);
                dt.Clear(); 
                Console.WriteLine($"{rowCount} registros cargados...");
            }
        }
        
        if (dt.Rows.Count > 0)
        {
            bulkCopy.WriteToServer(dt);
        }
        
        Console.WriteLine($"Importación masiva completada con éxito. Total registros insertados: {rowCount}");
    }

    private static void CalcularFacturacion(string connString, string codCliente, int year, int month) 
    {
        DateTime fechaInicio = new DateTime(year, month, 1);
        int daysInMonth = DateTime.DaysInMonth(year, month);
        DateTime fechaFin = new DateTime(year, month, daysInMonth);

        using var conn = new SqlConnection(connString);
        conn.Open();

        string sqlCliente = "SELECT Nombre, Tarifa FROM Clientes WHERE CodCliente = @Cliente";
        using var cmdCliente = new SqlCommand(sqlCliente, conn);
        cmdCliente.Parameters.AddWithValue("@Cliente", codCliente);

        string nombreCliente = "";
        decimal tarifa = 0m;

        using (var rdr = cmdCliente.ExecuteReader())
        {
            if (rdr.Read())
            {
                nombreCliente = rdr.GetString(0);
                tarifa = rdr.GetDecimal(1);
            }
            else
            {
                Console.WriteLine($"Error: El cliente '{codCliente}' no existe en la tabla Maestra de Clientes.");
                return;
            }
        }

        Console.WriteLine($"\nFACTURACIÓN PARA: {nombreCliente} ({codCliente}) - TARIFA: {tarifa:C2}/día");
        Console.WriteLine(new string('=', 60));

        // 2. Saldo Inicial
        string sqlSaldos = @"
            SELECT Estanteria, SUM(Cd) as Saldo
            FROM Movimientos
            WHERE CodCliente = @Cliente AND CAST(Fecha AS DATE) < @FechaInicio
            GROUP BY Estanteria
            HAVING SUM(Cd) > 0";

        using var cmdSaldos = new SqlCommand(sqlSaldos, conn);
        cmdSaldos.Parameters.AddWithValue("@Cliente", codCliente); 
        cmdSaldos.Parameters.AddWithValue("@FechaInicio", fechaInicio);
        
        var saldoEstanterias = new Dictionary<string, int>(); 
        using (var reader = cmdSaldos.ExecuteReader()) 
        {
            while (reader.Read()) 
            {
                saldoEstanterias[reader.GetString(0)] = reader.GetInt32(1);
            }
        }

        // 3. Movimientos del mes
        string sqlMes = @"
            SELECT CAST(Fecha AS DATE) as FechaDia, Estanteria, SUM(Cd) as Cambio
            FROM Movimientos
            WHERE CodCliente = @Cliente AND CAST(Fecha AS DATE) >= @FechaInicio AND CAST(Fecha AS DATE) <= @FechaFin
            GROUP BY CAST(Fecha AS DATE), Estanteria";

        using var cmdMes = new SqlCommand(sqlMes, conn);
        cmdMes.Parameters.AddWithValue("@Cliente", codCliente);
        cmdMes.Parameters.AddWithValue("@FechaInicio", fechaInicio);
        cmdMes.Parameters.AddWithValue("@FechaFin", fechaFin);

        var movimientosMes = new Dictionary<DateTime, List<(string Estanteria, int Cambio)>>();
        using (var reader = cmdMes.ExecuteReader()) 
        {
            while (reader.Read()) 
            {
                DateTime fecha = reader.GetDateTime(0);
                if (!movimientosMes.ContainsKey(fecha)) 
                    movimientosMes[fecha] = new List<(string, int)>();
                
                movimientosMes[fecha].Add((reader.GetString(1), reader.GetInt32(2)));
            }
        }

        // 4. Delegar al método puro de cálculo
        decimal totalMes = CalcularTotalMesEnMemoria(year, month, tarifa, saldoEstanterias, movimientosMes, (dia, estanterias, t, subtotal) => {
            Console.WriteLine($"{dia:dd/MM/yyyy}: {estanterias,3} estanterías ocupadas x {t:C2} = {subtotal:C2}");
        });

        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"TOTAL FACTURACIÓN DEL MES: {totalMes:C2}\n");
    }

    /// <summary>
    /// Método puramente en memoria que calcula la facturación mensual. 
    /// Ideal para pruebas unitarias sin dependencias de base de datos.
    /// </summary>
    public static decimal CalcularTotalMesEnMemoria(
        int year, 
        int month, 
        decimal tarifa, 
        Dictionary<string, int> saldoEstanterias, 
        Dictionary<DateTime, List<(string Estanteria, int Cambio)>> movimientosMes,
        Action<DateTime, int, decimal, decimal>? onDayCalculated = null)
    {
        int daysInMonth = DateTime.DaysInMonth(year, month);
        decimal totalMes = 0;

        for (int i = 1; i <= daysInMonth; i++) 
        {
            DateTime diaActual = new DateTime(year, month, i);
            
            if (movimientosMes.ContainsKey(diaActual)) 
            {
                foreach (var mov in movimientosMes[diaActual]) 
                {
                    if (!saldoEstanterias.ContainsKey(mov.Estanteria)) 
                        saldoEstanterias[mov.Estanteria] = 0;
                
                    saldoEstanterias[mov.Estanteria] += mov.Cambio;
                
                    if (saldoEstanterias[mov.Estanteria] <= 0) 
                        saldoEstanterias.Remove(mov.Estanteria);
                }
            }

            int estanteriasOcupadas = saldoEstanterias.Count;
            decimal facturacionDia = estanteriasOcupadas * tarifa;
            totalMes += facturacionDia;

            onDayCalculated?.Invoke(diaActual, estanteriasOcupadas, tarifa, facturacionDia);
        }

        return totalMes;
    }
}