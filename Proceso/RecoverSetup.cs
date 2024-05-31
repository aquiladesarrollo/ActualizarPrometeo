using System.Diagnostics;
using StartRobot;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace RecoverSetup
{
    class ClientesSetup
    {
        private static readonly StartBot startBot = new StartBot();

        public static void RecuperarArchivos(string cliente, out string[,] bonds, out string[,] cashFlow, out string[,] saldos, out string[,] valoresMercado, out string[,] VMsaldos)
        {
            string pathEntrada = startBot.cfgDic["rutaOrigen"];
            string path = Path.Combine(pathEntrada, cliente);
            string pathBonds = Path.Combine(path, "bonds.txt");
            string pathCashFlow = Path.Combine(path, "cashflow.txt");
            string pathSaldos = Path.Combine(path, "saldos.txt");
            string pathValoresMercado = Path.Combine(path, "valoresMercado.txt");
            string pathVMsaldos = Path.Combine(path, "VMsaldos.txt");

            bonds = GenerarMatriz(pathBonds);
            cashFlow = GenerarMatriz(pathCashFlow);
            saldos = GenerarMatriz(pathSaldos);
            valoresMercado = GenerarMatriz(pathValoresMercado);
            VMsaldos = GenerarMatriz(pathVMsaldos);
            GenerarExcelConsolidado(cliente, bonds, cashFlow, valoresMercado, VMsaldos);

            Log("Se recuperó la información para el setup del cliente: " + cliente);
        }



        private static string[,] GenerarMatriz(string path)
        {
            string[,]? matriz = null;
            if (File.Exists(path))
            {
                string[] lines = File.ReadAllLines(path);
                int filas = lines.Length;
                int columnas = lines[0].Split('|').Length;

                matriz = new string[filas, columnas];

                for (int i = 0; i < filas; i++)
                {
                    string[] elementos = lines[i].Split('|');
                    for (int j = 0; j < columnas; j++)
                    {
                        matriz[i, j] = elementos[j];
                    }
                }
            }

            return matriz;
        }
        public static void GenerarExcelConsolidado(string cliente, string[,] bonds, string[,] cashflow, string[,] valoresMercado, string[,] VMsaldos)
        {
            if (bonds == null)
            {
                return;
            }

            string rutaDeposito = startBot.cfgDic["rutaDataFeeder"];
            string fecha = DateTime.Now.ToString("dd-MM-yyyy");

            string pathDeposito = Path.Combine(rutaDeposito, fecha);

            if (!Directory.Exists(pathDeposito))
            {
                Directory.CreateDirectory(pathDeposito);
            }

            pathDeposito = Path.Combine(pathDeposito, cliente);

            if (!Directory.Exists(pathDeposito))
            {
                Directory.CreateDirectory(pathDeposito);
            }

            pathDeposito = pathDeposito + $"\\Consolidado_{cliente}_{fecha}.xlsx";
            if (File.Exists(pathDeposito))
            {
                return;
            }

            // Crear una instancia de Excel
            Excel.Application excelApp = new Excel.Application();

            // Crear un nuevo libro de Excel
            Excel.Workbook workbook = excelApp.Workbooks.Add();

            try
            {
                // Llenar la segunda hoja con la matriz cashFlow
                LlenarHoja(workbook, cashflow, "Cash Flow");

                // Llenar la primera hoja con la matriz bonds
                LlenarHoja(workbook, bonds, "Bonds");

                LlenarHoja(workbook, VMsaldos, "VM para Setup");

                // Llenar la cuarta hoja con la matriz valoresMercado
                LlenarHoja(workbook, valoresMercado, "Valores de Mercado");


                workbook.SaveAs(pathDeposito);

                Log("Se realizó el consolidado para el cliente: " + cliente);
            }
            catch (Exception ex)
            {
                Log("Error obteniendo el consolidado para el cliente: " + cliente);
            }
            finally
            {
                // Cerrar y liberar recursos
                workbook.Close();
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }
        private static void LlenarHoja(Excel.Workbook workbook, string[,] data, string sheetName)
        {
            // Agregar una nueva hoja al libro
            Excel.Worksheet worksheet = workbook.Sheets.Add();
            worksheet.Name = sheetName;

            // Obtener el rango de celdas de la hoja
            Excel.Range range = worksheet.Range["A1"].Resize[data.GetLength(0), data.GetLength(1)];

            // Crear una nueva matriz para almacenar los datos convertidos
            object[,] convertedData = new object[data.GetLength(0), data.GetLength(1)];

            // Convertir los datos a valores numéricos si es posible
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    if (double.TryParse(data[i, j], out double numericValue))
                    {
                        convertedData[i, j] = numericValue;
                    }
                    else
                    {
                        convertedData[i, j] = data[i, j];
                    }
                }
            }

            // Asignar los datos convertidos al rango de celdas utilizando Value2
            range.Value2 = convertedData;
        }

        private static void Log(string message)
        {
            DateTime miFechaActual = DateTime.Now;
            var st = new StackTrace();
            var sf = st.GetFrame(1);
            var currentMethodName = sf?.GetMethod()?.Name;

            string? metodo = currentMethodName;

            if (!Directory.Exists(startBot.cfgDic["rutaLogs"]))
            {
                Directory.CreateDirectory(startBot.cfgDic["rutaLogs"]);
            }

            string pathLog = startBot.cfgDic["rutaLogs"] + "\\" + miFechaActual.ToString("yyyyMMdd") + "_ActualizarPrometeoLog.txt";
            string className = nameof(ClientesSetup);

            using (System.IO.StreamWriter escritor = new System.IO.StreamWriter(pathLog, true))
            {
                escritor.WriteLine(">" + miFechaActual.ToString() + " > " + className + " > " + metodo + " > " + message);
            }
        }
    }
}