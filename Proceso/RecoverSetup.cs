using System.Diagnostics;
using StartRobot;

namespace RecoverSetup
{
    class ClientesSetup
    {
        private static readonly StartBot startBot = new StartBot();

        public static void RecuperarArchivos(string cliente, out string[,] bonds, out string[,] cashFlow, out string[,] saldos, out string[,] valoresMercado, out string[,] valoresMercadoAntiguos)
        {
            string pathEntrada = startBot.cfgDic["rutaOrigen"];
            string path = Path.Combine(pathEntrada, cliente);
            string pathBonds = Path.Combine(path, "bonds.txt");
            string pathCashFlow = Path.Combine(path, "cashflow.txt");
            string pathSaldos = Path.Combine(path, "saldos.txt");
            string pathValoresMercado = Path.Combine(path, "valoresMercado.txt");
            string pathValoresMercadoAntiguos = Path.Combine(path, "valoresMercadoAntiguos.txt");

            bonds = GenerarMatriz(pathBonds);
            cashFlow = GenerarMatriz(pathCashFlow);
            saldos = GenerarMatriz(pathSaldos);
            valoresMercado = GenerarMatriz(pathValoresMercado);
            valoresMercadoAntiguos = GenerarMatriz(pathValoresMercadoAntiguos);
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

            string pathLog = startBot.cfgDic["rutaLogs"] + "\\" + miFechaActual.ToString("yyyyMMdd") + "_SetupClientesLog.txt";
            string className = nameof(ClientesSetup);

            using (System.IO.StreamWriter escritor = new System.IO.StreamWriter(pathLog, true))
            {
                escritor.WriteLine(">" + miFechaActual.ToString() + " > " + className + " > " + metodo + " > " + message);
            }
        }
    }
}