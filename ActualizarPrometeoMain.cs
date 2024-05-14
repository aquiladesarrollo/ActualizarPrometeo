using System.Diagnostics;
using StartRobot;
using RecoverSetup;
using ExcelFill;
using Excel = Microsoft.Office.Interop.Excel;

class SetupClienteMain
{
    private static readonly StartBot startBot = new StartBot();
    private static readonly ClientesSetup recoverSetup = new ClientesSetup();
    private static readonly ExcelFiller excelFill = new ExcelFiller();

    public static void Main(string[] args)
    {
        startBot.IniciarBot();
        Dictionary<string, string> cfgDic = startBot.cfgDic;
        Tuple<int, string> estatus = startBot.RobotProductiveServerStatus(cfgDic);

        if (estatus.Item1 != 1)
        {
            return;
        }

        ObtenerCuentas(cfgDic, out string[] setupClientes, out string[] tipoPersona, out string[] anioSetup);
        CerrarExcel();

        for (int i = 0; i < setupClientes.Length; i++)
        {
            try
            {
                ClientesSetup.RecuperarArchivos(setupClientes[i], out string[,] bonds, out string[,] cashflow, out string[,] saldos, out string[,] valoresMercado, out string[,] VMsaldos);
                excelFill.GenerarExcel(setupClientes[i], tipoPersona[i], anioSetup[i], bonds, cashflow, saldos, valoresMercado, VMsaldos);
                RegistrarProcesados(setupClientes[i]);
            }
            catch (Exception ex)
            {
                RegistrarNoProcesados(setupClientes[i], ex.Message);
                Log(ex.Message);
            }
            finally
            {
                CerrarExcel();
            }
        }

        LimpiarTemp();
        FinEjecución();
    }

    private static void ObtenerCuentas(Dictionary<string, string> cfgDic, out string[] clientesSetup, out string[] tipoPersona, out string[] anioSetup)
    {
        string clientesPath = cfgDic["diccionarioSetup"];

        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(clientesPath);
        Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets["Setup"];

        Excel.Range columnRange = excelWorksheet.UsedRange;

        // Contar las celdas utilizadas en la columna
        int longitudColumna = columnRange.Rows.Count;

        // Ajustar el tamaño de los arrays
        clientesSetup = new string[longitudColumna - 1];
        tipoPersona = new string[longitudColumna - 1];
        anioSetup = new string[longitudColumna - 1];

        for (int i = 2; i <= longitudColumna; i++)
        {
            clientesSetup[i - 2] = excelWorksheet.Cells[i, 1].Value?.ToString();
            tipoPersona[i - 2] = excelWorksheet.Cells[i, 2].Value?.ToString();
            anioSetup[i - 2] = excelWorksheet.Cells[i, 3].Value?.ToString();
        }

        excelWorkbook.Close();
        excelApp.Quit();
    }

    private static void LimpiarTemp()
    {
        string temp = startBot.cfgDic["tempFolder"];
        string[] files = Directory.GetFiles(temp);

        foreach (string file in files)
        {
            File.Delete(file);
        }

        Log("Se eliminaron los archivos temporales");
    }

    private static void FinEjecución()
    {
        string mensaje = $"" +
            $"\r\n _____  ____  ____       ___      ___        ___  ____    ___    __  __ __    __  ____  ___   ____  \r\n|     ||    ||    \\     |   \\    /  _]      /  _]|    |  /  _]  /  ]|  |  |  /  ]|    |/   \\ |    \\ \r\n|   __| |  | |  _  |    |    \\  /  [_      /  [_ |__  | /  [_  /  / |  |  | /  /  |  ||     ||  _  |\r\n|  |_   |  | |  |  |    |  D  ||    _]    |    _]__|  ||    _]/  /  |  |  |/  /   |  ||  O  ||  |  |\r\n|   _]  |  | |  |  |    |     ||   [_     |   [_/  |  ||   [_/   \\_ |  :  /   \\_  |  ||     ||  |  |\r\n|  |    |  | |  |  |    |     ||     |    |     \\  `  ||     \\     ||     \\     | |  ||     ||  |  |\r\n|__|   |____||__|__|    |_____||_____|    |_____|\\____||_____|\\____| \\__,_|\\____||____|\\___/ |__|__|\r\n                                                                                                    \r\n";
        Log(mensaje);
    }

    private static void CerrarExcel()
    {
        Process[] processes = Process.GetProcessesByName("Excel");
        foreach (Process process in processes)
        {
            process.Kill();
            process.WaitForExit();
        }
    }

    private static void RegistrarProcesados(string cliente)
    {
        DateTime miFechaActual = DateTime.Now;

        if (!Directory.Exists(startBot.cfgDic["procesados"]))
        {
            Directory.CreateDirectory(startBot.cfgDic["procesados"]);
        }

        string pathLog = startBot.cfgDic["procesados"] + "\\" + "Procesados_" + miFechaActual.ToString("ddMMyyyy") + ".txt";

        using (System.IO.StreamWriter escritor = new System.IO.StreamWriter(pathLog, true))
        {
            escritor.WriteLine(cliente + "|" + miFechaActual.ToString("hh:mm:ss dd/MM/yyyy") + "|Procesado");
        }
    }

    private static void RegistrarNoProcesados(string cliente, string ex)
    {
        DateTime miFechaActual = DateTime.Now;

        if (!Directory.Exists(startBot.cfgDic["noProcesados"]))
        {
            Directory.CreateDirectory(startBot.cfgDic["noProcesados"]);
        }

        string pathLog = startBot.cfgDic["noProcesados"] + "\\" + "NoProcesados_" + miFechaActual.ToString("ddMMyyyy") + ".txt";

        using (System.IO.StreamWriter escritor = new System.IO.StreamWriter(pathLog, true))
        {
            escritor.WriteLine(cliente + "|" + miFechaActual.ToString("hh:mm:ss dd/MM/yyyy") + "|No Procesado|" + ex);
        }
    }

    private static void EliminarTemps()
    {
        string tempPath = startBot.cfgDic["tempFolder"];
        string[] archivos = Directory.GetFiles(tempPath);

        foreach (string archivo in archivos)
        {
            File.Delete(archivo);
        }
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
        string className = nameof(SetupClienteMain);

        using (System.IO.StreamWriter escritor = new System.IO.StreamWriter(pathLog, true))
        {
            escritor.WriteLine(">" + miFechaActual.ToString() + " > " + className + " > " + metodo + " > " + message);
        }
    }
}