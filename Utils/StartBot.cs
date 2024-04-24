using System.Xml;
using System.Diagnostics;

namespace StartRobot
{
    public class StartBot
    {
        private static string config = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Config", "configActualizarPrometeo.xml");
        private static Dictionary<string, string> _cfgDic = GetConfigDic();

        public Dictionary<string, string> cfgDic
        {
            get { return _cfgDic; }
            private set { _cfgDic = value; }
        }

        public void IniciarBot()
        {
            string message = @"
                                 _   _                         ______ _                      _   __        
                                | \ | |                       |  ____(_)                    (_) /_/        
                                |  \| |_   _  _____   ____ _  | |__   _  ___  ___ _   _  ___ _  ___  _ __  
                                | . ` | | | |/ _ \ \ / / _` | |  __| | |/ _ \/ __| | | |/ __| |/ _ \| '_ \ 
                                | |\  | |_| |  __/\ V / (_| | | |____| |  __/ (__| |_| | (__| | (_) | | | |
                                |_| \_|\__,_|\___| \_/ \__,_| |______| |\___|\___|\__,_|\___|_|\___/|_| |_|
                                                                    _/ |                                   
                                                                    |__/                                    
                                ";
            Log(message);

            if (cfgDic == null)
            {
                Log("Error obteniendo el archivo de configuración");
            }
            else
            {
                Log("Se obtuvo el archivo de configuración");
            }
        }

        private static Dictionary<string, string> GetConfigDic()
        {
            string duplicateVariables = String.Empty;
            // Se instancia el diccionario de configuración vacío
            Dictionary<string, string> cfgDic = new Dictionary<string, string>();
            string key = String.Empty;
            string value = String.Empty;
            try
            {
                // Lectura de configuración XML 
                XmlDocument document = new XmlDocument();
                string pathCfgFile = config;
                document.Load(pathCfgFile);
                // Agregar nodos al diccionario
                XmlNodeList addNodes = document.GetElementsByTagName("add");
                foreach (XmlElement nodo in addNodes)
                {
                    key = nodo.GetAttribute("key").Trim();
                    value = nodo.GetAttribute("value").Trim();

                    if (!cfgDic.ContainsKey(key))
                    {
                        cfgDic.Add(key, value);
                    }
                    else
                    {
                        if (String.IsNullOrEmpty(duplicateVariables))
                        {
                            duplicateVariables = key;
                        }
                        else
                        {
                            duplicateVariables = duplicateVariables + " | " + key;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                return cfgDic = null;
            }

            return cfgDic;
        }

        public Tuple<int, string> RobotProductiveServerStatus(Dictionary<string, string> cfgDic)
        {
            int status = 0;
            string msgCode = "Success";
            // Cargar información del cfgDic
            string D_config = cfgDic["D_config"];

            // Inicio de Tareas
            if (status != -1)
            {
                string msgLog = "";
                XmlDocument configXml = new XmlDocument();
                try
                {
                    string pathCfgFile = $"{D_config}\\configActualizarPrometeo.xml";
                    configXml.Load(pathCfgFile);
                    string missingFiles = String.Empty;
                    string llave = String.Empty;
                    string valor = String.Empty;
                    //	Inicia la validacion de carpetas y archivos que esten dentro de 
                    //	cualquier etiqueta con atributo [validate="true"]
                    XmlNodeList listNodes = configXml.SelectNodes("//*[@validate='true']");
                    Log("Iniciando la validación de archivos y directorios");

                    foreach (XmlElement node in listNodes)
                    {
                        XmlNodeList childNodes = node.SelectNodes("add");
                        foreach (XmlElement kvp in childNodes)
                        {
                            llave = kvp.GetAttribute("key").Trim();
                            valor = cfgDic[llave];
                            if (File.Exists(valor) || Directory.Exists(valor))
                            {
                                Log($"OK | {valor}");
                            }
                            else
                            {
                                missingFiles = missingFiles + valor + "\n";
                                Log($" x | {valor}");
                            }
                        }

                        if (!String.IsNullOrEmpty(missingFiles))
                        {
                            msgLog = $"\n\nFaltan los siguientes archivos/directorios: \n{missingFiles}\n\n";
                            Log(msgLog);
                            msgCode = $"Err003*{msgLog}";
                        }
                        else
                        {
                            status = 1;
                            Log("Se encontraron todos los archivos/directorios");
                        }
                    }
                }
                catch (Exception e)
                {
                    System.Console.WriteLine(e);
                    msgCode = "Err003";
                    status = -1;
                    Log($"Al realizar la verificación de las carpetas del robot\n{e}");
                }
            }
            return new Tuple<int, string>(status, msgCode);
        }

        private static void Log(string message)
        {
            DateTime miFechaActual = DateTime.Now;
            var st = new StackTrace();
            var sf = st.GetFrame(1);
            var currentMethodName = sf.GetMethod().Name;

            string metodo = currentMethodName;

            if (!Directory.Exists(_cfgDic["rutaLogs"]))
            {
                Directory.CreateDirectory(_cfgDic["rutaLogs"]);
            }

            string pathLog = _cfgDic["rutaLogs"] + "\\" + miFechaActual.ToString("yyyyMMdd") + "_ActualizarPrometeoLog.txt";
            string className = nameof(StartBot);

            using (System.IO.StreamWriter escritor = new System.IO.StreamWriter(pathLog, true))
            {
                escritor.WriteLine(">" + miFechaActual.ToString() + " > " + className + " > " + metodo + " > " + message);
            }
        }
    }
}