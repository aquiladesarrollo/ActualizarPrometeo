using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using StartRobot;
using System.Text;
using System.Globalization;
using System.Runtime.InteropServices;
using Newtonsoft.Json;

namespace ExcelFill
{
    class ExcelFiller
    {
        private static readonly StartBot startBot = new StartBot();

        public void GenerarExcel(string cliente, string tipoPersona, string anioSetup, string[,] bonds, string[,] cashflow, string[,] saldos, string[,] valoresMercado, string[,] valoresMercadoAntiguos)
        {
            // Verifica si existe alguna macro preexistente
            bool existeMacro = RevisiónExistenciaExcel(cliente, anioSetup);

            if (existeMacro)
            {
                BackupExcel(cliente, anioSetup);
            }
            else
            {
                CopiarPlantillaExcel(cliente, anioSetup, tipoPersona);
            }

            string hoy = DateTime.Today.ToString("dd-MM-yyyy");
            string path = startBot.cfgDic["rutaDeposito"] + $"\\{hoy}\\{cliente}\\MacroPatrimonial_{cliente}_{anioSetup}.xlsm";

            Excel.Application? excelApp = null;
            Excel.Workbook? excelWorkbook = null;
            Excel.Worksheet? excelWorksheet = null;

            if (bonds == null || cashflow == null || valoresMercado == null )
            {
                Log("El cliente " + cliente + " no tiene datos");
                return;
            }

            try
            {
                excelApp = new Excel.Application();
                //excelApp.Visible = true; 
                excelWorkbook = excelApp.Workbooks.Open(path);

                bool esPIC = false;
                if (tipoPersona == "PIC")
                {
                    esPIC = true;
                }

                string fechaSetup = $"01/01/{anioSetup}";
                string fechaMacro = string.Empty;

                if (existeMacro)
                {
                    fechaMacro = ObtenerUltimoMovimiento(excelWorkbook, esPIC);
                }

                AjusteBonds(bonds, esPIC, fechaSetup, fechaMacro, out string[,] newBonds);
                AjusteCashflow(cashflow, fechaSetup, fechaMacro, out string[,] newCashflow);
                ObtenerValorPortfolio(valoresMercado, out string valorPortafolio);

                ActualizarComprasYVentas(cliente, newBonds, excelWorkbook);
                ActualizarDivInt(cliente, newCashflow, excelWorkbook);
                ActualizarPortada(cliente, valorPortafolio, esPIC, fechaSetup, excelWorkbook);
                ActualizarBaseDatos(excelWorkbook);

                EjecutarMacros(cliente, excelApp, excelWorkbook);
                ActualizarValoresMercado(cliente, valoresMercado, valoresMercadoAntiguos, excelWorkbook, esPIC);
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
            finally
            {
                // Liberar la hoja de cálculo primero
                if (excelWorksheet != null)
                {
                    Marshal.ReleaseComObject(excelWorksheet);
                }

                // Cerrar el libro y liberar el objeto
                if (excelWorkbook != null)
                {
                    excelWorkbook.Close();
                    Marshal.ReleaseComObject(excelWorkbook);
                }

                // Cerrar Excel y liberar el objeto de la aplicación Excel
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
            }
        }

        private bool RevisiónExistenciaExcel(string cliente, string anioSetup)
        {
            string pathOrigen = startBot.cfgDic["rutaMacros"] + $"\\{cliente}\\{anioSetup}\\MacroPatrimonial_{cliente}_{anioSetup}.xlsm";
            if(File.Exists(pathOrigen))
            {
                return true;
            }

            return false;
        }

        private static void CopiarPlantillaExcel(string cliente, string anioSetup, string tipoPersona)
        {
            string hoy = DateTime.Today.ToString("dd-MM-yyyy");
            string pathOrigen = string.Empty;

            if (tipoPersona == "PF")
            {
                pathOrigen = startBot.cfgDic["macroPF"];
            }
            else if (tipoPersona == "PIC")
            {
                pathOrigen = startBot.cfgDic["macroPIC"];
            }

            string pathDeposito = startBot.cfgDic["rutaDeposito"];

            pathDeposito = Path.Combine(pathDeposito, hoy);

            if (!Directory.Exists(pathDeposito))
            {
                Directory.CreateDirectory(pathDeposito);
            }


            pathDeposito = Path.Combine(pathDeposito, cliente);

            if (!Directory.Exists(pathDeposito))
            {
                Directory.CreateDirectory(pathDeposito);
            }

            string nombreArchivo = $"MacroPatrimonial_{cliente}_{anioSetup}.xlsm";
            pathDeposito = Path.Combine(pathDeposito, nombreArchivo);
            
            File.Copy(pathOrigen, pathDeposito, true);

            Log("Se copió la plantilla MacroPatrimonial");
        }

        private void BackupExcel(string cliente, string anioSetup)
        {
            string hoy = DateTime.Today.ToString("dd-MM-yyyy");
            string pathOrigen = startBot.cfgDic["rutaMacros"] + $"\\{cliente}\\{anioSetup}\\MacroPatrimonial_{cliente}_{anioSetup}.xlsm";
            string pathDeposito = startBot.cfgDic["rutaDeposito"];
            //string pathDeposito = startBot.cfgDic["rutaPruebas"];
            string pathBackup = startBot.cfgDic["rutaBackup"];

            pathBackup = Path.Combine(pathBackup, hoy);
            pathDeposito = Path.Combine(pathDeposito, hoy);

            if (!Directory.Exists(pathBackup))
            {
                Directory.CreateDirectory(pathBackup);
            }

            if (!Directory.Exists(pathDeposito))
            {
                Directory.CreateDirectory(pathDeposito);
            }

            pathDeposito = Path.Combine(pathDeposito, cliente);

            if (!Directory.Exists(pathDeposito))
            {
                Directory.CreateDirectory(pathDeposito);
            }

            string nombreArchivo = $"MacroPatrimonial_{cliente}_{anioSetup}.xlsm";
            pathBackup = Path.Combine(pathBackup, nombreArchivo);
            pathDeposito = Path.Combine(pathDeposito, nombreArchivo);

            File.Copy(pathOrigen, pathBackup, true);
            File.Copy(pathOrigen, pathDeposito, true);

            Log("Se copió la plantilla MacroPatrimonial");
        }

        public static string ObtenerUltimoMovimiento(Excel.Workbook excelWorkbook, bool esPIC)
        {
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets["Portada"];
            string fechaMovimiento = string.Empty;

            if(!esPIC)
            {
                object valorCelda = excelWorksheet.Cells[5, 3].Value;
                if(valorCelda != null)
                {
                    fechaMovimiento = excelWorksheet.Cells[5, 3].Value.ToString("dd/MM/yyyy");
                }
            }
            else
            {
                object valorCelda = excelWorksheet.Cells[6, 3].Value;
                if (valorCelda != null)
                {
                    fechaMovimiento = excelWorksheet.Cells[6, 3].Value.ToString("dd/MM/yyyy");
                }
            }
            
            return fechaMovimiento;
        }

        private static void ActualizarComprasYVentas(string cliente, string[,] bonds, Excel.Workbook excelWorkbook)
        {
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets["Compras y Ventas"];

            Excel.Range columnRange = excelWorksheet.UsedRange;

            // Contar las celdas utilizadas en la columna
            int longitudColumna = columnRange.Rows.Count;

            int contador = 0;
            for (int i = 6; i < longitudColumna; i++)
            {
                string contenido = excelWorksheet.Cells[i, 3].Value;
                if (contenido == null)
                {
                    break;
                }
                contador++;
            }

            contador = contador + 6;

            for (int i = 0; i < bonds.GetLength(0); i++)
            {
                excelWorksheet.Cells[contador + i, 2] = bonds[i, 0]; // Fecha
                excelWorksheet.Cells[contador + i, 3] = bonds[i, 1]; // Cuenta 
                excelWorksheet.Cells[contador + i, 4] = bonds[i, 2]; // Nombre Instrumento
                excelWorksheet.Cells[contador + i, 5] = bonds[i, 3]; // ISIN
                excelWorksheet.Cells[contador + i, 6] = bonds[i, 4]; // Concepto
                excelWorksheet.Cells[contador + i, 7] = bonds[i, 11]; // Tipo Instrumento
                excelWorksheet.Cells[contador + i, 8] = bonds[i, 6]; // Moneda
                excelWorksheet.Cells[contador + i, 9] = bonds[i, 7]; // Unidades
                excelWorksheet.Cells[contador + i, 10] = bonds[i, 8]; // Precio total
            }

            excelWorkbook.Save();

            Log("Se llenó la hoja de Compras y Ventas");
        }

        private static void ActualizarDivInt(string cliente, string[,] cashflow, Excel.Workbook excelWorkbook)
        {
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets["Div-Int"];

            Excel.Range columnRange = excelWorksheet.UsedRange;

            // Contar las celdas utilizadas en la columna
            int longitudColumna = columnRange.Rows.Count;

            int contador = 0;
            for (int i = 6; i < longitudColumna; i++)
            {
                if (excelWorksheet.Cells[i, 3].Value == null)
                {
                    break;
                }
                contador++;
            }

            contador = contador + 6;

            for (int i = 0; i < cashflow.GetLength(0); i++)
            {
                excelWorksheet.Cells[contador + i, 2] = cashflow[i, 0].ToString(); // Fecha
                excelWorksheet.Cells[contador + i, 3] = cashflow[i, 1]; // Cuenta 
                excelWorksheet.Cells[contador + i, 4] = cashflow[i, 13]; // Concepto
                excelWorksheet.Cells[contador + i, 6] = cashflow[i, 8]; // Monto
                excelWorksheet.Cells[contador + i, 7] = cashflow[i, 4]; // Descripción
                excelWorksheet.Cells[contador + i, 5] = cashflow[i, 6]; // Moneda
            }

            excelWorkbook.Save();

            Log("Se llenó la hoja DivInt");
        }

        private static void ActualizarPortada(string cliente, string valorPortafolio, bool esPIC, string fechaSetup, Excel.Workbook excelWorkbook)
        {
            DateTime.TryParseExact(fechaSetup, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateSetup);
            string anio = string.Empty;
            DateTime fechaActualizacion;

            if (DateTime.Today.Year > dateSetup.Year)
            {
                anio = dateSetup.ToString("yyyy");
                fechaActualizacion = new DateTime(dateSetup.Year, 12, 31);
            }
            else
            {
                anio = DateTime.Today.ToString("yyyy");
                DateTime today = DateTime.Today;
                DayOfWeek currentDayOfWeek = today.DayOfWeek;

                // Calcular la cantidad de días para retroceder al martes pasado
                int daysToSubtract = (currentDayOfWeek - DayOfWeek.Tuesday + 7) % 7;

                // Obtener la fecha del martes pasado
                fechaActualizacion = today.AddDays(-daysToSubtract);
            }
            
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets["Portada"];

            bool esMx = false;

            if (esPIC)
            {
                string valorCelda = excelWorksheet.Cells[6, 2].Value.ToString();
                if (valorCelda != "Última fecha de act.")
                {
                    esMx = true;
                }
            }

            if (esPIC && !esMx)
            {
                excelWorksheet.Cells[1, 3] = cliente;
                Thread.Sleep(250);
                excelWorksheet.Cells[2, 3] = anio;
                Thread.Sleep(250);
                excelWorksheet.Cells[4, 3] = ObtenerTipoCambio().Result;
                Thread.Sleep(250);
                excelWorksheet.Cells[5, 3] = valorPortafolio;
                Thread.Sleep(250);
                excelWorksheet.Cells[6, 3] = fechaActualizacion.ToString("dd/MM/yyyy").Replace(".", "");

                excelWorkbook.Save();
            }
            else if (esMx)
            {
                excelWorksheet.Cells[1, 3] = cliente;
                Thread.Sleep(250);
                excelWorksheet.Cells[2, 3] = anio;
                Thread.Sleep(250);
                excelWorksheet.Cells[4, 3] = valorPortafolio;
                Thread.Sleep(250);
                excelWorksheet.Cells[5, 3] = fechaActualizacion.ToString("dd/MM/yyyy").Replace(".", "");

                excelWorkbook.Save();
            }
            else
            {
                excelWorksheet.Cells[1, 3] = cliente;
                Thread.Sleep(250);
                excelWorksheet.Cells[2, 3] = anio;
                Thread.Sleep(250);
                excelWorksheet.Cells[4, 3] = valorPortafolio;
                Thread.Sleep(250);
                excelWorksheet.Cells[5, 3] = fechaActualizacion.ToString("dd/MM/yyyy").Replace(".", "");
                Thread.Sleep(250);
                excelWorksheet.Cells[6, 3] = 1;

                excelWorkbook.Save();
            }

            Log("Se llenó la portada");
        }

        private static void EjecutarMacros(string cliente, Excel.Application excelApp, Excel.Workbook excelWorkbook)
        {
            string[] macros = new string[] { "Ajustar", "AjustarDiv" };

            try
            {
                foreach (string macro in macros)
                {
                    excelApp.Run(macro);

                    while (!excelApp.Ready)
                    {
                        System.Threading.Thread.Sleep(100);
                    }
                }

                excelWorkbook.Save();
                Log("Se ejecutaron los Macro");
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private static void ActualizarBaseDatos(Excel.Workbook excelWorkbook)
        {
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets["Base de Datos"];
            string[,] baseDatos = ObtenerMatrizBaseDatos();

            for (int i = 0; i < baseDatos.GetLength(0); i++)
            {
                if (DateTime.TryParse(baseDatos[i, 0], out DateTime fechaTipoCambio))
                {
                    baseDatos[i, 0] = fechaTipoCambio.ToString("MM/dd/yyyy");
                }

                if (DateTime.TryParse(baseDatos[i, 17], out DateTime mesINPC))
                {
                    baseDatos[i, 17] = mesINPC.ToString("MM/dd/yyyy");
                }
            }

            // Crear una nueva matriz para almacenar los datos convertidos
            object[,] convertedData = new object[baseDatos.GetLength(0), baseDatos.GetLength(1)];

            // Convertir los datos a valores numéricos si es posible
            for (int i = 0; i < baseDatos.GetLength(0); i++)
            {
                for (int j = 0; j < baseDatos.GetLength(1); j++)
                {
                    if (double.TryParse(baseDatos[i, j], out double numericValue))
                    {
                        convertedData[i, j] = numericValue;
                    }
                    else
                    {
                        convertedData[i, j] = baseDatos[i, j];
                    }
                }
            }

            // Llenar la tabla
            excelWorksheet.Range[excelWorksheet.Cells[9, 2], excelWorksheet.Cells[convertedData.GetLength(0) + 8, convertedData.GetLength(1) + 1]].Value = convertedData;

            // Dar formato a las fechas
            int lastRow = excelWorksheet.Cells[excelWorksheet.Rows.Count, "B"].End[Excel.XlDirection.xlUp].Row;
            Excel.Range range = excelWorksheet.Range["B9", $"B{lastRow}"];
            range.NumberFormat = "dd/MM/yyyy";

            // Dar formato a los números
            range = excelWorksheet.Range["C9", "Q9"].EntireColumn;
            range.NumberFormat = "#,##0.0000;(#,##0.0000);-";

            range = excelWorksheet.Range["S9"].EntireColumn;
            range.NumberFormat = "mmm-yy";

            excelWorkbook.Save();

            Log("Se actualizó la Base de Datos");
        }

        private static void ActualizarValoresMercado(string cliente, string[,] valoresMercado, string[,] valoresMercadoAntiguo, Excel.Workbook excelWorkbook, bool esPIC)
        {
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets["Inventario"];

            // Obtener el rango de la columna
            Excel.Range columnRange = excelWorksheet.UsedRange;

            // Contar las celdas utilizadas en la columna
            int longitudColumna = columnRange.Rows.Count;

            if (esPIC)
            {
                int contador = 0;
                for (int i = 7; i < longitudColumna; i++)
                {
                    object contenido = excelWorksheet.Cells[i, 3].Value;
                    if (contenido == null)
                    {
                        break;
                    }
                    contador++;
                }

                string[] isin = new string[contador];

                for (int i = 0; i < contador; i++)
                {
                    string contenido = excelWorksheet.Cells[i + 7, 3].Value.ToString();
                    isin[i] = contenido;
                }

                contador = 0;

                for (int i = 0; i < isin.Length; i++)
                {
                    if (!string.IsNullOrEmpty(isin[i]))
                    {
                        contador++;
                    }
                }

                string[] newIsin = new string[contador];
                string[] valores = new string[contador];
                contador = 0;

                for (int i = 0; i < isin.Length; i++)
                {
                    if (!string.IsNullOrEmpty(isin[i]))
                    {
                        newIsin[contador] = isin[i];
                        contador++;
                    }
                }

                for (int i = 0; i < newIsin.Length; i++)
                {
                    for (int j = 0; j < valoresMercado.GetLength(0); j++)
                    {
                        if (newIsin[i] == valoresMercado[j, 0])
                        {
                            valores[i] = valoresMercado[j, 5];
                        }
                    }
                }

                for (int i = 0; i < valores.Length; i++)
                {
                    excelWorksheet.Cells[i + 7, 10] = valores[i];
                    Thread.Sleep(50);
                }
            }
            else
            {
                int contador = 0;
                for (int i = 8; i < longitudColumna; i++)
                {
                    object contenido = excelWorksheet.Cells[i, 3].Value;
                    if (contenido == null)
                    {
                        break;
                    }
                    contador++;
                }

                string[] isin = new string[contador];

                for (int i = 0; i < contador; i++)
                {
                    string contenido = excelWorksheet.Cells[i + 8, 3].Value.ToString();
                    isin[i] = contenido;
                }

                contador = 0;

                for (int i = 0; i < isin.Length; i++)
                {
                    if (!string.IsNullOrEmpty(isin[i]))
                    {
                        contador++;
                    }
                }

                string[] newIsin = new string[contador];
                string[] valores = new string[contador];
                contador = 0;

                for (int i = 0; i < isin.Length; i++)
                {
                    if (!string.IsNullOrEmpty(isin[i]))
                    {
                        newIsin[contador] = isin[i];
                        contador++;
                    }
                }

                for (int i = 0; i < newIsin.Length; i++)
                {
                    for (int j = 0; j < valoresMercado.GetLength(0); j++)
                    {
                        if (newIsin[i] == valoresMercado[j, 0])
                        {
                            valores[i] = valoresMercado[j, 5];
                        }
                    }
                }

                for (int i = 0; i < valores.Length; i++)
                {
                    excelWorksheet.Cells[i + 8, 10] = valores[i];
                    Thread.Sleep(50);
                }
            }

            excelWorkbook.Save();
            Log("Se llenó la hoja de Valores de Mercado");
        }

        private static void AjusteBonds(string[,] bonds, bool esPIC, string fechaSetup, string fechaMacro, out string[,] newBonds)
        {
            newBonds = (string[,])bonds.Clone();
            string[,] matrizSIC = ObtenerMatrizSIC();

            if (!esPIC)
            {
                for (int i = 0; i < bonds.GetLength(0); i++)
                {
                    for (int j = 0; j < matrizSIC.GetLength(0); j++)
                    {
                        if (bonds[i, 3] == matrizSIC[j, 0])
                        {
                            bonds[i, 11] = "Acciones (SIC)";
                        }
                        else if (bonds[i, 11] == "Equity")
                        {
                            bonds[i, 11] = "Acciones (NO SIC)";
                        }
                        else if (bonds[i, 11] == "Fixed Income")
                        {
                            bonds[i, 11] = "Bonos";
                        }
                        else if (bonds[i, 11] == "Derivados")
                        {
                            bonds[i, 11] = "Derivados";
                        }
                    }
                }
            }
            else
            {
                for (int i = 0; i < bonds.GetLength(0); i++)
                {
                    for (int j = 0; j < matrizSIC.GetLength(0); j++)
                    {
                        if (bonds[i, 11] == "Equity")
                        {
                            bonds[i, 11] = "Acciones";
                        }
                        else if (bonds[i, 11] == "Fixed Income")
                        {
                            bonds[i, 11] = "Bonos";
                        }
                        else if (bonds[i, 11] == "Derivados")
                        {
                            bonds[i, 11] = "Derivados";
                        }
                    }
                }
            }

            DateTime.TryParseExact(fechaSetup, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fechaSetupActual);
            DateTime limiteFechaSuperior = new DateTime();

            if (fechaSetupActual.Year == DateTime.Today.Year)
            {
                limiteFechaSuperior = DateTime.Today;
            }
            else
            {
                limiteFechaSuperior = new DateTime(fechaSetupActual.Year, 12, 31);
            }

            int contador = 0;

            for (int i = 0; i < bonds.GetLength(0); i++)
            {
                string fechaComparar = bonds[i, 0];
                DateTime fecha;
                if (DateTime.TryParseExact(fechaComparar, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
                {
                    // Revisa todos los años anteriores al actual
                    if (fecha <= limiteFechaSuperior)
                    {
                        contador++;
                    }
                }
            }

            newBonds = new string[contador, bonds.GetLength(1)];
            contador = 0;

            for (int i = 0; i < bonds.GetLength(0); i++)
            {
                string fechaComparar = bonds[i, 0];
                DateTime fecha;
                if (DateTime.TryParseExact(fechaComparar, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
                {
                    // Revisa todos los años anteriores al actual
                    if (fecha <= limiteFechaSuperior)
                    {
                        for (int j = 0; j < bonds.GetLength(1); j++)
                        {
                            newBonds[contador, j] = bonds[i, j];
                        }
                        contador++;
                    }
                }
            }

            for (int i = 0; i < newBonds.GetLength(0); i++)
            {
                string fechaComparar = newBonds[i, 0];
                DateTime fecha;
                if (DateTime.TryParseExact(fechaComparar, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
                {
                    // Revisa todos los años anteriores al actual
                    if (fecha.Year < fechaSetupActual.Year)
                    {
                        newBonds[i, 1] = "Saldos Iniciales";
                    }
                }
            }

            if(!string.IsNullOrEmpty(fechaMacro))
            {
                DateTime.TryParseExact(fechaMacro, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fechaMacroActual);
                contador = 0;

                for (int i = 0; i < newBonds.GetLength(0); i++)
                {
                    string fechaComparar = newBonds[i, 0];
                    DateTime fecha;
                    if (DateTime.TryParseExact(fechaComparar, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
                    {
                        // Revisa todos los años anteriores al actual
                        if (fecha > fechaMacroActual)
                        {
                            contador++;
                        }
                    }
                }

                string[,] newBondsMacro = new string[contador, newBonds.GetLength(1)];
                contador = 0;

                for (int i = 0; i < newBonds.GetLength(0); i++)
                {
                    string fechaComparar = newBonds[i, 0];
                    DateTime fecha;
                    if (DateTime.TryParseExact(fechaComparar, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
                    {
                        // Revisa todos los años anteriores al actual
                        if (fecha > fechaMacroActual)
                        {
                            for (int j = 0; j < newBonds.GetLength(1); j++)
                            {
                                newBondsMacro[contador, j] = newBonds[i, j];
                            }
                            contador++;
                        }
                    }
                }

                newBonds = (string[,])newBondsMacro.Clone();
            }

            string[,] newBonds2 = (string[,])newBonds.Clone();

            // Se omiten estos instrumentos
            string[] omisionesInstrumento = startBot.cfgDic["omisionesInstrumentoBonds"].Split(",");
            contador = 0;

            for (int i = 0; i < newBonds2.GetLength(0); i++)
            {
                string instrumento = newBonds2[i, 2];
                if (!ContieneSubcadena(instrumento, omisionesInstrumento))
                {
                    contador++;
                }
            }

            newBonds = new string[contador, newBonds2.GetLength(1)];
            contador = 0;

            for (int i = 0; i < newBonds2.GetLength(0); i++)
            {
                string instrumento = newBonds2[i, 2];
                if (!ContieneSubcadena(instrumento, omisionesInstrumento))
                {
                    for (int j = 0; j < newBonds2.GetLength(1); j++)
                    {
                        newBonds[contador, j] = newBonds2[i, j];
                    }
                    contador++;
                }
            }

            for (int i = 0; i < newBonds.GetLength(0); i++)
            {
                if (newBonds[i, 4] == "Buy")
                {
                    newBonds[i, 4] = "Compra";
                }
                else if (newBonds[i, 4] == "Sell")
                {
                    newBonds[i, 4] = "Venta";
                }
            }

            for (int i = 0; i < newBonds.GetLength(0); i++)
            {
                string moneda = newBonds[i, 6];

                switch (moneda)
                {
                    case "MXN":
                        newBonds[i, 6] = "Mx";
                        break;
                    case "GBP":
                        newBonds[i, 6] = "Libra";
                        break;
                    case "":
                        newBonds[i, 6] = "USD";
                        break;
                    default:
                        break;
                }
            }

            for (int i = 0; i < newBonds.GetLength(0); i++)
            {
                double valor;
                double unidades;
                if (!string.IsNullOrEmpty(newBonds[i, 7]) && !string.IsNullOrEmpty(newBonds[i, 8]) && newBonds[i, 4] != "Split")
                {
                    valor = double.Parse(newBonds[i, 7]);
                    valor = Math.Abs(valor);
                    unidades = double.Parse(newBonds[i, 8]);
                    unidades = Math.Abs(unidades);
                    newBonds[i, 7] = valor.ToString();
                    newBonds[i, 8] = unidades.ToString();
                }
            }

            for (int i = 0; i < newBonds.GetLength(0); i++)
            {
                if (ContieneSubcadena(newBonds[i, 4], startBot.cfgDic["compra"].Split(",")))
                {
                    if (newBonds[i, 4] == "Cover Short")
                    {
                        newBonds[i, 11] = "Opciones Acciones";
                    }
                    else if (newBonds[i, 4] == "(Cancellation) Reinvestment")
                    {
                        double valor = double.Parse(newBonds[i, 7]);
                        valor = valor * (-1);
                        newBonds[i, 7] = valor.ToString();
                    }
                    newBonds[i, 4] = "Compra";
                }
                else if (ContieneSubcadena(newBonds[i, 4], startBot.cfgDic["venta"].Split(",")))
                {
                    if (newBonds[i, 4] == "Sell Short")
                    {
                        newBonds[i, 11] = "Opciones Acciones";
                    }
                    newBonds[i, 4] = "Venta";
                }
                else if (ContieneSubcadena(newBonds[i, 4], startBot.cfgDic["compra-"].Split(",")) && !string.IsNullOrEmpty(newBonds[i, 7]) && !string.IsNullOrEmpty(newBonds[i, 8]))
                {
                    newBonds[i, 4] = "Compra-";
                    double valor = double.Parse(newBonds[i, 7]) * (-1);
                    double unidades = double.Parse(newBonds[i, 8]) * (-1);
                    newBonds[i, 7] = valor.ToString();
                    newBonds[i, 8] = unidades.ToString();
                }
                else if (ContieneSubcadena(newBonds[i, 4], startBot.cfgDic["split"].Split(",")) && !string.IsNullOrEmpty(newBonds[i, 7]) && !string.IsNullOrEmpty(newBonds[i, 8]))
                {
                    double valor = double.Parse(newBonds[i, 7]);
                    double unidades = double.Parse(newBonds[i, 8]);
                    if (unidades > 0)
                    {
                        newBonds[i, 4] = "Compra";
                        newBonds[i, 7] = Math.Abs(valor).ToString();
                        newBonds[i, 8] = Math.Abs(unidades).ToString();
                    }
                    else
                    {
                        newBonds[i, 4] = "Compra-";
                        valor = valor * (-1);
                        newBonds[i, 7] = valor.ToString();
                        newBonds[i, 8] = unidades.ToString();

                    }
                }
            }

            Log("Se ajustó la matriz Bonds");
        }

        private static void AjusteCashflow(string[,] cashflow, string fechaSetup, string fechaMacro, out string[,] newCashflow)
        {
            DateTime.TryParseExact(fechaSetup, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fechaSetupActual);
            DateTime limiteFechaSuperior = new DateTime();

            if (fechaSetupActual.Year == DateTime.Today.Year)
            {
                limiteFechaSuperior = DateTime.Today;
            }
            else
            {
                limiteFechaSuperior = new DateTime(fechaSetupActual.Year, 12, 31);
            }

            int contador = 0;

            for (int i = 0; i < cashflow.GetLength(0); i++)
            {
                string fechaComparar = cashflow[i, 0];
                DateTime fecha;
                if (DateTime.TryParseExact(fechaComparar, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
                {
                    // Revisa todos los años anteriores al actual
                    if (fechaSetupActual <= fecha && fecha <= limiteFechaSuperior)
                    {
                        contador++;
                    }
                }
            }

            newCashflow = new string[contador, cashflow.GetLength(1)];
            contador = 0;

            for (int i = 0; i < cashflow.GetLength(0); i++)
            {
                string fechaComparar = cashflow[i, 0];
                DateTime fecha;
                if (DateTime.TryParseExact(fechaComparar, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
                {
                    // Revisa todos los años anteriores al actual
                    if (fechaSetupActual <= fecha && fecha <= limiteFechaSuperior)
                    {
                        for (int j = 0; j < cashflow.GetLength(1); j++)
                        {
                            newCashflow[contador, j] = cashflow[i, j];
                        }
                        contador++;
                    }
                }
            }

            if(!string.IsNullOrEmpty(fechaMacro))
            {
                DateTime.TryParseExact(fechaMacro, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime fechaMacroActual);
                contador = 0;

                for (int i = 0; i < newCashflow.GetLength(0); i++)
                {
                    string fechaComparar = newCashflow[i, 0];
                    DateTime fecha;
                    if (DateTime.TryParseExact(fechaComparar, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
                    {
                        // Revisa todos los años anteriores al actual
                        if (fecha > fechaMacroActual)
                        {
                            contador++;
                        }
                    }
                }

                string[,] newCashflowMacro = new string[contador, newCashflow.GetLength(1)];
                contador = 0;

                for (int i = 0; i < newCashflow.GetLength(0); i++)
                {
                    string fechaComparar = newCashflow[i, 0];
                    DateTime fecha;
                    if (DateTime.TryParseExact(fechaComparar, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out fecha))
                    {
                        // Revisa todos los años anteriores al actual
                        if (fecha > fechaMacroActual)
                        {
                            for (int j = 0; j < newCashflow.GetLength(1); j++)
                            {
                                newCashflowMacro[contador, j] = newCashflow[i, j];
                            }
                            contador++;
                        }
                    }
                }

                newCashflow = (string[,])newCashflowMacro.Clone();
            }

            string[,] cashflowType = new string[newCashflow.GetLength(0), newCashflow.GetLength(1) + 1];

            for (int i = 0; i < newCashflow.GetLength(0); i++)
            {
                for (int j = 0; j < newCashflow.GetLength(1); j++)
                {
                    cashflowType[i, j] = newCashflow[i, j];
                }
            }

            newCashflow = (string[,])cashflowType.Clone();

            for (int i = 0; i < newCashflow.GetLength(0); i++)
            {
                if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["dividendos"].Split(",")))
                {
                    newCashflow[i, 13] = "Dividendos";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["otrosIngresos"].Split(",")))
                {
                    newCashflow[i, 13] = "Otros Ingresos";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["impuestos"].Split(",")))
                {
                    newCashflow[i, 13] = "Impt. Ret.";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["comisiones"].Split(",")))
                {
                    newCashflow[i, 13] = "Comisiones";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["gastos"].Split(",")))
                {
                    newCashflow[i, 13] = "Gastos";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["aportacionCapital"].Split(",")))
                {
                    newCashflow[i, 13] = "Aportación de Capital";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["retiroCapital"].Split(",")))
                {
                    newCashflow[i, 13] = "Retiro de Capital";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["prestamosRecibidos"].Split(",")))
                {
                    newCashflow[i, 13] = "Préstamos Recibidos";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["pagoPrestamos"].Split(",")))
                {
                    newCashflow[i, 13] = "Pago de Préstamos";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["interesesGasto"].Split(",")))
                {
                    newCashflow[i, 13] = "Intereses (Gasto)";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["intereses"].Split(",")))
                {
                    newCashflow[i, 13] = "Intereses (Ingreso)";
                }
                else if (ContieneSubcadena(newCashflow[i, 4], startBot.cfgDic["noDeducibles"].Split(",")))
                {
                    newCashflow[i, 13] = "No Deducibles";
                }
                else
                {
                    newCashflow[i, 13] = "****";
                }
            }

            for (int i = 0; i < newCashflow.GetLength(0); i++)
            {
                if (!string.IsNullOrEmpty(newCashflow[i, 8]))
                {
                    double valor = double.Parse(newCashflow[i, 8]);
                    valor = Math.Abs(valor);

                    if (newCashflow[i, 4].IndexOf("Cancellation", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        valor = valor * (-1);
                    }

                    newCashflow[i, 8] = valor.ToString();
                }
            }

            Log("Se ajustó la matriz Cashflow");
        }

        private static void ObtenerValorPortfolio(string[,] valoresMercado, out string valorPortafolio)
        {
            double portfolio = 0;

            for (int i = 1; i < valoresMercado.GetLength(0); i++)
            {
                double valor = double.Parse(valoresMercado[i, 5]);
                valor = Math.Round(valor, 3);
                portfolio += valor;
            }

            valorPortafolio = portfolio.ToString();
        }

        private static bool ContieneSubcadena(string texto, string[] subcadenas)
        {
            foreach (string subcadena in subcadenas)
            {
                if (texto.IndexOf(subcadena, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            return false;
        }

        private static string[,] ObtenerMatrizSIC()
        {
            string tempPath = startBot.cfgDic["tempFolder"];
            string pathBivaSIC = tempPath + "\\" + "bivaSIC.txt";

            if (!File.Exists(pathBivaSIC))
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(startBot.cfgDic["bivaSic"]);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets["SIC"];
                Excel.Range excelRange = excelWorksheet.UsedRange;

                int rowCount = excelRange.Rows.Count;
                int colsCount = excelRange.Columns.Count;

                string[,] matrizSic = new string[rowCount, colsCount];

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colsCount; j++)
                    {
                        Excel.Range xlTmp = (Excel.Range)excelRange.Cells[i, j];
                        object vacio = xlTmp.Value;

                        if (vacio != null)
                        {
                            matrizSic[i - 1, j - 1] = xlTmp.Value.ToString();
                        }
                    }
                }

                StringBuilder txtContent = new StringBuilder();

                int rows = matrizSic.GetLength(0);
                int cols = matrizSic.GetLength(1);

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        //Agrega el valor actual y una coma(excepto para el último valor de la fila)
                        txtContent.Append(matrizSic[i, j]);
                        if (j < cols - 1)
                            txtContent.Append("|");
                    }

                    //Agrega un salto de línea después de cada fila(excepto para la última fila)
                    if (i < rows - 1)
                        txtContent.AppendLine();
                }

                //Guarda el contenido en el archivo
                string directoryPath = tempPath + "\\" + "bivaSIC.txt";
                File.WriteAllText(directoryPath, txtContent.ToString());
                Log("Se leyó el archivo BivaSIC");
                return matrizSic;
            }
            else
            {
                string[,]? matrizSic = null;
                string[] lines = File.ReadAllLines(pathBivaSIC);
                int filas = lines.Length;
                int columnas = lines[0].Split('|').Length;

                matrizSic = new string[filas, columnas];

                for (int i = 0; i < filas; i++)
                {
                    string[] elementos = lines[i].Split('|');
                    for (int j = 0; j < columnas; j++)
                    {
                        matrizSic[i, j] = elementos[j];
                    }
                }

                Log("Se leyó el archivo BivaSIC");
                return matrizSic;
            }
        }

        private static string[,] ObtenerMatrizBaseDatos()
        {
            string tempPath = startBot.cfgDic["tempFolder"];
            string pathBaseDatos = tempPath + "\\" + "baseDatos.txt";

            if (!File.Exists(pathBaseDatos))
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(startBot.cfgDic["baseDeDatos"]);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets["Base de Datos"];
                Excel.Range excelRange = excelWorksheet.UsedRange;

                int rowCount = excelRange.Rows.Count;
                int colsCount = 20;

                string[,] matrizBd = new string[rowCount, 19];

                for (int i = 8; i <= rowCount; i++)
                {
                    for (int j = 2; j <= colsCount; j++)
                    {
                        Excel.Range xlTmp = (Excel.Range)excelRange.Cells[i, j];
                        object vacio = xlTmp.Value;

                        if (vacio != null)
                        {
                            matrizBd[i - 8, j - 2] = xlTmp.Value.ToString();
                        }
                    }
                }

                excelWorkbook.Close();
                excelApp.Quit();

                int contador = 0;
                for (int i = 0; i < matrizBd.GetLength(0); i++)
                {
                    if (!string.IsNullOrEmpty(matrizBd[i, 0]))
                    {
                        contador++;
                    }
                }

                string[,] newMatrizBd = (string[,])matrizBd.Clone();
                matrizBd = new string[contador, newMatrizBd.GetLength(1)];
                contador = 0;

                for (int i = 0; i < matrizBd.GetLength(0); i++)
                {
                    if (!string.IsNullOrEmpty(newMatrizBd[i, 0]))
                    {
                        for (int j = 0; j < matrizBd.GetLength(1); j++)
                        {
                            matrizBd[contador, j] = newMatrizBd[i, j];
                        }
                        contador++;
                    }
                }

                StringBuilder txtContent = new StringBuilder();

                int rows = matrizBd.GetLength(0);
                int cols = matrizBd.GetLength(1);

                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        //Agrega el valor actual y una coma(excepto para el último valor de la fila)
                        txtContent.Append(matrizBd[i, j]);
                        if (j < cols - 1)
                            txtContent.Append("|");
                    }

                    //Agrega un salto de línea después de cada fila(excepto para la última fila)
                    if (i < rows - 1)
                        txtContent.AppendLine();
                }

                //Guarda el contenido en el archivo
                string directoryPath = tempPath + "\\" + "baseDatos.txt";
                File.WriteAllText(directoryPath, txtContent.ToString());
                Log("Se leyó el archivo Base de Datos");
                return matrizBd;
            }
            else
            {
                string[,]? matrizBd = null;
                string[] lines = File.ReadAllLines(pathBaseDatos);
                int filas = lines.Length;
                int columnas = lines[0].Split('|').Length;

                matrizBd = new string[filas, columnas];

                for (int i = 0; i < filas; i++)
                {
                    string[] elementos = lines[i].Split('|');
                    for (int j = 0; j < columnas; j++)
                    {
                        matrizBd[i, j] = elementos[j];
                    }
                }

                Log("Se leyó el archivo Base de Datos");
                return matrizBd;
            }
        }

        static async Task<string> ObtenerTipoCambio()
        {
            var client = new HttpClient();
            // Obtener la fecha del viernes pasado
            DateTime fechaViernesPasado = ObtenerViernesPasado();

            // Construir la URL con la fecha del viernes pasado
            var url = $"https://www.banxico.org.mx/SieAPIRest/service/v1/series/SF43718/datos/{fechaViernesPasado.ToString("yyyy-MM-dd")}/{fechaViernesPasado.ToString("yyyy-MM-dd")}";

            var request = new HttpRequestMessage
            {
                Method = HttpMethod.Get,
                RequestUri = new Uri(url),
                Headers =
                {
                    { "Bmx-Token", "784f2443cbeacddd3603a01fe71caa078a09144ebbfbfddb9aece05bc74ed309" },
                },
            };

            using (var response = await client.SendAsync(request))
            {
                response.EnsureSuccessStatusCode();
                var responseBody = await response.Content.ReadAsStringAsync();

                // Deserializar la respuesta JSON
                var data = JsonConvert.DeserializeObject<RootObject>(responseBody);

                // Acceder a la fecha y dato
                var tipoCambio = data?.bmx?.series?[0]?.datos?[0]?.dato;

                return tipoCambio;
            }
        }

        private static DateTime ObtenerViernesPasado()
        {
            DateTime hoy = DateTime.Today;
            DateTime viernesPasado = hoy;

            while (viernesPasado.DayOfWeek != DayOfWeek.Friday)
            {
                viernesPasado = viernesPasado.AddDays(-1);
            }

            return viernesPasado;
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
            string className = nameof(ExcelFiller);

            using (System.IO.StreamWriter escritor = new System.IO.StreamWriter(pathLog, true))
            {
                escritor.WriteLine(">" + miFechaActual.ToString() + " > " + className + " > " + metodo + " > " + message);
            }
        }
    }

    public class Dato
    {
        public string fecha { get; set; }
        public string dato { get; set; }
    }

    public class Serie
    {
        public List<Dato> datos { get; set; }
    }

    public class Bmx
    {
        public List<Serie> series { get; set; }
    }

    public class RootObject
    {
        public Bmx bmx { get; set; }
    }
}