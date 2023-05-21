// -----------------------------------------------------------------------------
// ReordenarAsignaciones                                    (21/may/23 00.19)
// Invertir las asignaciones:
// de controles a tipo de datos y viceversa
//
// En este código se utilizan las extensiones AsDecimal, AsInteger, AsDouble, AsDate, AsDateTime y AsTimeSpan
// que extán definidas en una clase (módulo de VB) llamado Extensiones.
//
// Para ver ejemplos, mira en Ejemplos.txt
//
// (c) Guillermo (elGuille) Som, 2023
// -----------------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;
using System.Text;

namespace ReordenarAsignaciones;

internal class Program
{
    // La versión 11
    //#error version

    // Intentar no pasar de estas marcas: 60 caracteres. 2         3         4         5         6
    //                                ---------|---------|---------|---------|---------|---------|
    //[COPIAR]AppDescripcionCopia = " usando snk público y como dotnet tool"

    /// <summary>
    /// La versión actual.
    /// </summary>
    public static string AppFileVersion { get; } = "1.2";

    /// <summary>
    /// La fecha de modificación.
    /// </summary>
    public static string AppFechaVersion { get; } = "21-may-2023";

    /// <summary>
    /// El nombre de la aplicación.
    /// </summary>
    public static string AppName { get; set; } = "Reordenar Asignaciones";

    /// <summary>
    /// El color del texto de la consola.
    /// </summary>
    private static ConsoleColor ColorTexto { get; set; } = ConsoleColor.Yellow;

    /// <summary>
    /// El color del texto de la parte derecha del título.
    /// </summary>
    private static ConsoleColor ColorTituloDerecha { get; set; } = ConsoleColor.Green;

    /// <summary>
    /// El color del texto de la parte izquierda del título.
    /// </summary>
    private static ConsoleColor ColorTituloIzquierda { get; set; } = ConsoleColor.Cyan;

    /// <summary>
    /// La longitud de las líneas del título.
    /// </summary>
    private static int LongitudLineaTitulo { get; set; } = 80;

    /// <summary>
    /// El retorno de carro según sea Windows u otro sistema.
    /// </summary>
    public static string CrLf => RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "\r" : "\r\n";

    [STAThread]
    static void Main(string[] args)
    {
        MostrarTitulo();
        if (args.Length == 0)
        {
            // pedir la opción a usar
            Console.WriteLine("Debes indicar en la línea de comandos lo que hay que procesar.");
        }
        else
        {
            // se supone que se ha indicado lo que hay que hacer.
            string res = InvertirAsignaciones(args);
            // Copiar en el portapapeles
            //ClipBoard.Set(res);
            Console.WriteLine(res);
        }
        Console.WriteLine();
        Console.WriteLine("Pulsa INTRO para terminar.");
        Console.ReadLine();
    }

    /// <summary>
    /// Mostrar el título en la app de consola.
    /// </summary>
    private static void MostrarTitulo()
    {
        //string titulo = $"{AppName} del {AppFechaVersion} - v{AppFileVersion}";
        string tituloIzq = $"{AppName} - v{AppFileVersion}";
        string tituloDer = $"revisión del {AppFechaVersion}";
        int espaciosMid = LongitudLineaTitulo - tituloIzq.Length - tituloDer.Length;
        string tituloSep = new string(' ', espaciosMid);
        string titulo = $"{tituloIzq}{tituloSep}{tituloDer}";

        Console.Clear();
        Console.ForegroundColor = ColorTexto;
        Console.Title = titulo;

        StringBuilder sb = new();
        string lineas = new('-', LongitudLineaTitulo);
        Console.WriteLine(lineas);
        //Console.WriteLine($"{titulo}");
        Console.ForegroundColor = ColorTituloIzquierda;
        Console.Write($"{tituloIzq}");
        Console.Write($"{tituloSep}");
        Console.ForegroundColor = ColorTituloDerecha;
        Console.WriteLine($"{tituloDer}");
        Console.ForegroundColor = ColorTexto;
        sb.AppendLine();
        sb.AppendLine($"Utilidad para invertir las asignaciones.");
        sb.AppendLine(lineas);
        Console.WriteLine(sb.ToString());
        Console.ForegroundColor = ConsoleColor.White;
    }

    /// <summary>
    /// Invertir las asignaciones del array indicado.
    /// </summary>
    /// <param name="args"></param>
    /// <returns>Una cadena con la conversión realizada.</returns>
    /// <remarks>El formato del array es: {"parteIzquierda", "=", "parteDerecha"}
    /// <para>Es decir la asignación estará en tres valores consecutivos:</para>
    /// La parte izquierda, el signo igual y la parte derecha.
    /// </remarks>
    private static string InvertirAsignaciones(string[] args)
    {
        StringBuilder sb = new();
        // Al indicarse en la línea de comandos el formato será {"parteIzquierda", "=", "parteDerecha"}
        for (int i = 0; i < args.Length - 2; i += 3)
        {
            string leftPart = args[i].Trim();
            string rightPart = args[i + 2].Trim().TrimEnd(';');
            if (rightPart.Contains("txt", StringComparison.OrdinalIgnoreCase))
            {
                int j = rightPart.IndexOf(".Text.");
                if (j > -1)
                {
                    if (rightPart.Contains(".AsInteger()"))
                    {
                        leftPart += ".ToString()";
                    }
                    else if (rightPart.Contains(".AsDecimalInt()"))
                    {
                        leftPart += ".ToString()";
                    }
                    else if (rightPart.Contains(".AsDecimal()"))
                    {
                        leftPart += ".ToString(\"0.##\")";
                    }
                    else if (rightPart.Contains(".AsDouble()"))
                    {
                        leftPart += ".ToString(\"0.##\")";
                    }
                    else if (rightPart.Contains(".AsDate()"))
                    {
                        leftPart += ".ToString(\"dd/MM/yyyy\")";
                    }
                    else if (rightPart.Contains(".AsDateTime()"))
                    {
                        leftPart += ".ToString(\"dd/MM/yyyy HH:mm\")";
                    }
                    else if (rightPart.Contains(".AsTimeSpan()"))
                    {
                        leftPart += ".ToString(\"hh\\\\:mm\")";
                    }

                    rightPart = rightPart.Substring(0, j + 5);

                }
                else
                {
                    j = rightPart.IndexOf(".ToString(");
                    if (j > -1)
                    {
                        rightPart = rightPart.Substring(0, j - 1);
                    }
                }
            }
            else
            {
                int j = leftPart.IndexOf(".Text.");
                if (j > -1)
                {
                    if (leftPart.Contains(".AsInteger()"))
                    {
                        rightPart += ".ToString()";
                    }
                    else if (leftPart.Contains(".AsDecimalInt()"))
                    {
                        rightPart += ".ToString()";
                    }
                    else if (leftPart.Contains(".AsDecimal()"))
                    {
                        rightPart += ".ToString(\"0.##\")";
                    }
                    else if (leftPart.Contains(".AsDouble()"))
                    {
                        rightPart += ".ToString(\"0.##\")";
                    }
                    else if (leftPart.Contains(".AsDate()"))
                    {
                        rightPart += ".ToString(\"dd/MM/yyyy\")";
                    }
                    else if (leftPart.Contains(".AsDateTime()"))
                    {
                        rightPart += ".ToString(\"dd/MM/yyyy HH:mm\")";
                    }
                    else if (leftPart.Contains(".AsTimeSpan()"))
                    {
                        rightPart += ".ToString(\"hh\\\\:mm\")";
                    }

                    leftPart = leftPart.Substring(0, j + 5);
                }
                else
                {
                    if (rightPart.Contains(".ToString()"))
                    {
                        leftPart += ".AsInteger()";
                    }
                    else if (rightPart.Contains(".ToString(hh\\\\:mm)"))
                    {
                        leftPart += ".AsTimeSpan()";
                    }
                    else if (rightPart.Contains(".ToString(dd/MM/yyyy HH:mm)"))
                    {
                        leftPart += ".AsDateTime()";
                    }
                    else if (rightPart.Contains(".ToString(dd/MM/yyyy)"))
                    {
                        leftPart += ".AsDate()";
                    }
                    else if (rightPart.Contains(".ToString(0.##)"))
                    {
                        leftPart += ".AsDecimal()";
                    }
                    j = rightPart.IndexOf(".ToString(");
                    if (j > - 1)
                    {
                        rightPart = rightPart.Substring(0, j);
                    }
                }
            }
            sb.Append(rightPart);
            sb.Append(" = ");
            sb.Append(leftPart);
            sb.AppendLine(";");
        }

        return sb.ToString();
    }
}