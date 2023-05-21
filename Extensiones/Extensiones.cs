// -----------------------------------------------------------------------------
// Módulo para extender el funcionamiento de algunos controles      (14/Jul/19)
// No hacer llamada a los métodos de ConversorTipos                 (09/Sep/19)
// 
// Le agrego extensiones que tengo en gsEvaluaColorearCodigo        (14/Oct/20)
// 
// Las conversiones AsTIPO las sobrecargo para usar con String      (28/Mar/21)
// Agrego las definiciones de EsPar, EsImpar, etc.                  (31/Mar/21)
// Agrego la conversión AsDateTime (devuelve fecha y hora)          (09/Oct/21)
// FechaSP utiliza AsDaTime en lugar de AsDate
// (que devuelve solo la parte de la fecha).
// Agrego la propiedad CultureES para usar CultireInfo (es-ES)      (17/Oct/21)
// Agrego la extensión Singular                                     (11/may/23)
//
// Convertido a C# a partir de Extensiones.vb               (21/may/23 20.55)
// con https://converter.telerik.com/
// 
// (c) Guillermo (elGuille) Som, 2019, 2020, 2021, 2023
// -----------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;

/// <summary>
/// Clase Extensiones.
/// </summary>
public static class Extensiones
{
    /// <summary>
    /// Para obtener una fecha con el último día del mes de la fecha indicada.
    /// </summary>
    /// <param name="fecha">La fecha de la que se obtendrá con el último día del mes.</param>
    /// <returns>Una fecha con el último día del mes de la fecha indicada.</returns>
    public static DateTime UltimoDiaMes(this DateTime fecha)
    {
        return new DateTime(fecha.Year, fecha.Month, DateTime.DaysInMonth(fecha.Year, fecha.Month));
    }

    /// <summary>
    /// Comprobar el contenido del texto para aceptar solo:
    /// números y el signo + que es lo que se acepta en los teléfonos
    /// </summary>
    /// <remarks>Si añadirPais es True, se añade +34 si no hay código asignado.</remarks>
    /// <returns>El teléfono con el código internacional (en realidad el + delante si tiene más de 9 caracteres o el +34), 
    /// si no se han indicado números en el texto apsado, se devuelve la misma cadena.</returns>
    public static string ValidarTextoTelefono(this string txt, bool añadirPais)
    {
        var charTelef = "+0123456789".ToCharArray();
        var txtOriginal = txt;

        foreach (var c in txt)
        {
            if (Array.IndexOf<char>(charTelef, c) == -1)
                txt = txt.Replace(c.ToString(), "");
        }

        // Si es una cadena vacía, devolver lo que se indicó (24/abr/23 09.03)
        if (string.IsNullOrWhiteSpace(txt))
            return txtOriginal;

        // esto hacerlo si se indica                                 (18/Abr/19)
        if (añadirPais)
        {
            if (txt.StartsWith("+") == false)
            {
                // Si la longitud es mayor de 9                      (18/Jul/19)
                // tendrá código internacional
                if (txt.Length > 9)
                {
                    // si empieza por 00, quitárselo                 (18/Jul/19)
                    if (txt.StartsWith("00"))
                        txt = txt.Substring(2);
                    txt = "+" + txt;
                }
                else
                    txt = "+34" + txt;
            }
        }

        return txt;
    }

    /// <summary>
    /// Guarda en el fichero indicado los datos de la colección de tipo List(Of String)
    /// </summary>
    /// <param name="fic">El fichero en el que guardar los datos</param>
    /// <param name="colDatos">Lista de tipo cadena a guardar</param>
    /// <remarks>12/Ago/19</remarks>
    public static void GuardarDatos(this string fic, List<string> colDatos)
    {
        // guardar los datos en el fichero
        using (System.IO.StreamWriter sw = new System.IO.StreamWriter(fic, false, Encoding.Default))
        {
            for (var i = 0; i <= colDatos.Count - 1; i++)
                sw.WriteLine(colDatos[i]);
            sw.Close();
        }
    }

    /// <summary>
    /// Devuelve un valor de tipo String 1 si el valor es True, 0 si es False o nulo si es nulo.
    /// </summary>
    /// <param name="valor">El valor de tipo Boolean? a convertir en cadena.</param>
    /// <returns></returns>
    public static string TriState(this bool? valor)
    {
        if (valor.HasValue)
        {
            if (valor.Value == true)
                return "1";
            else
                return "0";
        }
        else
            return "nulo";
    }

    /// <summary>
    /// Convierte un valor de una cadena en un valor Boolean?
    /// </summary>
    /// <param name="valor">La cadena a convertir (ver las notas).</param>
    /// <returns>Un valor Boolean?</returns>
    /// <remarks>
    /// Se tienen en cuenta las cadenas (sin distinguir entre mayúsculas y minúsculas):
    /// 1, Verdadeo y True para true
    /// 0, Falso y False para false
    /// Cadena vacía, nula o con otro valor de los indicados, se devuelve nulo.
    /// </remarks>
    public static bool? TriState(this string valor)
    {
        if (string.IsNullOrWhiteSpace(valor))
            return default(Boolean?);
        switch (valor.ToLower())
        {
            case "1":
            case "verdadero":
            case "true":
                {
                    return true;
                }

            case "0":
            case "falso":
            case "false":
                {
                    return false;
                }

            default:
                {
                    return default(Boolean?);
                }
        }
    }

    /// <summary>
    /// Comprobar si la fecha indicada está entre las otras 2.
    /// </summary>
    /// <param name="fecha">Fecha a comprobar.</param>
    /// <param name="fecha1">Fecha desde.</param>
    /// <param name="fecha2">Fecha hasta.</param>
    /// <returns>True si la fecha está entre las otras 2.</returns>
    /// <remarks>No se comprueban las horas, solo las fechas.</remarks>
    public static bool FechasBetween(this DateTime fecha, DateTime fecha1, DateTime fecha2)
    {
        return fecha.Date >= fecha1.Date && fecha.Date <= fecha2.Date;
    }

    /// <summary>
    /// Comprueba si la fecha indicada está entre la fecha1 y la fecha1 más los días indicados.
    /// </summary>
    /// <param name="fecha">Fecha a comprobar.</param>
    /// <param name="fecha1">Fecha desde.</param>
    /// <param name="dias">Número de días a sumar a fecha1</param>
    /// <returns>True si la fecha está entre las otras 2.</returns>
    /// <remarks>No se comprueban las horas, solo las fechas.</remarks>
    public static bool FechasBetween(this DateTime fecha, DateTime fecha1, int dias)
    {
        return fecha.Date >= fecha1.Date && fecha.Date <= fecha1.Date.AddDays(dias);
    }

    /// <summary>
    /// Busca en el texto todas las cadenas indicadas, 
    /// devuelve true si todos los indicados NO están en el texto.
    /// </summary>
    /// <param name="texto"></param>
    /// <param name="buscar"></param>
    /// <returns></returns>
    public static bool NoContieneTodos(this string texto, params string[] buscar)
    {
        if (buscar == null || buscar.Length == 0)
            return false;

        int cuantos = 0;

        texto = texto.QuitarTildes();
        for (var i = 0; i <= buscar.Length - 1; i++)
        {
            var index = texto.IndexOf(buscar[i].Trim().QuitarTildes(), StringComparison.CurrentCultureIgnoreCase);
            if (index == -1)
                cuantos += 1;
        }
        return cuantos == buscar.Length;
    }

    /// <summary>
    /// Busca en el texto todas las cadenas indicadas, cuenta los encontrados y devuelve true o false según los haya encontrado todos o no.
    /// </summary>
    /// <param name="texto"></param>
    /// <param name="buscar"></param>
    /// <returns></returns>
    public static bool ContieneTodos(this string texto, params string[] buscar)
    {
        if (buscar == null || buscar.Length == 0)
            return false;

        int cuantos = 0;

        texto = texto.QuitarTildes();
        for (var i = 0; i <= buscar.Length - 1; i++)
        {
            var index = texto.IndexOf(buscar[i].Trim().QuitarTildes(), StringComparison.CurrentCultureIgnoreCase);
            if (index == -1)
                break;
            cuantos += 1;
        }
        return cuantos == buscar.Length;
    }

    /// <summary>
    /// Busca en texto cualquiera de las cadenas indicadas y devuelve la posición o -1 si no se ha encontrado.
    /// Tiene preferencia que no se encuentre lo buscado.
    /// </summary>
    /// <param name="texto">La cadena en la que se buscará.</param>
    /// <param name="buscar">Array de tipo cadena con lo que se quiere buscar.</param>
    /// <returns>La posición en base cero en la cadena de cualquiera de los valores de buscar, o -1 si no se ha encontrado.</returns>
    /// <remarks>No diferencia entre mayúsculas y minúsculas y busca sin diferenciar con tildes y se quitan los espacios extras.</remarks>
    public static int ContieneNinguno(this string texto, params string[] buscar)
    {
        int index = -1;
        texto = texto.QuitarTildes();
        for (var i = 0; i <= buscar.Length - 1; i++)
        {
            index = texto.IndexOf(buscar[i].Trim().QuitarTildes(), StringComparison.CurrentCultureIgnoreCase);
            if (index == -1)
                break;
        }
        return index;
    }

    /// <summary>
    /// Busca en texto cualquiera de las cadenas indicadas y devuelve la posición o -1 si no se ha encontrado.
    /// Tiene preferencia que se encuentre lo buscado.
    /// </summary>
    /// <param name="texto">La cadena en la que se buscará.</param>
    /// <param name="buscar">Array de tipo cadena con lo que se quiere buscar.</param>
    /// <returns>La posición en base cero en la cadena de cualquiera de los valores de buscar, o -1 si no se ha encontrado.</returns>
    /// <remarks>No diferencia entre mayúsculas y minúsculas y busca sin diferenciar con tildes y se quitan los espacios extras.</remarks>
    public static int ContieneAlguno(this string texto, params string[] buscar)
    {
        int index = -1;
        texto = texto.QuitarTildes();
        for (var i = 0; i <= buscar.Length - 1; i++)
        {
            index = texto.IndexOf(buscar[i].Trim().QuitarTildes(), StringComparison.CurrentCultureIgnoreCase);
            if (index > -1)
                break; // Return index
        }
        return index;
    }

    /// <summary>
    /// Ajusta el ancho de la cadena según el valor indicado.
    ///     '''</summary>
    ///     '''<remarks>No quita los espacios extras.</remarks>
    public static string AjustarAncho(this string cadena, int ancho)
    {
        System.Text.StringBuilder sb = new System.Text.StringBuilder(new string(' ', ancho));
        return (cadena + sb.ToString()).Substring(0, ancho); // .Trim()
    }

    /// <summary>
    /// Ajusta el ancho de la cadena según el valor indicado.
    /// </summary>
    /// <param name="cadena">El texto a ajustar.</param>
    /// <param name="ancho">El ancho total a devolver.</param>
    /// <param name="conSeparador">Si se muestra con el sepadador en medio o '...' al final. Si es false, se ignoran los valores de alPrincipio y alFinal.</param>
    /// <param name="alPrincipio">Cuántos caracteres al principio, -1 indica la mitad. Al valor indicado/calculado se le restarán 2.</param>
    /// <param name="alFinal">Cuántos caracteres al final, -1 indica la mitad. Al valor indicado/calculado se le restarán 3.</param>
    /// <returns>La nueva cadena</returns>
    /// <remarks>
    /// Entre el principio y el final de la cadena se añadirá ' [-] '
    /// esos 5 caracteres se descuentan del total a mostrar, 
    /// por tanto si el ancho a mostrar es pequeño no tiene mucho sentido usar esta función.
    /// </remarks>
    public static string AjustarAnchoPartido(this string cadena, int ancho, bool conSeparador, int alPrincipio = -1, int alFinal = -1)
    {
        // Si la cadena ya tiene los caracteres a devolver, no hacer nada
        if (string.IsNullOrWhiteSpace(cadena) == false && cadena.Length == ancho)
            return cadena;

        cadena = cadena.Trim();
        var longCadena = cadena.Length;

        if (conSeparador)
        {
            if (alPrincipio == -1)
                alPrincipio = ancho / 2;
            if (alFinal == -1)
                alFinal = ancho / 2;

            // ajustar los anchos del final y el principio
            alPrincipio -= 2;
            alFinal -= 3;
            if (alPrincipio > longCadena / 2)
                alPrincipio = longCadena / 2;
            if (alFinal > longCadena / 2)
                alFinal = longCadena / 2;

            // Solo hacerlo si hay algo en la cadena que ocupe más de el ancho - 5 (del separador)
            if (longCadena > ancho - 5)
                cadena = cadena.Substring(0, alPrincipio) + " [-] " + cadena.Substring(longCadena - alFinal);
        }
        else if (longCadena > ancho)
            cadena = cadena.Substring(0, ancho - 3) + "...";

        StringBuilder sb = new StringBuilder();
        sb.Append(cadena);
        sb.Append(new string(' ', ancho));

        return sb.ToString().Substring(0, ancho);
    }

    // 
    // Definir esta propiedad para que esté compartido.              (17/Oct/21)
    // 

    private static CultureInfo _CultureES;

    /// <summary>
    /// Devuelve un objeto con el valor de la cultura en español-España (es-ES).
    /// </summary>
    /// <returns>Un objeto del tipo <see cref="CultureInfo"/> con la cultura para es-ES.</returns>
    public static CultureInfo CultureES
    {
        get
        {
            if (_CultureES == null)
                _CultureES = CultureInfo.CreateSpecificCulture("es-ES");
            return _CultureES;
        }
    }

    /// <summary>
    /// Devuelve una cadena de la fecha y hora indicada usando un formato en español.
    /// </summary>
    /// <param name="laFecha">Un tipo Object que representa una fecha.</param>
    /// <param name="formato">El formato a usar.</param>
    /// <returns>Una cadena con el formato indicado, pero en cultura es-ES.</returns>
    public static string FechaSP(object laFecha, string formato)
    {
        var fec = laFecha.ToString().AsDateTime();
        if (fec.Year == 1900)
            return laFecha.ToString();
        // Dim culture = System.Globalization.CultureInfo.CreateSpecificCulture("es-ES")
        var sfec = fec.ToString(formato, CultureES);
        return sfec;
    }

    /// <summary>
    /// Devuelve una cadena de la fecha indicada usando un formato en español.
    /// </summary>
    /// <param name="laFecha">La fecha a convertir en cadena.</param>
    /// <param name="formato">El formato a usar.</param>
    /// <returns>Una cadena con el formato indicado, pero en cultura es-ES.</returns>
    public static string FechaSP(this DateTime laFecha, string formato)
    {
        // Dim culture = System.Globalization.CultureInfo.CreateSpecificCulture("es-ES")
        var sfec = laFecha.ToString(formato, CultureES);
        return sfec;
    }

    /// <summary>
    /// Convierte la cadena indicada en singular.
    /// </summary>
    /// <param name="plural"></param>
    /// <param name="conES"></param>
    /// <param name="conN"></param>
    /// <param name="variasPalabras"></param>
    /// <returns></returns>
    public static string Singular(this string plural, bool conES = false, bool conN = false, bool variasPalabras = false)
    {
        var mayusculas = plural == plural.ToLower();

        if (variasPalabras)
        {
            StringBuilder sbSingular = new StringBuilder();
            var col = Palabras(plural);
            for (var i = 0; i <= col.Count - 1; i++)
            {
                col[i] = col[i].Trim().Singular(conES: conES, conN: conN);
                sbSingular.Append($"{col[i]} ");
            }
            plural = sbSingular.ToString().TrimEnd();
        }
        else if (conN)
            plural = plural.TrimEnd("es".ToCharArray());
        else if (conES)
            plural = plural.TrimEnd("s".ToCharArray());
        if (mayusculas)
            return plural.ToLower();

        return plural;
    }

    /// <summary>
    /// Devuelve el plural del texto indicado, según el valor sea distinto de 1.
    /// </summary>
    /// <param name="n">El valor a tener en cuenta (será plural si es distinto de 1).</param>
    /// <param name="singular">La palabra a pluralizar.</param>
    /// <param name="conES">Si el plural debe finalizar con ES en lugar de con S.</param>
    /// <param name="conN">Si el plural debe finalizar con N (queda -> quedan).</param>
    /// <param name="variasPalabras">Si se incluyen varias palabras separadas por espacio.</param>
    /// <returns>La cadena pluralizada o la indicada si no es plural.</returns>
    /// <remarks>Si la palabra en singular es en mayúsculas se devuelve en mayúsculas.</remarks>
    public static string Plural(this string singular, int n, bool conES = false, bool conN = false, bool variasPalabras = false)
    {
        return Plural(n, singular, conES, conN, variasPalabras);
    }

    /// <summary>
    /// Devuelve el plural del texto indicado, según el valor sea distinto de 1.
    /// </summary>
    /// <param name="n">El valor a tener en cuenta (será plural si es distinto de 1).</param>
    /// <param name="singular">La palabra a pluralizar.</param>
    /// <param name="conES">Si el plural debe finalizar con ES en lugar de con S.</param>
    /// <param name="conN">Si el plural debe finalizar con N (queda -> quedan).</param>
    /// <param name="variasPalabras">Si se incluyen varias palabras separadas por espacio.</param>
    /// <returns>La cadena pluralizada o la indicada si no es plural.</returns>
    /// <remarks>Si la palabra en singular es en mayúsculas se devuelve en mayúsculas.</remarks>
    public static string Plural(this int n, string singular, bool conES = false, bool conN = false, bool variasPalabras = false)
    {
        var mayusculas = singular == singular.ToUpper();

        if (n != 1)
        {
            // Poner primero si son varias palabras. v1.10.28.1 (02/sep/22 15.06)
            if (variasPalabras)
            {
                var col = Palabras(singular);
                singular = "";
                for (var i = 0; i <= col.Count - 1; i++)
                {
                    col[i] = col[i].Trim().Plural(n, conES: conES, conN: conN);
                    singular += col[i] + " ";
                }
                singular = singular.TrimEnd();
            }
            else if (conN)
                singular += "n";
            else if (conES)
                singular += "es";
            else
                singular += "s";
        }
        if (mayusculas)
            return singular.ToUpper();
        return singular;
    }

    /// <summary>
    /// Devuelve true si el número indicado es múltiplo exacto del indicado en veces.
    /// </summary>
    /// <param name="numero">El número a comprobar.</param>
    /// <param name="veces">Las veces a comprobar.</param>
    /// <returns>True si el módulo resultante es cero.</returns>
    /// <remarks>Por ejemplo: 126 es múltiplo exacto de 7 = sí, 125 no lo es.</remarks>
    public static bool EsMultiplo(this int numero, int veces)
    {
        // return numero % veces == 0
        return (numero % veces) == 0;
    }

    /// <summary>
    /// Devuelve true si el número indicado es cumple las veces indicadas.
    /// </summary>
    /// <param name="numero">El número a comprobar.</param>
    /// <param name="veces">Las veces a comprobar.</param>
    /// <returns>True si el módulo resultante es cero.</returns>
    /// <remarks>Por ejemplo: 126 es múltiplo exacto de 7 = sí, 125 no lo es.</remarks>
    public static bool EsVeces(this int numero, int veces)
    {
        return EsMultiplo(numero, veces);
    }

    /// <summary>
    /// Devuelve si un número es par
    /// </summary>
    /// <remarks>27/May/19</remarks>
    public static bool IsEven(this int n)
    {
        return EsPar(n);
    }

    /// <summary>
    /// Devuelve si un número es par
    /// </summary>
    /// <remarks>27/May/19</remarks>
    public static bool EsPar(this int n)
    {
        // return value % 2 == 0
        return (n % 2) == 0;
    }

    /// <summary>
    /// Devuelve si un número es impar
    /// </summary>
    /// <remarks>27/May/19</remarks>
    public static bool IsOdd(this int n)
    {
        return EsImpar(n);
    }

    /// <summary>
    /// Devuelve si un número es impar
    /// </summary>
    /// <remarks>27/May/19</remarks>
    public static bool EsImpar(this int n)
    {
        return (n % 2) != 0;
    }

    /// <summary>
    /// Quitar de una cadena un texto indicado (que será el predeterminado cuando está vacío).
    /// Por ejemplo si el texto grisáceo es Buscar... y
    /// se empezó a escribir en medio del texto (o en cualquier parte)
    /// BuscarL... se quitará Buscar... y se dejará L.
    /// Antes de hacer cambios se comprueba si el texto predeterminado está al completo 
    /// en el texto en el que se hará el cambio.
    /// </summary>
    /// <param name="texto">El texto en el que se hará la sustitución.</param>
    /// <param name="predeterminado">El texto a quitar.</param>
    /// <returns>Una cadena con el texto predeterminado quitado.</returns>
    /// <remarks>18/Oct/2020 actualizado 24/Oct/2020</remarks>
    public static string QuitarPredeterminado(this string texto, string predeterminado)
    {
        var cuantos = predeterminado.Length;
        var k = 0;

        for (var i = 0; i <= predeterminado.Length - 1; i++)
        {
            var j = texto.IndexOf(predeterminado[i]);
            if (j == -1)
                continue;
            k += 1;
        }
        // si k es distinto de cuantos es que no están todos lo caracteres a quitar
        if (k != cuantos)
            return texto;

        for (var i = 0; i <= predeterminado.Length - 1; i++)
        {
            var j = texto.IndexOf(predeterminado[i]);
            if (j == -1)
                continue;
            if (j == 0)
                texto = texto.Substring(j + 1);
            else
                texto = texto.Substring(0, j) + texto.Substring(j + 1);
        }

        return texto;
    }

    /// <summary>
    /// Devuelve true si el texto indicado contiene alguna letra del alfabeto.
    /// Incluída la Ñ y vocales con tilde.
    /// </summary>
    /// <param name="texto"></param>
    /// <returns></returns>
    /// <remaks>14/Oct/2020</remaks>
    public static bool ContieneLetras(this string texto)
    {
        var letras = "abcdefghijklmnñopqurstuvwxyzáéíóúü";
        return texto.ToLower().IndexOfAny(letras.ToCharArray()) > -1;
    }

    /// <summary>
    /// Quitar las tildes de una cadena.
    /// </summary>
    /// <param name="s">La cadena a extender donde se buscarán las tildes.</param>
    /// <remarks>
    /// 03/Ago/2020
    /// 27/Jun/2021 Usando StringBuilder en vez de concatenación.
    /// </remarks>
    public static string QuitarTildes(this string s)
    {
        var tildes1 = "ÁÉÍÓÚÜáéíóúü";
        var tildes0 = "AEIOUUaeiouu";
        // Dim res = ""
        StringBuilder res = new StringBuilder();
        int j;

        for (var i = 0; i <= s.Length - 1; i++)
        {
            // Dim j = tildes1.IndexOf(s(i))
            j = tildes1.IndexOf(s[i]);
            if (j > -1)
                // res &= tildes0.Substring(j, 1)
                res.Append(tildes0.Substring(j, 1));
            else
                // res &= s(i)
                res.Append(s[i]);
        }
        return res.ToString();
    }

    // 
    // Conversiones (AsTIPO) usando cadenas en vez de controles      (28/Mar/21)
    // 

    // 28/Mar/21
    /// <summary>
    /// Devuelve un valor Integer de la cadena indicada.
    /// </summary>
    /// <param name="txt">La cadena a extender</param>
    /// <remarks>17-oct-22: La conversión falla si el texto tiene decimales</remarks>
    public static int AsInteger(this string txt)
    {
        int i = 0;

        // La conversión falla si el texto tiene decimales. (17/oct/22 12.16)
        // Si falla: Convertir primero a double y redondearlo.
        if (int.TryParse(txt, ref i) == false)
            i = System.Convert.ToInt32(AsDoubleInt(txt));

        return i;
    }

    /// <summary>
    /// Devuelve solo la parte de la fecha de la cadena indicada.
    /// </summary>
    /// <param name="txt">La cadena a extender</param>
    /// <remarks>28/Mar/21</remarks>
    public static DateTime AsDate(this string txt)
    {
        return AsDateTime(txt).Date;
    }

    /// <summary>
    /// Devuelve un valor DateTime (fecha y hora) de la cadena indicada.
    /// </summary>
    /// <param name="txt">La cadena a extender</param>
    /// <remarks>09/Oct/21</remarks>
    public static DateTime AsDateTime(this string txt)
    {
        DateTime d = new DateTime(1900, 1, 1, 0, 0, 0);

        if (!(string.IsNullOrWhiteSpace(txt) || txt.Equals(DBNull.Value)))
        {
            // Comprobar si tiene caracteres para cambiar            (07/Oct/20)
            txt = txt.Replace(".", "/").Replace("-", "/");

            // Dim culture = Globalization.CultureInfo.CreateSpecificCulture("es-ES")
            var styles = System.Globalization.DateTimeStyles.None;

            // Usar siempre la conversión al estilo de España        (29/Mar/21)
            // Devuelve false si no se ha convertido y la fecha es DateTime.MinValue
            // Comprobarlo así por si falla la conversión.   (15/abr/23 06.29)
            if (DateTime.TryParse(txt, CultureES, styles, ref d) == false)
            {
                // Volver a intentarlo                               (01/May/21)
                // (por si la fecha está en formato "guiri")
                if (DateTime.TryParse(txt, ref d) == false)
                    d = new DateTime(1900, 1, 1, 0, 0, 0);
            }
        }
        else
            // asignar el 01/01/1900 si es un valor en blanco        (07/Jul/15)
            d = new DateTime(1900, 1, 1, 0, 0, 0);

        return d;
    }

    /// <summary>
    /// Devuelve el valor Integer redondeado (usando Math.Round) de la cadena indicada y tratado como Decimal.
    /// </summary>
    /// <param name="txt">La cadena a extender</param>
    /// <remarks>28/Mar/21</remarks>
    public static int AsDecimalInt(this string txt)
    {
        return System.Convert.ToInt32(Math.Round(txt.AsDecimal()));
    }

    /// <summary>
    /// Devuelve un entero redondeado (usando Math.Round) del decimal indicado.
    /// </summary>
    /// <param name="txt"></param>
    /// <returns></returns>
    /// <remarks>01/Abr/21</remarks>
    public static int AsDecimalInt(this decimal txt)
    {
        return System.Convert.ToInt32(Math.Round(txt));
    }

    /// <summary>
    /// Devuelve un valor Decimal de la cadena indicada.
    /// </summary>
    /// <param name="txt">La cadena a extender</param>
    /// <remarks>28/Mar/21</remarks>
    public static decimal AsDecimal(this string txt)
    {
        decimal d = 0;

        // La conversión con decimales da problemas
        // Dim style = NumberStyles.Number Or NumberStyles.AllowCurrencySymbol 'Or NumberStyles.AllowDecimalPoint

        if (string.IsNullOrWhiteSpace(txt))
            txt = "0";

        // Si tiene símbolo de moneda, quitarlo. v1.10.13.2 (26/jul/22 21.51)
        if (txt.IndexOfAny("€$".ToCharArray()) > -1)
            txt = txt.Replace("€", "").Replace("$", "").Trim();

        // Decimal.TryParse(txt, style, CultureES, d)
        // No convertir en español por si se usa otro idioma para los decimales. v1.10.13.2 (26/jul/22 21.45)

        // Si todo fue bien, devuelve el valor convertido, si el texto indicado no es un número devuelve 0
        decimal.TryParse(txt, ref d);

        return d;
    }

    /// <summary>
    /// Devuelve un valor Double de la cadena indicada.
    /// </summary>
    /// <param name="txt">La cadena a extender</param>
    /// <remarks>18-sep-22: v1.20.2.14</remarks>
    public static double AsDouble(this string txt)
    {
        double d = 0;

        if (string.IsNullOrWhiteSpace(txt))
            txt = "0";

        double.TryParse(txt, ref d);

        return d;
    }

    /// <summary>
    /// Devuelve un valor Double de la cadena indicada y después lo redondea.
    /// </summary>
    /// <param name="txt">La cadena a extender</param>
    /// <remarks>17-oct-22</remarks>
    public static double AsDoubleInt(this string txt)
    {
        double d = 0;

        if (string.IsNullOrWhiteSpace(txt))
            txt = "0";

        double.TryParse(txt, ref d);
        d = Math.Round(d);

        return d;
    }

    /// <summary>
    /// Devuelve un valor Boolean de la cadena indicada.
    /// </summary>
    /// <param name="txt">La cadena a extender</param>
    /// <remarks>28/Mar/21</remarks>
    public static bool AsBoolean(this string txt)
    {
        if (txt == "" || txt == "0")
            return false;
        return System.Convert.ToBoolean(txt);
    }

    /// <summary>
    /// Devuelve un valor TimeSpan de la cadena indicada.
    /// </summary>
    /// <param name="txt">La cadena a extender</param>
    /// <remarks>28/Mar/21</remarks>
    public static TimeSpan AsTimeSpan(this string txt)
    {
        TimeSpan c = new TimeSpan(0, 0, 0);

        if (string.IsNullOrWhiteSpace(txt))
            return c;

        // Solo cambiar los puntos por : si no tiene :               (21/Jun/21)
        // ya que pueden ser milisegundos...
        if (txt.Contains(".") && txt.Contains(":") == false)
            txt = txt.Replace(".", ":");
        else if (txt.Contains(":") == false)
            txt += ":00";
        if (txt == ":00")
            txt = "00:00";

        TimeSpan.TryParse(txt, ref c);

        return c;
    }

    // 
    // De las extensiones de gsEvaluarColorearCodigo
    // 

    // 
    // Extensiones reemplazar si no está lo que se va a reemplazar   (04/Oct/20)
    // 

    /// <summary>
    /// Reemplazar buscar por poner si no está poner.
    /// </summary>
    /// <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    /// <param name="buscar">La cadena a buscar sin distinguir entre mayúsculas y minúsculas.</param>
    /// <param name="poner">La cadena a poner si previamente no está.</param>
    /// <returns>Una cadena con los cambios realizados.</returns>
    public static string ReplaceSiNoEstaPoner(this string texto, string buscar, string poner)
    {
        var j = texto.IndexOf(poner);
        // si está lo que se quiere poner, devolver la cadena actual sin cambios
        if (j > -1)
            return texto;

        return texto.Replace(buscar, poner);
    }

    /// <summary>
    /// Reemplazar buscar por poner si no está poner.
    /// </summary>
    /// <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    /// <param name="buscar">La cadena a buscar usando la compración indicada.</param>
    /// <param name="poner">La cadena a poner si previamente no está.</param>
    /// <param name="comparar">El tipo de comparación a relizar: Ordinal o OrdinalIgnoreCase.</param>
    /// <returns>Una cadena con los cambios realizados.</returns>
    public static string ReplaceSiNoEstaPoner(this string texto, string buscar, string poner, StringComparison comparar)
    {
        var j = texto.IndexOf(poner, comparar);
        // si está lo que se quiere poner, devolver la cadena actual sin cambios
        if (j > -1)
            return texto;

        // esta sobrecarga está en la versión 5.0.0.0 no en la 4.0.0.0
        // Return texto.Replace(buscar, poner, comparar)
        if (comparar == StringComparison.OrdinalIgnoreCase)
        {
            // Return texto.Replace(buscar, poner)
            int i;
            do
            {
                i = texto.IndexOf(buscar, comparar);
                if (i == -1)
                    break;
                if (i > 0)
                    texto = poner + texto.Substring(i + buscar.Length);
                else
                    texto = texto.Substring(0, i) + poner + texto.Substring(i + buscar.Length);
            }
            while (true); // While i > -1
            return texto;
        }
        else
            return texto.Replace(buscar, poner);
    }

    /// <summary>
    /// Reemplazar buscar por poner si no está poner.
    /// </summary>
    /// <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    /// <param name="buscar">La cadena a buscar (palabra completa) usando la comparación indicada.</param>
    /// <param name="poner">La cadena a poner si previamente no está.</param>
    /// <param name="comparar">El tipo de comparación a relizar: Ordinal o OrdinalIgnoreCase.</param>
    /// <returns>Una cadena con los cambios realizados.</returns>
    public static string ReplaceWordSiNoEstaPoner(this string texto, string buscar, string poner, StringComparison comparar)
    {
        var j = texto.IndexOf(poner, comparar);
        // si está lo que se quiere poner, devolver la cadena actual sin cambios
        if (j > -1)
            return texto;

        return ReplaceWord(texto, buscar, poner, comparar);
    }

    // 
    // Extensión quitar todos los espacios
    // 

    /// <summary>
    /// Quitar todos los espacios que tenga la cadena,
    /// incluidos los que están entre palabras.
    /// </summary>
    /// <param name="texto">Cadena a la que se quitarán los espacios.</param>
    /// <returns>Una nueva cadena con TODOS los espacios quitados.</returns>
    public static string QuitarTodosLosEspacios(this string texto)
    {
        MatchCollection col = Regex.Matches(texto, @"\S+");
        StringBuilder sb = new StringBuilder();
        foreach (Match m in col)
            sb.Append(m.Value);

        return sb.ToString();
    }

    // 
    // Extensión contar palabras y saber las palabras usando Regex.
    // 

    /// <summary>
    /// Contar las palabras de una cadena de texto usando <see cref="Regex"/>.
    /// </summary>
    /// <param name="texto">El texto con las palabras a contar.</param>
    /// <returns>Un valor entero con el número de palabras</returns>
    /// <example>
    /// Adaptado usando una cadena en vez del Text del RichTextBox
    /// (sería del RichTextBox para WinForms)
    /// El código lo he adaptado de:
    /// https://social.msdn.microsoft.com/Forums/en-US/
    ///     81e438ed-9d35-47d7-a800-1fabab0f3d52/
    ///     c-how-to-add-a-word-counter-to-a-richtextbox
    ///     ?forum=csharplanguage
    /// </example>
    public static int CuantasPalabras(this string texto)
    {
        MatchCollection col = Regex.Matches(texto, @"[\W]+");
        return col.Count;
    }

    // 
    // Extensiones de cadena y cambiar a mayúsculas/minúsculas       (01/Oct/20)
    // 

    public enum CasingValues : int
    {
        /// <summary>
        ///         ''' No se hacen cambios
        ///         ''' </summary>
        Normal,
        /// <summary>
        ///         ''' Todas las letras a mayúsculas
        ///         ''' </summary>
        Upper,
        /// <summary>
        ///         ''' Todas las letras a minúsculas.
        ///         ''' </summary>
        Lower,
        /// <summary>
        ///         ''' La primera letra de cada palabra a mayúsculas.
        ///         ''' </summary>
        Title,
        /// <summary>
        ///         ''' La primera letra de cada palabra en mayúsculas.
        ///         ''' Equivalente a <see cref="Title"/>.
        ///         ''' </summary>
        FirstToUpper,
        /// <summary>
        ///         ''' La primera letra de cada palabra en minúsculas
        ///         ''' </summary>
        FirstToLower
    }

    /// <summary>
    /// Cambia el texto a Upper, Lower, TitleCase/FirstToUpper o FirstToLower.
    /// Se devuelve una nueva cadena con los cambios.
    /// Valores posibles:
    /// Normal
    /// Upper
    /// Lower
    /// Title o FirstToLower
    /// FirstToLower
    /// </summary>
    /// <param name="text">La cadena a la que se aplicará</param>
    /// <param name="queCase">Un valor </param>
    /// <returns>Una cadena con los cambios</returns>
    public static string CambiarCase(this string text, CasingValues queCase)
    {
        switch (queCase)
        {
            case CasingValues.Lower:
                {
                    text = text.ToLower();
                    break;
                }

            case CasingValues.Upper:
                {
                    text = text.ToUpper();
                    break;
                }

            case CasingValues.Title:
            case CasingValues.FirstToUpper // Title
     :
                {
                    text = ToTitle(text);
                    break;
                }

            case CasingValues.FirstToLower // camelCase
     :
                {
                    text = ToLowerFirst(text);
                    break;
                }

            default:
                {
                    break;
                }
        }

        return text;
    }

    /// <summary>
    /// Devuelve una cadena en tipo Título
    /// la primera letra de cada palabra en mayúsculas.
    /// Usando System.Globalization.CultureInfo.CurrentCulture
    /// que es más eficaz que
    /// System.Threading.Thread.CurrentThread.CurrentCulture
    /// </summary>
    public static string ToTitle(this string text)
    {
        // según la documentación usar CultureInfo.CurrentCulture es más eficaz
        // que CurrentThread.CurrentCulture
        var cultureInfo = System.Globalization.CultureInfo.CurrentCulture;
        var txtInfo = cultureInfo.TextInfo;
        if (text == null)
            return "";
        return txtInfo.ToTitleCase(text);
    }

    /// <summary>
    /// Devuelve la cadena indicada con el primer carácter en minúsculas.
    /// Si tiene espacios delante, pone en minúscula el primer carácter que no sea espacio.
    /// </summary>
    /// <param name="text"></param>
    /// <returns></returns>
    public static string ToLowerFirstChar(this string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return text;

        StringBuilder sb = new StringBuilder();
        var b = false;
        for (var i = 0; i <= text.Length - 1; i++)
        {
            if (!b && !char.IsWhiteSpace(text[i]))
            {
                sb.Append(text[i].ToString().ToLower());
                b = true;
            }
            else
                sb.Append(text[i]);
        }

        return sb.ToString();
    }

    /// <summary>
    /// Convierte en minúsculas el primer carácter de cada palabra en la cadena indicada.
    /// </summary>
    public static string ToLowerFirst(this string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return text;

        var col = Palabras(text);
        StringBuilder sb = new StringBuilder();
        for (var i = 0; i <= col.Count - 1; i++)
            sb.AppendFormat("{0}", col[i].ToLowerFirstChar());

        return sb.ToString();
    }

    /// <summary>
    /// Devuelve una cadena en tipo Titulo o nada si es nothing
    /// </summary>
    /// <remarks>25/May/19</remarks>
    public static string ToTitle(this object obj)
    {
        // según la documentación usar CultureInfo.CurrentCulture es más eficaz
        // que CurrentThread.CurrentCulture
        if (obj == null || obj.Equals(DBNull.Value))
            return "";
        // Return ToTitle(obj.ToString)
        return obj.ToString().ToTitle();
    }

    /// <summary>
    /// Devuelve una cadena en minúsculas usando la cultura actual.
    /// </summary>
    public static string ToLower(string text)
    {
        var cultureInfo = System.Globalization.CultureInfo.CurrentCulture;
        var txtInfo = cultureInfo.TextInfo;
        if (text == null)
            return "";
        return txtInfo.ToLower(text);
    }

    /// <summary>
    /// Devuelve una cadena en minúsculas usando la cultura actual.o nada si es nulo.
    /// </summary>
    /// <remarks>25/May/19</remarks>
    public static string ToLower(object obj)
    {
        if (obj == null || obj.Equals(DBNull.Value))
            return "";
        return ToLower(obj.ToString());
    }

    /// <summary>
    /// Devuelve una cadena en mayúsculas usando la cultura actual.
    /// </summary>
    public static string ToUpper(string text)
    {
        var cultureInfo = System.Globalization.CultureInfo.CurrentCulture;
        var txtInfo = cultureInfo.TextInfo;
        if (text == null)
            return "";
        return txtInfo.ToUpper(text);
    }

    /// <summary>
    /// Devuelve una cadena en mayúsculas usando la cultura actual o nada si es nulo.
    /// </summary>
    /// <remarks>25/May/19</remarks>
    public static string ToUpper(object obj)
    {
        if (obj == null || obj.Equals(DBNull.Value))
            return "";
        return ToUpper(obj.ToString());
    }

    /// <summary>
    /// Devuelve una cadena o nada si es nulo
    /// </summary>
    /// <remarks>25/May/19</remarks>
    public static string ToStringVacia(object obj)
    {
        if (obj == null || obj.Equals(DBNull.Value))
            return "";
        return obj.ToString();
    }

    /// <summary>
    /// Devuelve un espacio si es nulo o es una cadena vacía
    /// </summary>
    /// <remarks>25/May/19</remarks>
    public static string ToStringUnEspacio(object obj)
    {
        if (obj == null || obj.Equals(DBNull.Value))
            return " ";
        if (obj.ToString() == "")
            return " ";
        return obj.ToString();
    }

    /// <summary>
    /// Devuelve una lista con las palabras del texto indicado.
    /// </summary>
    /// <param name="text">La cadena de la que se extraerán las palabras.</param>
    /// <returns></returns>
    /// <remarks>
    /// En realidad no devuelve solo las palabras,
    /// ya que cada elemento contendrá los espacios y otros símbolos que estén con esa palabra:
    /// Si la palabra tiene espacios delante también los añade, si tiene un espacio o un símbolo detrás
    /// también lo añade.
    /// Si al final hay espacios en blanco, los elimina.
    /// </remarks>
    /// <example>    Private Sub Hola(str As String) 
    /// Devolverá: "    Private ", "Sub ", "Hola(", "str ", "As ", "String)"
    /// </example>
    public static List<string> Palabras(this string text)
    {
        // busca palabra con (o sin) espacios delante (\s*),
        // cualquier cosa (.),
        // una o más palabras (\w+) y
        // cualquier cosa (.)
        var s = @"\s*.\w+.";
        var res = Regex.Matches(text, s);
        List<string> col = new List<string>();
        foreach (Match m in res)
            col.Add(m.Value);

        return col;
    }

    // 
    // Versiones si se comprueban mayúsculas y minúsculas            (04/Oct/20)
    // 

    /// <summary>
    /// Reemplaza todas las ocurrencias de 'buscar' por 'poner' en el texto,
    /// teniendo en cuenta mayúsculas y minúsculas en la cadena a buscar.
    /// </summary>
    /// <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    /// <param name="buscar">El valor a buscar (palabra completa) distingue mayúsculas y minúsculas.</param>
    /// <param name="poner">El nuevo valor a poner.</param>
    /// <returns>Una cadena con los cambios.</returns>
    public static string ReplaceWordOrdinal(this string texto, string buscar, string poner)
    {
        return ReplaceWord(texto, buscar, poner, StringComparison.Ordinal);
    }

    /// <summary>
    /// Reemplaza todas las ocurrencias de 'buscar' por 'poner' en el texto,
    /// ignorando mayúsculas y minúsculas en la cadena a buscar.
    /// </summary>
    /// <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    /// <param name="buscar">El valor a buscar (palabra completa) sin distinguir mayúsculas y minúsculas.</param>
    /// <param name="poner">El nuevo valor a poner.</param>
    /// <returns>Una cadena con los cambios.</returns>
    public static string ReplaceWordIgnoreCase(this string texto, string buscar, string poner)
    {
        return ReplaceWord(texto, buscar, poner, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Devuelve una nueva cadena en la que todas las apariciones de oldValue
    /// en la instancia actual se reemplazan por newValue, teniendo en cuenta
    /// que se buscarán palabras completas.
    /// </summary>
    /// <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    /// <param name="buscar">El valor a buscar (palabra completa).</param>
    /// <param name="poner">El nuevo valor a poner.</param>
    /// <param name="comparar">El tipo de comparación Ordinal / OrdinalIgnoreCase.</param>
    /// <returns>Una cadena con los cambios.</returns>
    /// <remarks>Código convertido del original en C# de palota:
    /// https://stackoverflow.com/a/62782791/14338047</remarks>
    public static string ReplaceWord(this string texto, string buscar, string poner, StringComparison comparar)
    {
        var IsWordChar = char c => char.IsLetterOrDigit(c) || c == '_';

        StringBuilder sb = null;
        int p = 0;
        int j = 0;

        while (j < texto.Length && __Assign(ref j, texto.IndexOf(buscar, j, comparar)) >= 0)
        {
            if ((j == 0 || !IsWordChar(texto[j - 1])) && (j + buscar.Length == texto.Length || !IsWordChar(texto[j + buscar.Length])))
            {
                sb = sb ?? new StringBuilder();
                sb.Append(texto, p, j - p);
                sb.Append(poner);
                j += buscar.Length;
                p = j;
            }
            else
                j += 1;
        }

        if (sb == null)
            return texto;
        sb.Append(texto, p, texto.Length - p);
        return sb.ToString();
    }

    /// <summary>
    /// Función para la equivalencia en C# de:
    /// while (j &lt; text.Length &amp;&amp; (j = unvalor) >=0 )
    /// </summary>
    /// <typeparam name="T">El tipo de datos</typeparam>
    /// <param name="target">La variable a la que se le asignará el valor de la expresión de value</param>
    /// <param name="value">Expresión con el valor a asignar a target</param>
    /// <returns>Devuelve el valor de value</returns>
    private static T __Assign<T>(ref T target, T value)
    {
        target = value;
        return value;
    }
}
