'------------------------------------------------------------------------------
' Módulo para extender el funcionamiento de algunos controles       (14/Jul/19)
' No hacer llamada a los métodos de ConversorTipos                  (09/Sep/19)
'
' Le agrego extensiones que tengo en gsEvaluaColorearCodigo         (14/Oct/20)
'
' Las conversiones AsTIPO las sobrecargo para usar con String       (28/Mar/21)
' Agrego las definiciones de EsPar, EsImpar, etc.                   (31/Mar/21)
' Agrego la conversión AsDateTime (devuelve fecha y hora)           (09/Oct/21)
'   FechaSP utiliza AsDaTime en lugar de AsDate
'   (que devuelve solo la parte de la fecha).
' Agrego la propiedad CultureES para usar CultireInfo (es-ES)       (17/Oct/21)
' Agrego la extensión Singular                                      (11/may/23)
'
' (c) Guillermo (elGuille) Som, 2019, 2020, 2021, 2023
'------------------------------------------------------------------------------
Option Strict On
Option Infer On

Imports System
Imports System.Text
Imports System.Collections.Generic
Imports System.Linq
'Imports System.Xml.Linq
'Imports System.IO
Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Globalization

''' <summary>
''' Clase Extensiones definida en KNDatos.
''' </summary>
Public Module Extensiones

    ''' <summary>
    ''' Para obtener una fecha con el último día del mes de la fecha indicada.
    ''' </summary>
    ''' <param name="fecha">La fecha de la que se obtendrá con el último día del mes.</param>
    ''' <returns>Una fecha con el último día del mes de la fecha indicada.</returns>
    <Extension>
    Public Function UltimoDiaMes(fecha As Date) As Date
        Return New Date(fecha.Year, fecha.Month, DateTime.DaysInMonth(fecha.Year, fecha.Month))
    End Function

    ''' <summary>
    ''' Comprobar el contenido del texto para aceptar solo:
    ''' números y el signo + que es lo que se acepta en los teléfonos
    ''' </summary>
    ''' <remarks>Si añadirPais es True, se añade +34 si no hay código asignado.</remarks>
    ''' <returns>El teléfono con el código internacional (en realidad el + delante si tiene más de 9 caracteres o el +34), 
    ''' si no se han indicado números en el texto apsado, se devuelve la misma cadena.</returns>
    <Extension>
    Public Function ValidarTextoTelefono(txt As String, añadirPais As Boolean) As String
        Dim charTelef = "+0123456789".ToCharArray()
        Dim txtOriginal = txt

        For Each c In txt
            If Array.IndexOf(Of Char)(charTelef, c) = -1 Then
                txt = txt.Replace(c.ToString, "")
            End If
        Next

        ' Si es una cadena vacía, devolver lo que se indicó (24/abr/23 09.03)
        If String.IsNullOrWhiteSpace(txt) Then
            Return txtOriginal
        End If

        ' esto hacerlo si se indica                                 (18/Abr/19)
        If añadirPais Then
            If txt.StartsWith("+") = False Then
                ' Si la longitud es mayor de 9                      (18/Jul/19)
                ' tendrá código internacional
                If txt.Length > 9 Then
                    ' si empieza por 00, quitárselo                 (18/Jul/19)
                    If txt.StartsWith("00") Then
                        txt = txt.Substring(2)
                    End If
                    txt = "+" & txt
                Else
                    txt = "+34" & txt
                End If
            End If
        End If

        Return txt
    End Function

    ''' <summary>
    ''' Guarda en el fichero indicado los datos de la colección de tipo List(Of String)
    ''' </summary>
    ''' <param name="fic">El fichero en el que guardar los datos</param>
    ''' <param name="colDatos">Lista de tipo cadena a guardar</param>
    ''' <remarks>12/Ago/19</remarks>
    <Extension>
    Public Sub GuardarDatos(fic As String, colDatos As List(Of String))
        ' guardar los datos en el fichero
        Using sw As New System.IO.StreamWriter(fic, False, Encoding.Default)
            For i = 0 To colDatos.Count - 1
                sw.WriteLine(colDatos(i))
            Next
            sw.Close()
        End Using
    End Sub

    ''' <summary>
    ''' Devuelve un valor de tipo String 1 si el valor es True, 0 si es False o nulo si es nulo.
    ''' </summary>
    ''' <param name="valor">El valor de tipo Boolean? a convertir en cadena.</param>
    ''' <returns></returns>
    <Extension>
    Public Function TriState(valor As Boolean?) As String
        If valor.HasValue Then
            If valor.Value = True Then
                Return "1"
            Else 'If valor.Value = False Then
                Return "0"
            End If
        Else
            Return "nulo"
        End If
    End Function

    ''' <summary>
    ''' Convierte un valor de una cadena en un valor Boolean?
    ''' </summary>
    ''' <param name="valor">La cadena a convertir (ver las notas).</param>
    ''' <returns>Un valor Boolean?</returns>
    ''' <remarks>
    ''' Se tienen en cuenta las cadenas (sin distinguir entre mayúsculas y minúsculas):
    ''' 1, Verdadeo y True para true
    ''' 0, Falso y False para false
    ''' Cadena vacía, nula o con otro valor de los indicados, se devuelve nulo.
    ''' </remarks>
    <Extension>
    Public Function TriState(valor As String) As Boolean?
        If String.IsNullOrWhiteSpace(valor) Then
            Return Nothing
        End If
        Select Case valor.ToLower()
            Case "1", "verdadero", "true"
                Return True
            Case "0", "falso", "false"
                Return False
            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>
    ''' Comprobar si la fecha indicada está entre las otras 2.
    ''' </summary>
    ''' <param name="fecha">Fecha a comprobar.</param>
    ''' <param name="fecha1">Fecha desde.</param>
    ''' <param name="fecha2">Fecha hasta.</param>
    ''' <returns>True si la fecha está entre las otras 2.</returns>
    ''' <remarks>No se comprueban las horas, solo las fechas.</remarks>
    <Extension>
    Public Function FechasBetween(fecha As Date, fecha1 As Date, fecha2 As Date) As Boolean
        Return fecha.Date >= fecha1.Date AndAlso fecha.Date <= fecha2.Date
    End Function

    ''' <summary>
    ''' Comprueba si la fecha indicada está entre la fecha1 y la fecha1 más los días indicados.
    ''' </summary>
    ''' <param name="fecha">Fecha a comprobar.</param>
    ''' <param name="fecha1">Fecha desde.</param>
    ''' <param name="dias">Número de días a sumar a fecha1</param>
    ''' <returns>True si la fecha está entre las otras 2.</returns>
    ''' <remarks>No se comprueban las horas, solo las fechas.</remarks>
    <Extension>
    Public Function FechasBetween(fecha As Date, fecha1 As Date, dias As Integer) As Boolean
        Return fecha.Date >= fecha1.Date AndAlso fecha.Date <= fecha1.Date.AddDays(dias)
    End Function

    ''' <summary>
    ''' Busca en el texto todas las cadenas indicadas, 
    ''' devuelve true si todos los indicados NO están en el texto.
    ''' </summary>
    ''' <param name="texto"></param>
    ''' <param name="buscar"></param>
    ''' <returns></returns>
    <Extension>
    Public Function NoContieneTodos(texto As String, ParamArray buscar As String()) As Boolean
        If buscar Is Nothing OrElse buscar.Length = 0 Then Return False

        Dim cuantos As Integer = 0

        texto = texto.QuitarTildes()
        For i = 0 To buscar.Length - 1
            Dim index = texto.IndexOf(buscar(i).Trim().QuitarTildes(), StringComparison.CurrentCultureIgnoreCase)
            If index = -1 Then
                cuantos += 1
            End If
        Next
        Return cuantos = buscar.Length
    End Function

    ''' <summary>
    ''' Busca en el texto todas las cadenas indicadas, cuenta los encontrados y devuelve true o false según los haya encontrado todos o no.
    ''' </summary>
    ''' <param name="texto"></param>
    ''' <param name="buscar"></param>
    ''' <returns></returns>
    <Extension>
    Public Function ContieneTodos(texto As String, ParamArray buscar As String()) As Boolean
        If buscar Is Nothing OrElse buscar.Length = 0 Then Return False

        Dim cuantos As Integer = 0

        texto = texto.QuitarTildes()
        For i = 0 To buscar.Length - 1
            Dim index = texto.IndexOf(buscar(i).Trim().QuitarTildes(), StringComparison.CurrentCultureIgnoreCase)
            If index = -1 Then Exit For
            cuantos += 1
        Next
        Return cuantos = buscar.Length
    End Function

    ''' <summary>
    ''' Busca en texto cualquiera de las cadenas indicadas y devuelve la posición o -1 si no se ha encontrado.
    ''' Tiene preferencia que no se encuentre lo buscado.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se buscará.</param>
    ''' <param name="buscar">Array de tipo cadena con lo que se quiere buscar.</param>
    ''' <returns>La posición en base cero en la cadena de cualquiera de los valores de buscar, o -1 si no se ha encontrado.</returns>
    ''' <remarks>No diferencia entre mayúsculas y minúsculas y busca sin diferenciar con tildes y se quitan los espacios extras.</remarks>
    <Extension>
    Public Function ContieneNinguno(texto As String, ParamArray buscar As String()) As Integer
        Dim index As Integer = -1
        texto = texto.QuitarTildes()
        For i = 0 To buscar.Length - 1
            index = texto.IndexOf(buscar(i).Trim().QuitarTildes(), StringComparison.CurrentCultureIgnoreCase)
            If index = -1 Then Exit For
        Next
        Return index
    End Function

    ''' <summary>
    ''' Busca en texto cualquiera de las cadenas indicadas y devuelve la posición o -1 si no se ha encontrado.
    ''' Tiene preferencia que se encuentre lo buscado.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se buscará.</param>
    ''' <param name="buscar">Array de tipo cadena con lo que se quiere buscar.</param>
    ''' <returns>La posición en base cero en la cadena de cualquiera de los valores de buscar, o -1 si no se ha encontrado.</returns>
    ''' <remarks>No diferencia entre mayúsculas y minúsculas y busca sin diferenciar con tildes y se quitan los espacios extras.</remarks>
    <Extension>
    Public Function ContieneAlguno(texto As String, ParamArray buscar As String()) As Integer
        Dim index As Integer = -1
        texto = texto.QuitarTildes()
        For i = 0 To buscar.Length - 1
            index = texto.IndexOf(buscar(i).Trim().QuitarTildes(), StringComparison.CurrentCultureIgnoreCase)
            If index > -1 Then Exit For 'Return index
        Next
        Return index
    End Function

    '''<summary>
    ''' Ajusta el ancho de la cadena según el valor indicado.
    '''</summary>
    '''<remarks>No quita los espacios extras.</remarks>
    <Extension>
    Public Function AjustarAncho(cadena As String, ancho As Integer) As String
        Dim sb As New System.Text.StringBuilder(New String(" "c, ancho))
        Return (cadena & sb.ToString()).Substring(0, ancho) '.Trim()
    End Function

    ''' <summary>
    ''' Ajusta el ancho de la cadena según el valor indicado.
    ''' </summary>
    ''' <param name="cadena">El texto a ajustar.</param>
    ''' <param name="ancho">El ancho total a devolver.</param>
    ''' <param name="conSeparador">Si se muestra con el sepadador en medio o '...' al final. Si es false, se ignoran los valores de alPrincipio y alFinal.</param>
    ''' <param name="alPrincipio">Cuántos caracteres al principio, -1 indica la mitad. Al valor indicado/calculado se le restarán 2.</param>
    ''' <param name="alFinal">Cuántos caracteres al final, -1 indica la mitad. Al valor indicado/calculado se le restarán 3.</param>
    ''' <returns>La nueva cadena</returns>
    ''' <remarks>
    ''' Entre el principio y el final de la cadena se añadirá ' [-] '
    ''' esos 5 caracteres se descuentan del total a mostrar, 
    ''' por tanto si el ancho a mostrar es pequeño no tiene mucho sentido usar esta función.
    ''' </remarks>
    <Extension>
    Public Function AjustarAnchoPartido(cadena As String, ancho As Integer,
                                        conSeparador As Boolean,
                                        Optional alPrincipio As Integer = -1,
                                        Optional alFinal As Integer = -1) As String
        ' Si la cadena ya tiene los caracteres a devolver, no hacer nada
        If String.IsNullOrWhiteSpace(cadena) = False AndAlso cadena.Length = ancho Then
            Return cadena
        End If

        cadena = cadena.Trim()
        Dim longCadena = cadena.Length

        If conSeparador Then
            If alPrincipio = -1 Then
                alPrincipio = ancho \ 2
            End If
            If alFinal = -1 Then
                alFinal = ancho \ 2
            End If

            ' ajustar los anchos del final y el principio
            alPrincipio -= 2
            alFinal -= 3
            If alPrincipio > longCadena \ 2 Then
                alPrincipio = longCadena \ 2
            End If
            If alFinal > longCadena \ 2 Then
                alFinal = longCadena \ 2
            End If

            ' Solo hacerlo si hay algo en la cadena que ocupe más de el ancho - 5 (del separador)
            If longCadena > ancho - 5 Then
                cadena = cadena.Substring(0, alPrincipio) & " [-] " & cadena.Substring(longCadena - alFinal)
            End If
        Else
            If longCadena > ancho Then
                cadena = cadena.Substring(0, ancho - 3) & "..."
            End If
        End If

        Dim sb As New StringBuilder()
        sb.Append(cadena)
        sb.Append(New String(" "c, ancho))

        Return sb.ToString().Substring(0, ancho)
    End Function

    '
    ' Definir esta propiedad para que esté compartido.              (17/Oct/21)
    '

    Private _CultureES As CultureInfo

    ''' <summary>
    ''' Devuelve un objeto con el valor de la cultura en español-España (es-ES).
    ''' </summary>
    ''' <returns>Un objeto del tipo <see cref="CultureInfo"/> con la cultura para es-ES.</returns>
    Public ReadOnly Property CultureES As CultureInfo
        Get
            If _CultureES Is Nothing Then
                _CultureES = CultureInfo.CreateSpecificCulture("es-ES")
            End If
            Return _CultureES
        End Get
    End Property

    ''' <summary>
    ''' Devuelve una cadena de la fecha y hora indicada usando un formato en español.
    ''' </summary>
    ''' <param name="laFecha">Un tipo Object que representa una fecha.</param>
    ''' <param name="formato">El formato a usar.</param>
    ''' <returns>Una cadena con el formato indicado, pero en cultura es-ES.</returns>
    Public Function FechaSP(laFecha As Object, formato As String) As String
        Dim fec = laFecha.ToString().AsDateTime()
        If fec.Year = 1900 Then
            Return laFecha.ToString()
        End If
        'Dim culture = System.Globalization.CultureInfo.CreateSpecificCulture("es-ES")
        Dim sfec = fec.ToString(formato, CultureES)
        Return sfec
    End Function

    ''' <summary>
    ''' Devuelve una cadena de la fecha indicada usando un formato en español.
    ''' </summary>
    ''' <param name="laFecha">La fecha a convertir en cadena.</param>
    ''' <param name="formato">El formato a usar.</param>
    ''' <returns>Una cadena con el formato indicado, pero en cultura es-ES.</returns>
    <Extension>
    Public Function FechaSP(laFecha As Date, formato As String) As String
        'Dim culture = System.Globalization.CultureInfo.CreateSpecificCulture("es-ES")
        Dim sfec = laFecha.ToString(formato, CultureES)
        Return sfec
    End Function

    ''' <summary>
    ''' Convierte la cadena indicada en singular.
    ''' </summary>
    ''' <param name="plural"></param>
    ''' <param name="conES"></param>
    ''' <param name="conN"></param>
    ''' <param name="variasPalabras"></param>
    ''' <returns></returns>
    <Extension>
    Public Function Singular(plural As String,
                             Optional conES As Boolean = False,
                             Optional conN As Boolean = False,
                             Optional variasPalabras As Boolean = False) As String
        Dim mayusculas = plural = plural.ToLower()

        If variasPalabras Then
            Dim sbSingular As New StringBuilder()
            Dim col = Palabras(plural)
            For i = 0 To col.Count - 1
                col(i) = col(i).Trim().Singular(conES:=conES, conN:=conN)
                sbSingular.Append($"{col(i)} ")
            Next
            plural = sbSingular.ToString().TrimEnd()
        ElseIf conN Then
            plural = plural.TrimEnd("es".ToCharArray())
        ElseIf conES Then
            plural = plural.TrimEnd("s".ToCharArray())
        End If
        If mayusculas Then
            Return plural.ToLower()
        End If

        Return plural
    End Function

    ''' <summary>
    ''' Devuelve el plural del texto indicado, según el valor sea distinto de 1.
    ''' </summary>
    ''' <param name="n">El valor a tener en cuenta (será plural si es distinto de 1).</param>
    ''' <param name="singular">La palabra a pluralizar.</param>
    ''' <param name="conES">Si el plural debe finalizar con ES en lugar de con S.</param>
    ''' <param name="conN">Si el plural debe finalizar con N (queda -> quedan).</param>
    ''' <param name="variasPalabras">Si se incluyen varias palabras separadas por espacio.</param>
    ''' <returns>La cadena pluralizada o la indicada si no es plural.</returns>
    ''' <remarks>Si la palabra en singular es en mayúsculas se devuelve en mayúsculas.</remarks>
    <Extension>
    Public Function Plural(singular As String, n As Integer,
                           Optional conES As Boolean = False,
                           Optional conN As Boolean = False,
                           Optional variasPalabras As Boolean = False) As String
        Return Plural(n, singular, conES, conN, variasPalabras)
    End Function

    ''' <summary>
    ''' Devuelve el plural del texto indicado, según el valor sea distinto de 1.
    ''' </summary>
    ''' <param name="n">El valor a tener en cuenta (será plural si es distinto de 1).</param>
    ''' <param name="singular">La palabra a pluralizar.</param>
    ''' <param name="conES">Si el plural debe finalizar con ES en lugar de con S.</param>
    ''' <param name="conN">Si el plural debe finalizar con N (queda -> quedan).</param>
    ''' <param name="variasPalabras">Si se incluyen varias palabras separadas por espacio.</param>
    ''' <returns>La cadena pluralizada o la indicada si no es plural.</returns>
    ''' <remarks>Si la palabra en singular es en mayúsculas se devuelve en mayúsculas.</remarks>
    <Extension>
    Public Function Plural(n As Integer, singular As String,
                           Optional conES As Boolean = False,
                           Optional conN As Boolean = False,
                           Optional variasPalabras As Boolean = False) As String
        Dim mayusculas = singular = singular.ToUpper()

        If n <> 1 Then
            ' Poner primero si son varias palabras. v1.10.28.1 (02/sep/22 15.06)
            If variasPalabras Then
                Dim col = Palabras(singular)
                singular = ""
                For i = 0 To col.Count - 1
                    col(i) = col(i).Trim().Plural(n, conES:=conES, conN:=conN)
                    singular &= col(i) & " "
                Next
                singular = singular.TrimEnd()
            ElseIf conN Then
                singular &= "n"
            ElseIf conES Then
                singular &= "es"
            Else
                singular &= "s"
            End If

            'If conN Then
            '    singular &= "n"
            'ElseIf conES Then
            '    singular &= "es"
            'ElseIf variasPalabras Then
            '    Dim col = Palabras(singular)
            '    singular = ""
            '    For i = 0 To col.Count - 1
            '        col(i) = col(i).Trim().Plural(n, conES:=conES, conN:=conN)
            '        singular &= col(i) & " "
            '    Next
            '    singular = singular.TrimEnd()
            'Else
            '    singular &= "s"
            'End If
        End If
        If mayusculas Then
            Return singular.ToUpper
        End If
        Return singular
    End Function

    ''' <summary>
    ''' Devuelve true si el número indicado es múltiplo exacto del indicado en veces.
    ''' </summary>
    ''' <param name="numero">El número a comprobar.</param>
    ''' <param name="veces">Las veces a comprobar.</param>
    ''' <returns>True si el módulo resultante es cero.</returns>
    ''' <remarks>Por ejemplo: 126 es múltiplo exacto de 7 = sí, 125 no lo es.</remarks>
    <Extension>
    Public Function EsMultiplo(numero As Integer, veces As Integer) As Boolean
        'return numero % veces == 0
        Return (numero Mod veces) = 0
    End Function

    ''' <summary>
    ''' Devuelve true si el número indicado es cumple las veces indicadas.
    ''' </summary>
    ''' <param name="numero">El número a comprobar.</param>
    ''' <param name="veces">Las veces a comprobar.</param>
    ''' <returns>True si el módulo resultante es cero.</returns>
    ''' <remarks>Por ejemplo: 126 es múltiplo exacto de 7 = sí, 125 no lo es.</remarks>
    <Extension>
    Public Function EsVeces(numero As Integer, veces As Integer) As Boolean
        Return EsMultiplo(numero, veces)
    End Function

    ''' <summary>
    ''' Devuelve si un número es par
    ''' </summary>
    ''' <remarks>27/May/19</remarks>
    <Extension>
    Public Function IsEven(n As Integer) As Boolean
        Return EsPar(n)
    End Function

    ''' <summary>
    ''' Devuelve si un número es par
    ''' </summary>
    ''' <remarks>27/May/19</remarks>
    <Extension>
    Public Function EsPar(n As Integer) As Boolean
        'return value % 2 == 0
        Return (n Mod 2) = 0
    End Function

    ''' <summary>
    ''' Devuelve si un número es impar
    ''' </summary>
    ''' <remarks>27/May/19</remarks>
    <Extension>
    Public Function IsOdd(n As Integer) As Boolean
        Return EsImpar(n)
    End Function

    ''' <summary>
    ''' Devuelve si un número es impar
    ''' </summary>
    ''' <remarks>27/May/19</remarks>
    <Extension>
    Public Function EsImpar(n As Integer) As Boolean
        Return (n Mod 2) <> 0
    End Function

    ''' <summary>
    ''' Quitar de una cadena un texto indicado (que será el predeterminado cuando está vacío).
    ''' Por ejemplo si el texto grisáceo es Buscar... y
    ''' se empezó a escribir en medio del texto (o en cualquier parte)
    ''' BuscarL... se quitará Buscar... y se dejará L.
    ''' Antes de hacer cambios se comprueba si el texto predeterminado está al completo 
    ''' en el texto en el que se hará el cambio.
    ''' </summary>
    ''' <param name="texto">El texto en el que se hará la sustitución.</param>
    ''' <param name="predeterminado">El texto a quitar.</param>
    ''' <returns>Una cadena con el texto predeterminado quitado.</returns>
    ''' <remarks>18/Oct/2020 actualizado 24/Oct/2020</remarks>
    <Extension>
    Public Function QuitarPredeterminado(texto As String, predeterminado As String) As String
        Dim cuantos = predeterminado.Length
        Dim k = 0

        For i = 0 To predeterminado.Length - 1
            Dim j = texto.IndexOf(predeterminado(i))
            If j = -1 Then Continue For
            k += 1
        Next
        ' si k es distinto de cuantos es que no están todos lo caracteres a quitar
        If k <> cuantos Then
            Return texto
        End If

        For i = 0 To predeterminado.Length - 1
            Dim j = texto.IndexOf(predeterminado(i))
            If j = -1 Then Continue For
            If j = 0 Then
                texto = texto.Substring(j + 1)
            Else
                texto = texto.Substring(0, j) & texto.Substring(j + 1)
            End If
        Next

        Return texto
    End Function

    ''' <summary>
    ''' Devuelve true si el texto indicado contiene alguna letra del alfabeto.
    ''' Incluída la Ñ y vocales con tilde.
    ''' </summary>
    ''' <param name="texto"></param>
    ''' <returns></returns>
    ''' <remaks>14/Oct/2020</remaks>
    <Extension>
    Public Function ContieneLetras(texto As String) As Boolean
        Dim letras = "abcdefghijklmnñopqurstuvwxyzáéíóúü"
        Return texto.ToLower().IndexOfAny(letras.ToCharArray) > -1
    End Function

    ''' <summary>
    ''' Quitar las tildes de una cadena.
    ''' </summary>
    ''' <param name="s">La cadena a extender donde se buscarán las tildes.</param>
    ''' <remarks>
    ''' 03/Ago/2020
    ''' 27/Jun/2021 Usando StringBuilder en vez de concatenación.
    ''' </remarks>
    <Extension>
    Public Function QuitarTildes(s As String) As String
        Dim tildes1 = "ÁÉÍÓÚÜáéíóúü"
        Dim tildes0 = "AEIOUUaeiouu"
        'Dim res = ""
        Dim res As New StringBuilder()
        Dim j As Integer

        For i = 0 To s.Length - 1
            'Dim j = tildes1.IndexOf(s(i))
            j = tildes1.IndexOf(s(i))
            If j > -1 Then
                'res &= tildes0.Substring(j, 1)
                res.Append(tildes0.Substring(j, 1))
            Else
                'res &= s(i)
                res.Append(s(i))
            End If
        Next
        Return res.ToString()
    End Function

    '
    ' Conversiones (AsTIPO) usando cadenas en vez de controles      (28/Mar/21)
    '

    ' 28/Mar/21
    ''' <summary>
    ''' Devuelve un valor Integer de la cadena indicada.
    ''' </summary>
    ''' <param name="txt">La cadena a extender</param>
    ''' <remarks>17-oct-22: La conversión falla si el texto tiene decimales</remarks>
    <Extension>
    Public Function AsInteger(txt As String) As Integer
        Dim i As Integer = 0

        ' La conversión falla si el texto tiene decimales. (17/oct/22 12.16)
        ' Si falla: Convertir primero a double y redondearlo.
        If Integer.TryParse(txt, i) = False Then
            i = CInt(AsDoubleInt(txt))
        End If

        Return i
    End Function

    ''' <summary>
    ''' Devuelve solo la parte de la fecha de la cadena indicada.
    ''' </summary>
    ''' <param name="txt">La cadena a extender</param>
    ''' <remarks>28/Mar/21</remarks>
    <Extension>
    Public Function AsDate(txt As String) As Date
        Return AsDateTime(txt).Date
    End Function

    ''' <summary>
    ''' Devuelve un valor DateTime (fecha y hora) de la cadena indicada.
    ''' </summary>
    ''' <param name="txt">La cadena a extender</param>
    ''' <remarks>09/Oct/21</remarks>
    <Extension>
    Public Function AsDateTime(txt As String) As Date
        Dim d As New Date(1900, 1, 1, 0, 0, 0)

        If Not (String.IsNullOrWhiteSpace(txt) OrElse txt.Equals(DBNull.Value)) Then
            ' Comprobar si tiene caracteres para cambiar            (07/Oct/20)
            txt = txt.Replace(".", "/").Replace("-", "/")

            'Dim culture = Globalization.CultureInfo.CreateSpecificCulture("es-ES")
            Dim styles = Globalization.DateTimeStyles.None

            ' Usar siempre la conversión al estilo de España        (29/Mar/21)
            ' Devuelve false si no se ha convertido y la fecha es DateTime.MinValue
            ' Comprobarlo así por si falla la conversión.   (15/abr/23 06.29)
            If Date.TryParse(txt, CultureES, styles, d) = False Then
                ' Volver a intentarlo                               (01/May/21)
                ' (por si la fecha está en formato "guiri")
                If Date.TryParse(txt, d) = False Then
                    d = New Date(1900, 1, 1, 0, 0, 0)
                End If
            End If
            'If d.Year < 1900 Then
            '    ' Volver a intentarlo                               (01/May/21)
            '    ' (por si la fecha está en formato "guiri")
            '    Date.TryParse(txt, d)
            '    If d.Year < 1900 Then
            '        d = New Date(1900, 1, 1, 0, 0, 0)
            '    End If
            'End If
        Else
            ' asignar el 01/01/1900 si es un valor en blanco        (07/Jul/15)
            d = New Date(1900, 1, 1, 0, 0, 0)
        End If

        Return d
    End Function

    ''' <summary>
    ''' Devuelve el valor Integer redondeado (usando Math.Round) de la cadena indicada y tratado como Decimal.
    ''' </summary>
    ''' <param name="txt">La cadena a extender</param>
    ''' <remarks>28/Mar/21</remarks>    
    <Extension>
    Public Function AsDecimalInt(txt As String) As Integer
        Return CInt(Math.Round(txt.AsDecimal()))
    End Function

    ''' <summary>
    ''' Devuelve un entero redondeado (usando Math.Round) del decimal indicado.
    ''' </summary>
    ''' <param name="txt"></param>
    ''' <returns></returns>
    ''' <remarks>01/Abr/21</remarks>
    <Extension>
    Public Function AsDecimalInt(txt As Decimal) As Integer
        Return CInt(Math.Round(txt))
    End Function

    ''' <summary>
    ''' Devuelve un valor Decimal de la cadena indicada.
    ''' </summary>
    ''' <param name="txt">La cadena a extender</param>
    ''' <remarks>28/Mar/21</remarks>
    <Extension>
    Public Function AsDecimal(txt As String) As Decimal
        Dim d As Decimal = 0

        ' La conversión con decimales da problemas
        'Dim style = NumberStyles.Number Or NumberStyles.AllowCurrencySymbol 'Or NumberStyles.AllowDecimalPoint

        If String.IsNullOrWhiteSpace(txt) Then
            txt = "0"
        End If

        ' Si tiene símbolo de moneda, quitarlo. v1.10.13.2 (26/jul/22 21.51)
        If txt.IndexOfAny("€$".ToCharArray()) > -1 Then
            txt = txt.Replace("€", "").Replace("$", "").Trim()
        End If

        'Decimal.TryParse(txt, style, CultureES, d)
        ' No convertir en español por si se usa otro idioma para los decimales. v1.10.13.2 (26/jul/22 21.45)

        ' Si todo fue bien, devuelve el valor convertido, si el texto indicado no es un número devuelve 0
        Decimal.TryParse(txt, d)

        Return d
    End Function

    ''' <summary>
    ''' Devuelve un valor Double de la cadena indicada.
    ''' </summary>
    ''' <param name="txt">La cadena a extender</param>
    ''' <remarks>18-sep-22: v1.20.2.14</remarks>
    <Extension>
    Public Function AsDouble(txt As String) As Double
        Dim d As Double = 0

        If String.IsNullOrWhiteSpace(txt) Then
            txt = "0"
        End If

        Double.TryParse(txt, d)

        Return d
    End Function

    ''' <summary>
    ''' Devuelve un valor Double de la cadena indicada y después lo redondea.
    ''' </summary>
    ''' <param name="txt">La cadena a extender</param>
    ''' <remarks>17-oct-22</remarks>
    <Extension>
    Public Function AsDoubleInt(txt As String) As Double
        Dim d As Double = 0

        If String.IsNullOrWhiteSpace(txt) Then
            txt = "0"
        End If

        Double.TryParse(txt, d)
        d = Math.Round(d)

        Return d
    End Function

    ''' <summary>
    ''' Devuelve un valor Boolean de la cadena indicada.
    ''' </summary>
    ''' <param name="txt">La cadena a extender</param>
    ''' <remarks>28/Mar/21</remarks>
    <Extension>
    Public Function AsBoolean(txt As String) As Boolean
        If txt = "" OrElse txt = "0" Then
            Return False
        End If
        Return CBool(txt)
    End Function

    ''' <summary>
    ''' Devuelve un valor TimeSpan de la cadena indicada.
    ''' </summary>
    ''' <param name="txt">La cadena a extender</param>
    ''' <remarks>28/Mar/21</remarks>
    <Extension>
    Public Function AsTimeSpan(txt As String) As TimeSpan
        Dim c As New TimeSpan(0, 0, 0)

        If String.IsNullOrWhiteSpace(txt) Then
            Return c
        End If

        ' Solo cambiar los puntos por : si no tiene :               (21/Jun/21)
        ' ya que pueden ser milisegundos...
        If txt.Contains(".") AndAlso txt.Contains(":") = False Then
            txt = txt.Replace(".", ":")
        ElseIf txt.Contains(":") = False Then
            txt &= ":00"
        End If
        If txt = ":00" Then txt = "00:00"

        TimeSpan.TryParse(txt, c)

        Return c
    End Function

    '
    ' De las extensiones de gsEvaluarColorearCodigo
    '

    '
    ' Extensiones reemplazar si no está lo que se va a reemplazar   (04/Oct/20)
    '

    ''' <summary>
    ''' Reemplazar buscar por poner si no está poner.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">La cadena a buscar sin distinguir entre mayúsculas y minúsculas.</param>
    ''' <param name="poner">La cadena a poner si previamente no está.</param>
    ''' <returns>Una cadena con los cambios realizados.</returns>
    <Extension>
    Public Function ReplaceSiNoEstaPoner(texto As String, buscar As String, poner As String) As String

        Dim j = texto.IndexOf(poner)
        ' si está lo que se quiere poner, devolver la cadena actual sin cambios
        If j > -1 Then Return texto

        Return texto.Replace(buscar, poner)
    End Function

    ''' <summary>
    ''' Reemplazar buscar por poner si no está poner.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">La cadena a buscar usando la compración indicada.</param>
    ''' <param name="poner">La cadena a poner si previamente no está.</param>
    ''' <param name="comparar">El tipo de comparación a relizar: Ordinal o OrdinalIgnoreCase.</param>
    ''' <returns>Una cadena con los cambios realizados.</returns>
    <Extension>
    Public Function ReplaceSiNoEstaPoner(texto As String, buscar As String, poner As String,
                                         comparar As StringComparison) As String

        Dim j = texto.IndexOf(poner, comparar)
        ' si está lo que se quiere poner, devolver la cadena actual sin cambios
        If j > -1 Then Return texto

        ' esta sobrecarga está en la versión 5.0.0.0 no en la 4.0.0.0
        'Return texto.Replace(buscar, poner, comparar)
        If comparar = StringComparison.OrdinalIgnoreCase Then
            'Return texto.Replace(buscar, poner)
            Dim i As Integer
            Do
                i = texto.IndexOf(buscar, comparar)
                If i = -1 Then Exit Do
                If i > 0 Then
                    texto = poner & texto.Substring(i + buscar.Length)
                Else
                    texto = texto.Substring(0, i) & poner & texto.Substring(i + buscar.Length)
                End If
            Loop 'While i > -1
            Return texto
        Else
            Return texto.Replace(buscar, poner)
        End If
    End Function

    ''' <summary>
    ''' Reemplazar buscar por poner si no está poner.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">La cadena a buscar (palabra completa) usando la comparación indicada.</param>
    ''' <param name="poner">La cadena a poner si previamente no está.</param>
    ''' <param name="comparar">El tipo de comparación a relizar: Ordinal o OrdinalIgnoreCase.</param>
    ''' <returns>Una cadena con los cambios realizados.</returns>
    <Extension>
    Public Function ReplaceWordSiNoEstaPoner(texto As String, buscar As String, poner As String,
                                             comparar As StringComparison) As String
        Dim j = texto.IndexOf(poner, comparar)
        ' si está lo que se quiere poner, devolver la cadena actual sin cambios
        If j > -1 Then Return texto

        Return ReplaceWord(texto, buscar, poner, comparar)
    End Function

    '
    ' Extensión quitar todos los espacios
    '

    ''' <summary>
    ''' Quitar todos los espacios que tenga la cadena,
    ''' incluidos los que están entre palabras.
    ''' </summary>
    ''' <param name="texto">Cadena a la que se quitarán los espacios.</param>
    ''' <returns>Una nueva cadena con TODOS los espacios quitados.</returns>
    <Extension>
    Public Function QuitarTodosLosEspacios(texto As String) As String
        Dim col As MatchCollection = Regex.Matches(texto, "\S+")
        Dim sb As New StringBuilder
        For Each m As Match In col
            sb.Append(m.Value)
        Next

        Return sb.ToString
    End Function

    '
    ' Extensión contar palabras y saber las palabras usando Regex.
    '

    ''' <summary>
    ''' Contar las palabras de una cadena de texto usando <see cref="Regex"/>.
    ''' </summary>
    ''' <param name="texto">El texto con las palabras a contar.</param>
    ''' <returns>Un valor entero con el número de palabras</returns>
    ''' <example>
    ''' Adaptado usando una cadena en vez del Text del RichTextBox
    ''' (sería del RichTextBox para WinForms)
    ''' El código lo he adaptado de:
    ''' https://social.msdn.microsoft.com/Forums/en-US/
    '''     81e438ed-9d35-47d7-a800-1fabab0f3d52/
    '''     c-how-to-add-a-word-counter-to-a-richtextbox
    '''     ?forum=csharplanguage
    ''' </example>
    <Extension>
    Public Function CuantasPalabras(texto As String) As Integer
        Dim col As MatchCollection = Regex.Matches(texto, "[\W]+")
        Return col.Count
    End Function

    '
    ' Extensiones de cadena y cambiar a mayúsculas/minúsculas       (01/Oct/20)
    '

    Public Enum CasingValues As Integer
        ''' <summary>
        ''' No se hacen cambios
        ''' </summary>
        Normal
        ''' <summary>
        ''' Todas las letras a mayúsculas
        ''' </summary>
        Upper
        ''' <summary>
        ''' Todas las letras a minúsculas.
        ''' </summary>
        Lower
        ''' <summary>
        ''' La primera letra de cada palabra a mayúsculas.
        ''' </summary>
        Title
        ''' <summary>
        ''' La primera letra de cada palabra en mayúsculas.
        ''' Equivalente a <see cref="Title"/>.
        ''' </summary>
        FirstToUpper
        ''' <summary>
        ''' La primera letra de cada palabra en minúsculas
        ''' </summary>
        FirstToLower

    End Enum

    ''' <summary>
    ''' Cambia el texto a Upper, Lower, TitleCase/FirstToUpper o FirstToLower.
    ''' Se devuelve una nueva cadena con los cambios.
    ''' Valores posibles:
    ''' Normal
    ''' Upper
    ''' Lower
    ''' Title o FirstToLower
    ''' FirstToLower
    ''' </summary>
    ''' <param name="text">La cadena a la que se aplicará</param>
    ''' <param name="queCase">Un valor </param>
    ''' <returns>Una cadena con los cambios</returns>
    <Extension>
    Public Function CambiarCase(text As String, queCase As CasingValues) As String
        Select Case queCase
            Case CasingValues.Lower
                text = text.ToLower
            Case CasingValues.Upper
                text = text.ToUpper
            'Case CasingValues.Normal
            Case CasingValues.Title, CasingValues.FirstToUpper ' Title
                text = ToTitle(text)
            Case CasingValues.FirstToLower 'camelCase
                text = ToLowerFirst(text)
            Case Else ' Normal
                '
        End Select

        Return text
    End Function

    ''' <summary>
    ''' Devuelve una cadena en tipo Título
    ''' la primera letra de cada palabra en mayúsculas.
    ''' Usando System.Globalization.CultureInfo.CurrentCulture
    ''' que es más eficaz que
    ''' System.Threading.Thread.CurrentThread.CurrentCulture
    ''' </summary>
    <Extension>
    Public Function ToTitle(text As String) As String
        ' según la documentación usar CultureInfo.CurrentCulture es más eficaz
        ' que CurrentThread.CurrentCulture
        Dim cultureInfo = System.Globalization.CultureInfo.CurrentCulture
        Dim txtInfo = cultureInfo.TextInfo
        If text Is Nothing Then
            Return ""
        End If
        Return txtInfo.ToTitleCase(text)
    End Function

    ''' <summary>
    ''' Devuelve la cadena indicada con el primer carácter en minúsculas.
    ''' Si tiene espacios delante, pone en minúscula el primer carácter que no sea espacio.
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns></returns>
    <Extension>
    Public Function ToLowerFirstChar(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then
            Return text
        End If

        Dim sb As New StringBuilder
        Dim b = False
        For i = 0 To text.Length - 1
            If Not b AndAlso Not Char.IsWhiteSpace(text(i)) Then
                sb.Append(text(i).ToString.ToLower)
                b = True
            Else
                sb.Append(text(i))
            End If
        Next

        Return sb.ToString
    End Function

    ''' <summary>
    ''' Convierte en minúsculas el primer carácter de cada palabra en la cadena indicada.
    ''' </summary>
    <Extension>
    Public Function ToLowerFirst(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then
            Return text
        End If

        Dim col = Palabras(text)
        Dim sb As New StringBuilder
        For i = 0 To col.Count - 1
            sb.AppendFormat("{0}", col(i).ToLowerFirstChar)
        Next

        Return sb.ToString
    End Function

    ''' <summary>
    ''' Devuelve una cadena en tipo Titulo o nada si es nothing
    ''' </summary>
    ''' <remarks>25/May/19</remarks>
    <Extension>
    Public Function ToTitle(obj As Object) As String
        ' según la documentación usar CultureInfo.CurrentCulture es más eficaz
        ' que CurrentThread.CurrentCulture
        If obj Is Nothing OrElse obj.Equals(DBNull.Value) Then
            Return ""
        End If
        'Return ToTitle(obj.ToString)
        Return obj.ToString().ToTitle
    End Function

    ''' <summary>
    ''' Devuelve una cadena en minúsculas usando la cultura actual.
    ''' </summary>
    Public Function ToLower(text As String) As String
        Dim cultureInfo = System.Globalization.CultureInfo.CurrentCulture
        Dim txtInfo = cultureInfo.TextInfo
        If text Is Nothing Then
            Return ""
        End If
        Return txtInfo.ToLower(text)
    End Function

    ''' <summary>
    ''' Devuelve una cadena en minúsculas usando la cultura actual.o nada si es nulo.
    ''' </summary>
    ''' <remarks>25/May/19</remarks>
    Public Function ToLower(obj As Object) As String
        If obj Is Nothing OrElse obj.Equals(DBNull.Value) Then
            Return ""
        End If
        Return ToLower(obj.ToString)
    End Function

    ''' <summary>
    ''' Devuelve una cadena en mayúsculas usando la cultura actual.
    ''' </summary>
    Public Function ToUpper(text As String) As String
        Dim cultureInfo = System.Globalization.CultureInfo.CurrentCulture
        Dim txtInfo = cultureInfo.TextInfo
        If text Is Nothing Then
            Return ""
        End If
        Return txtInfo.ToUpper(text)
    End Function

    ''' <summary>
    ''' Devuelve una cadena en mayúsculas usando la cultura actual o nada si es nulo.
    ''' </summary>
    ''' <remarks>25/May/19</remarks>
    Public Function ToUpper(obj As Object) As String
        If obj Is Nothing OrElse obj.Equals(DBNull.Value) Then
            Return ""
        End If
        Return ToUpper(obj.ToString)
    End Function

    ''' <summary>
    ''' Devuelve una cadena o nada si es nulo
    ''' </summary>
    ''' <remarks>25/May/19</remarks>
    Public Function ToStringVacia(obj As Object) As String
        If obj Is Nothing OrElse obj.Equals(DBNull.Value) Then
            Return ""
        End If
        Return obj.ToString
    End Function

    ''' <summary>
    ''' Devuelve un espacio si es nulo o es una cadena vacía
    ''' </summary>
    ''' <remarks>25/May/19</remarks>
    Public Function ToStringUnEspacio(obj As Object) As String
        If obj Is Nothing OrElse obj.Equals(DBNull.Value) Then
            Return " "
        End If
        If obj.ToString = "" Then Return " "
        Return obj.ToString
    End Function

    ''' <summary>
    ''' Devuelve una lista con las palabras del texto indicado.
    ''' </summary>
    ''' <param name="text">La cadena de la que se extraerán las palabras.</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' En realidad no devuelve solo las palabras,
    ''' ya que cada elemento contendrá los espacios y otros símbolos que estén con esa palabra:
    ''' Si la palabra tiene espacios delante también los añade, si tiene un espacio o un símbolo detrás
    ''' también lo añade.
    ''' Si al final hay espacios en blanco, los elimina.
    ''' </remarks>
    ''' <example>    Private Sub Hola(str As String) 
    ''' Devolverá: "    Private ", "Sub ", "Hola(", "str ", "As ", "String)"
    ''' </example>
    <Extension>
    Public Function Palabras(text As String) As List(Of String)
        ' busca palabra con (o sin) espacios delante (\s*),
        ' cualquier cosa (.),
        ' una o más palabras (\w+) y
        ' cualquier cosa (.)
        Dim s = "\s*.\w+."
        Dim res = Regex.Matches(text, s)
        Dim col As New List(Of String)
        For Each m As Match In res
            col.Add(m.Value)
        Next

        Return col
    End Function

    '
    ' Versiones si se comprueban mayúsculas y minúsculas            (04/Oct/20)
    '

    ''' <summary>
    ''' Reemplaza todas las ocurrencias de 'buscar' por 'poner' en el texto,
    ''' teniendo en cuenta mayúsculas y minúsculas en la cadena a buscar.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">El valor a buscar (palabra completa) distingue mayúsculas y minúsculas.</param>
    ''' <param name="poner">El nuevo valor a poner.</param>
    ''' <returns>Una cadena con los cambios.</returns>
    <Extension>
    Public Function ReplaceWordOrdinal(texto As String, buscar As String, poner As String) As String
        Return ReplaceWord(texto, buscar, poner, StringComparison.Ordinal)
    End Function

    ''' <summary>
    ''' Reemplaza todas las ocurrencias de 'buscar' por 'poner' en el texto,
    ''' ignorando mayúsculas y minúsculas en la cadena a buscar.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">El valor a buscar (palabra completa) sin distinguir mayúsculas y minúsculas.</param>
    ''' <param name="poner">El nuevo valor a poner.</param>
    ''' <returns>Una cadena con los cambios.</returns>
    <Extension>
    Public Function ReplaceWordIgnoreCase(texto As String, buscar As String, poner As String) As String
        Return ReplaceWord(texto, buscar, poner, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' Devuelve una nueva cadena en la que todas las apariciones de oldValue
    ''' en la instancia actual se reemplazan por newValue, teniendo en cuenta
    ''' que se buscarán palabras completas.
    ''' </summary>
    ''' <param name="texto">La cadena en la que se hará la búsqueda y el reemplazo.</param>
    ''' <param name="buscar">El valor a buscar (palabra completa).</param>
    ''' <param name="poner">El nuevo valor a poner.</param>
    ''' <param name="comparar">El tipo de comparación Ordinal / OrdinalIgnoreCase.</param>
    ''' <returns>Una cadena con los cambios.</returns>
    ''' <remarks>Código convertido del original en C# de palota:
    ''' https://stackoverflow.com/a/62782791/14338047</remarks>
    <Extension>
    Public Function ReplaceWord(texto As String, buscar As String, poner As String,
                                comparar As StringComparison) As String
        Dim IsWordChar = Function(c As Char) Char.IsLetterOrDigit(c) OrElse c = "_"c

        Dim sb As StringBuilder = Nothing
        Dim p As Integer = 0, j As Integer = 0

        ' Comprueba sin distinguir mayúsculas y minúsculas          (04/Oct/20)
        'Dim ordinal = StringComparison.Ordinal
        'Dim noOrdinal = StringComparison.OrdinalIgnoreCase

        Do While j < texto.Length AndAlso __Assign(j, texto.IndexOf(buscar, j, comparar)) >= 0
            If (j = 0 OrElse Not IsWordChar(texto(j - 1))) AndAlso
                (j + buscar.Length = texto.Length OrElse Not IsWordChar(texto(j + buscar.Length))) Then

                sb = If(sb, New StringBuilder())
                sb.Append(texto, p, j - p)
                sb.Append(poner)
                j += buscar.Length
                p = j
            Else
                j += 1
            End If
        Loop

        If sb Is Nothing Then Return texto
        sb.Append(texto, p, texto.Length - p)
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Función para la equivalencia en C# de:
    ''' while (j &lt; text.Length &amp;&amp; (j = unvalor) >=0 )
    ''' </summary>
    ''' <typeparam name="T">El tipo de datos</typeparam>
    ''' <param name="target">La variable a la que se le asignará el valor de la expresión de value</param>
    ''' <param name="value">Expresión con el valor a asignar a target</param>
    ''' <returns>Devuelve el valor de value</returns>
    Private Function __Assign(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function

End Module
