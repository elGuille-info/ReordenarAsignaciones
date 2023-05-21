# ReordenarAsignaciones

Aplicación de consola para cambiar el orden de las asignaciones.  (revisión del 21-may-2023)

Versión para .NET 6.

En NuGet puedes descargar el paquete e instalarlo como Tool de .NET.


```
Por ejemplo esta asignación:
    LaFactura.Activa = chkFacActiva.IsChecked;
Lo convertirá en:
    chkFacActiva.IsChecked = LaFactura.Activa;

Y al revés:
    chkFacActiva.IsChecked = LaFactura.Activa;
Lo convertirá en:
    LaFactura.Activa = chkFacActiva.IsChecked;
```

## Conversiones especializadas

En el código utilizo conversiones extras de unas extensiones que suelo utilizar con una clase/módulo llamado Extensiones.

> En la carpeta Extensiones está el código de Visual Basic y el convertido a C# (sin revisar).

Esas extensiones convierten una cadena de texto en un tipo específico:
```
    AsDecimal, AsDecimalInt, AsInteger, AsDouble, AsDate, AsDateTime y AsTimeSpan.

Los equivalentes en ToString son:
	AsInteger       .ToString()
	AsDecimalInt    .ToString()
	AsDecimal       .ToString("0.##")
	AsDouble        .ToString("0.##")
	AsDate          .ToString("dd/MM/yyyy")
	AsDateTime      .ToString("dd/MM/yyyy HH:mm")
	AsTimeSpan      .ToString("hh\\:mm")
                        En Visual Basic sería: .ToString("hh\:mm")
```

