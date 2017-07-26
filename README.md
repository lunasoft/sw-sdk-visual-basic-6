
![VB6](http://findicons.com/files/icons/1803/msdn/128/ms_visual_studio.png)



Compatibilidad
-------------
* CFDI 3.3
* Visual Basic 6

Dependencias
------------

* MSXML



----------------
Instalaci&oacute;n
---------
Se descarga el proyecto sw.services y se agrega al proyecto en donde se desea implementar.



#### Aunteticaci&oacute;n #####
El servicio de Autenticación es utilizado principalmente para obtener el **token** el cual sera utilizado para poder timbrar nuestro CFDI (xml) ya emitido (sellado), para poder utilizar este servicio es necesario que cuente con un **usuario** y **contraseña** para posteriormente obtenga el token, usted puede utilizar los que estan en este ejemplo para el ambiente de **Pruebas**.


**Obtener Token
Se declaran funciones de la clase Authentication.

```vb
 Dim auth As New Authentication
 MsgBox auth.Token("http://services.test.sw.com.mx", "demo", "123456789")
```
La respuesta es en formato string.

```json
{
  "data": {
    "token": "hdgshgdhsdsd....."
  },
  "status": "success"
}
```

O en caso de error

```json
{
    "message": "SuFacturacion AU1000. Error No Controlado. No se pudo generar el token de autenticación.",
    "messageDetail": "",
    "status": "error"
}
```

#### Timbrar #####
Recibe el contenido de un **XML** ya emitido (sellado) en formato **String**,    posteriormente si la factura y el token son correctos devuelve el complemento timbre en un string (**TFD**), en caso regresa un response con detalles del error.


* Timbrar

Para poder hacer uso de las funciones de Timbrado de la DLL, debes declarar la referencia de la función en el código VB6.
```vb
 Dim stamp As New Stamp
 'host, xml a timbrar, version de timbrado, token de autenticacion
 MsgBox stamp.Stamp("http://services.test.sw.com.mx", xml, "v1", token)

```

La respuesta es un string :


```json
{
  "data": {
    "tfd": "<?xml version=\"1.0\" encoding=\"utf-8\"?><tfd:TimbreFiscalDigital xsi:schemaLocation=\"http://www.sat.gob.mx/TimbreFiscalDigital http://www.sat.gob.mx/sitio_internet/cfd/TimbreFiscalDigital/TimbreFiscalDigitalv11.xsd\" Version=\"1.1\" UUID=\"e29e4be7-e2d9-4d94-a2f8-81e30c893394\" FechaTimbrado=\"2017-05-11T17:56:00\" RfcProvCertif=\"AAA010101AAA\" SelloCFD=\"YHvkKPCGUhxHRoqk8vAnNeiHVNo5KaGYa3EBU1yMOiiTNnUASQJZxFkNbn52RUMtnepI1IAXDh7FlqCm5Vjofh3vLSJFCl8A+KUYO/GRoiYXOqwPpIhBMs9JPDXnshQzgDeL4NCd6/dSuQj3hdCVZCPgUnyYjRaFUBtqfJKTuIyP3n1o0QHq9pNvQTe+I6pumMcZoK2cWsFcgj3gZ++qO/SeV8bcWpXWGVQ43dvMCggI/z3q6sMTli6TcqoLYjS/aXmtKcPXE7Lay9uEGUNXlRaNDeGFyhtRh4ABGcFzIUuOVu1aPoq5s9wX81CaYx7hgTHFg74vNVGmxbTUwMbDSg==\" NoCertificadoSAT=\"20001000000300022323\" SelloSAT=\"bMiJXKzuMoEpOS1JKY2k+WMVEwXzhT5sNx2/WkNpp6OmoXVoahVsrBQLCCuwSbusQWIpueRRL1b8s3OoLdDDqYBKPfPIBqmwa3ZpbAQCcwv91+mMKyraDGBViXLZLhvGE7hy+tiH7PE632CjS5gSeIjXlUk3/BNKPD9tio+pSmlvWV62cPoDzJm3u7KZqNB2jWoJiYT+od6VYiibvaQ90TLT+uAkyw2jzbwOdoJZuucqfOOpO8X2vSk7NGTA5M84brTIuOlF2YLCz3LQhmzGR6WEtUUQE0LdqFvKdd+0GUeY/q6eWllv3XEIK1rw4uIM8rzQT1+D4uXslV9b3V56SA==\" xmlns:tfd=\"http://www.sat.gob.mx/TimbreFiscalDigital\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" />"
  },
  "status": "success"
}
```
O en caso de error:

```json
{
  "message": "CFDI33102 - El resultado de la digestión debe ser igual al resultado de la desencripción del sello.",
  "messageDetail": "CadenaOriginal: ||3.3|RogueOne|HNFK231|2017-05-11T16:48:22|01|20001000000300022763|201.00|MXN|1|603.28|I|PUE|06300|TME960709LR2|INMOB EDMA SA DE CV|601|AAA010101AAA|Rodolfo Carranza Ramos|G03|50211503|UT421511|1|H87|Pieza|Cigarros|200.00|200.00|200.00|002|Tasa|0.160000|32.08|232.00|003|Tasa|1.600000|371.20|002|Tasa|0.160000|32.08|003|Tasa|1.600000|371.20|403.28||",
  "status": "error"
}
```

