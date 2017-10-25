
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
Se descarga el proyecto sw.services y se agrega al proyecto en donde se desea implementar (Tiene 5 Clases y 1 módulo).



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
Tenemos cuatro tipos distintos de timbrado dependiendo los elementos que tengamos ser&aacute; el m&aacute; conveniente a utilizar, las opciones de las que disponemos son **StampV1, StampV2, StampV3 y StampV4**. 

Para poder hacer uso de las funciones de cancelaci&oacute; de la clase, debes importar la clase en **Project / Add Class Module / Existing / Stamp.cls**  as&iacute; como declarar la referencia de la función en el código VB6.
```vb
 Dim stamp As New Stamp
 'host, xml a timbrar, token de autenticacion
' Hacemos referencia al metódo de timbrado a utilizar StampV1, StampV2, StampV3, StampV4
 MsgBox stamp.StampV1("http://services.test.sw.com.mx", xml, token)

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

#### Cancelar #####
Tenemos cuatro tipos distintos de cancelaci&oacute;n dependiendo los elementos que tengamos ser&aacute; el m&aacute; conveniente a utilizar, las opciones de las que disponemos son Cancelaci&oacute;n por **XML**, **CSD**, **PFX** y **UUID**. 

Para poder hacer uso de las funciones de cancelaci&oacute; de la clase, debes importar la clase en **Project / Add Class Module / Existing / Cancelation.cls**  as&iacute; como declarar la referencia de la función en el código VB6.


* Cancelaci&oacute;n por CSD

```vb
 	Dim Cancelation As New Cancelation
	Dim url As String
	url = "http://services.test.sw.com.mx"
 ' host  
 ' archivo *.Cer en formato base64 
 ' archivo *.Key en formato base64
 ' password del archivo Key
 ' UUID del documento a cancelar
 ' token de autenticacion
 
 MsgBox Cancelation.CancelByCSD("http://services.test.sw.com.mx", token, b64Cer, b64Key, PasswordKey, UUID)

```

    

La respuesta es un string :


```json
{
    "data": {
        "acuse": "<?xml version=\"1.0\" encoding=\"utf-8\"?><Acuse xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" Fecha=\"2017-06-27T11:00:54.8788503\" RfcEmisor=\"LAN7008173R5\"><Folios xmlns=\"http://cancelacfd.sat.gob.mx\"><UUID>3EAEABC9-EA41-4627-9609-C6856B78E2B1</UUID><EstatusUUID>202</EstatusUUID></Folios><Signature Id=\"SelloSAT\" xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo><CanonicalizationMethod Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\" /><SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#hmac-sha512\" /><Reference URI=\"\"><Transforms><Transform Algorithm=\"http://www.w3.org/TR/1999/REC-xpath-19991116\"><XPath>not(ancestor-or-self::*[local-name()='Signature'])</XPath></Transform></Transforms><DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha512\" /><DigestValue>yoO1MKUhUcokwUgyKt5GJbcXvSzZhMKOp2pGhtuwBVrk35Y8HW8s6gJ04liSamflJFNWwUzaFOIf7KpS0SKkaw==</DigestValue></Reference></SignedInfo><SignatureValue>7ZKbUqUVSXkd9Xo9Dm4xOzrqd+j8v3NQWH8HeIPH+opnTOTGNSlVu+a2cqKKB7vmbt2ZTyfsaNsZ+d7up0zEIw==</SignatureValue><KeyInfo><KeyName>00001088888810000001</KeyName><KeyValue><RSAKeyValue><Modulus>vAr6QLmcvW6auTg7a+Ogm0veNvqJ30rD3j0iSAHxGzGVrg1d0xl0Fj5l+JX9EivD+qhkSY7pfLnJoObLpQ3GGZZOOihJVS2tbJDmnn9TW8fKUOVg+jGhcnpCHaUPq/Poj8I2OVb3g7hiaREORm6tLtzOIjkOv9INXxIpRMx54cw46D5F1+0M7ECEVO8Jg+3yoI6OvDNBH+jABsj7SutmSnL1Tov/omIlSWausdbXqykcl10BLu2XiQAc6KLnl0+Ntzxoxk+dPUSdRyR7f3Vls6yUlK/+C/4FacbR+fszT0XIaJNWkHaTOoqz76Ax9XgTv9UuT67j7rdTVzTvAN363w==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue></KeyValue></KeyInfo></Signature></Acuse>",
        "uuid": {
            "3EAEABC9-EA41-4627-9609-C6856B78E2B1": "202"
        }
    },
    "status": "success"
}
```
O en caso de error:

```json
{
    "message": "Parámetros incompletos",
    "messageDetail": "Son necesarios el .Cer y el .Key en formato B64, la contraseña, el RFC y el UUID de la factura que necesita cancelar",
    "status": "error"
}
```
* Cancelaci&oacute;n por CSD

```vb
 Dim Cancelation As New Cancelation
 'host, xml a timbrar, token de autenticacion
 MsgBox Cancelation.CancelByCSD(url, token, b64Cer, b64Key, PasswordKey, UUID)
```

 * Cancelaci&oacute;n por PFX
  
 ```vb
 Dim Cancelation As New Cancelation
 'host, xml a timbrar, token de autenticacion

 MsgBox Cancelation.CancelByPFX(url, PFX, UUID, password, RFC, Token)
```

* Cancelaci&oacute;n por XML

 ```vb
 Dim Cancelation As New Cancelation
 'host, xml a timbrar, token de autenticacion

 MsgBox Cancelation.CancelByXML(url, XML, Token)
```

* Cancelaci&oacute;n por UUID

 ```vb
 Dim Cancelation As New Cancelation
 'host, UUID a cancelar, RFC ,Token de autenticacion 
 'cargar previamente en nuestro portal de administrador de timbres sus sellos digitales

 MsgBox Cancelation.CancelByUUID(url, RFC, UUID, Token)
```

La respuesta es un string :


```json
{
    "data": {
        "acuse": "<?xml version=\"1.0\" encoding=\"utf-8\"?><Acuse xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" Fecha=\"2017-06-27T11:00:54.8788503\" RfcEmisor=\"LAN7008173R5\"><Folios xmlns=\"http://cancelacfd.sat.gob.mx\"><UUID>3EAEABC9-EA41-4627-9609-C6856B78E2B1</UUID><EstatusUUID>202</EstatusUUID></Folios><Signature Id=\"SelloSAT\" xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><SignedInfo><CanonicalizationMethod Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\" /><SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#hmac-sha512\" /><Reference URI=\"\"><Transforms><Transform Algorithm=\"http://www.w3.org/TR/1999/REC-xpath-19991116\"><XPath>not(ancestor-or-self::*[local-name()='Signature'])</XPath></Transform></Transforms><DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha512\" /><DigestValue>yoO1MKUhUcokwUgyKt5GJbcXvSzZhMKOp2pGhtuwBVrk35Y8HW8s6gJ04liSamflJFNWwUzaFOIf7KpS0SKkaw==</DigestValue></Reference></SignedInfo><SignatureValue>7ZKbUqUVSXkd9Xo9Dm4xOzrqd+j8v3NQWH8HeIPH+opnTOTGNSlVu+a2cqKKB7vmbt2ZTyfsaNsZ+d7up0zEIw==</SignatureValue><KeyInfo><KeyName>00001088888810000001</KeyName><KeyValue><RSAKeyValue><Modulus>vAr6QLmcvW6auTg7a+Ogm0veNvqJ30rD3j0iSAHxGzGVrg1d0xl0Fj5l+JX9EivD+qhkSY7pfLnJoObLpQ3GGZZOOihJVS2tbJDmnn9TW8fKUOVg+jGhcnpCHaUPq/Poj8I2OVb3g7hiaREORm6tLtzOIjkOv9INXxIpRMx54cw46D5F1+0M7ECEVO8Jg+3yoI6OvDNBH+jABsj7SutmSnL1Tov/omIlSWausdbXqykcl10BLu2XiQAc6KLnl0+Ntzxoxk+dPUSdRyR7f3Vls6yUlK/+C/4FacbR+fszT0XIaJNWkHaTOoqz76Ax9XgTv9UuT67j7rdTVzTvAN363w==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue></KeyValue></KeyInfo></Signature></Acuse>",
        "uuid": {
            "3EAEABC9-EA41-4627-9609-C6856B78E2B1": "202"
        }
    },
    "status": "success"
}
```
O en caso de error:

```json
{
    "message": "Parámetros incompletos",
    "messageDetail": "Son necesarios el .Cer y el .Key en formato B64, la contraseña, el RFC y el UUID de la factura que necesita cancelar",
    "status": "error"
}
```

### Estado de Cuenta ###
Tenemos un solo tipo de funci&oacute;n para consulta de estado de cuenta  **AccountBalance**. 

Para poder hacer uso de la funci&oacute;n de estado de cuenta de la clase, debes importar la clase en **Project / Add Class Module / Existing / AccountBalance.cls**  as&iacute; como declarar la referencia de la función en el código VB6.

 ```vb
 Dim AccountBalance As New AccountBalance
 'host ,Token de autenticacion 

 MsgBox AccountBalance.AccountBalance(url, Token)
```


La respuesta es un string :


```json
{
    "data": {
        "idSaldoCliente": "126eac70-425d-4493-87af-93505bfca746",
        "idClienteUsuario": "05f731af-4c94-4d6e-aa87-7b19a16ff891",
        "saldoTimbres": 995026340,
        "timbresUtilizados": 1895963,
        "fechaExpiracion": "0001-01-01T00:00:00",
        "unlimited": false,
        "timbresAsignados": 0
    },
    "status": "success"
}
```
O en caso de error:

```json
{
"message": "Parámetros incompletos",
    "status": "error"
}
```


### Validaci&oacute;n de XML ###
Tenemos un solo tipo de funci&oacute;n para validaci&oacute;n de XML **Validate**. 

Para poder hacer uso de la funci&oacute;n de validaci&oacute; de la clase, debes importar la clase en **Project / Add Class Module / Existing / Validate.cls**  as&iacute; como declarar la referencia de la función en el código VB6.

 ```vb
 Dim Validate As New Validate
 'host , XML a validar, Token de autenticacion 

 MsgBox Validate.Validate(url, Token, XML)
 ```
 
La respuesta es un string :


```json
{
    "status": "success",
    "detail": [
        {
            "detail": [
                {
                    "message": "OK",
                    "messageDetail": "Validacion de Estructura Correcta",
                    "type": 1
                }
            ],
            "section": "CFDI33 - Validacion de Estructura"
        },
        {
            "detail": [
                {
                    "message": "NOM150-El nodo Nomina no se puede utilizar dentro del elemento ComplementoConcepto. ",
                    "messageDetail": null,
                    "type": 0
                }
            ],
            "section": "CFDI33 - Validaciones Proveedor Comprobante ( CFDI33 ) "
        },
        {
            "detail": [
                {
                    "message": "TFD11303 - El SelloSAT del TFD no es valido.",
                    "messageDetail": null,
                    "type": 0
                }
            ],
            "section": "CFDI33 - Validaciones Proveedor Complemento tfd:TimbreFiscalDigital"
        }
    ],
    "cadenaOriginalSAT": "||1.1|6b1df4e3-dcf6-4462-9fb3-05f78e9ca298|2017-08-11T11:04:43|AAA010101AAA|dWvsbRWFkjJn/7MaJXdoDHM7UaMHC7SemffmEE0nrFHGZP8hBSDSma14n8HpLa2wGgX7DJhKn7qWAEhWNDPCllCwHz4WyqzLwY3AkUFgXj1W6MvTqT6oEVJn78y/Xrk66j5fAd8Rd1h3/Oz1IXeCHvQgwSSKSOys4mTZ/r3VbRRWH0X84lgn3OZ6B2Mn3PaCTe9QmWeiQUnCkWUm6LEHvR0rOpclXpvUWUHiFmq1xx/vChrnxUeAgvvsDc0yWn+3tNU3+HxoTMIwCXHYj46iLDaCWb1jishekJhPWgz8XW7bpobZXLvso0IkT088hewypWANJF1WnXL4ui5Nu/ducQ==|20001000000300022323||",
    "cadenaOriginalComprobante": "||3.3|-|636373549714268002858|2017-08-03T11:02:51|99|30001000000300023708|17.00|10.00|MXN|1|7.00|N|PUE|66000|AAA010101AAA|NOMBRE DE PRUEBA EMISOR SA DE CV|601|AAAA010101AAA|NOMBRE DE PRUEBA RECEPTOR SA DE CV|G03|84111505|1|ACT|Pago de nómina|17.00|17.00|10.00|1.2|O|2016-11-15|2016-11-14|2016-11-15|15|15.00|10.00|2.00|01|AUAC4601138F9|IM|0|AAAA001030HSPBBB00|0|2016-10-03|P1M13D|01|Sí|08|02|1478130|Flota|Transportista|5|04|072|8154789562|0|0|NLE|BAJS721028S88|100.000|15.00|10.00|5.00|024|ÜÜÜ|Ü|10.00|5.00|10.00|011|ÜÜÜ|Ü|10.00|001|ÜÜÜ|Ü|2.00||",
    "uuid": "6b1df4e3-dcf6-4462-9fb3-05f78e9ca298",
    "statusSat": "No Encontrado",
    "statusCodeSat": "N - 602: Comprobante no encontrado."
}
```
