
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
 MsgBox auth.token("http://services.test.sw.com.mx", "demo", "123456789")
```
La respuesta es en formato string.


#### Timbrar #####
Recibe el contenido de un **XML** ya emitido (sellado) en formato **String**,    posteriormente si la factura y el token son correctos devuelve el complemento timbre en un string (**TFD**), en caso regresa un response con detalles del error.


* Timbrar

Para poder hacer uso de las funciones de Timbrado de la DLL, debes declarar la referencia de la función en el código VB6.
```vb
 Dim stamp As New stamp
 'host, xml a timbrar, version de timbrado, token de autenticacion
 MsgBox stamp.stampV1("http://services.test.sw.com.mx", xml, "v1", token)

```
