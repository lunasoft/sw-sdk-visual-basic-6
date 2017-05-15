
![VB6](http://findicons.com/files/icons/1803/msdn/128/ms_visual_studio.png)

A continuación encontrara la documentación necesaria para consumir nuestra librería DLL de C++ en una aplicación de Visual Basic 6 proveido por **SmarterWeb** para Timbrado de **CFDI 3.3**

Compatibilidad
-------------
* CFDI 3.3
* Visual Basic 6

Dependencias
------------
* [SW-SDK-CPP](https://github.com/lunasoft/sw-sdk-cpp)
* [CPPREST SDK](https://github.com/Microsoft/cpprestsdk)

----------------
Instalaci&oacute;n
---------
Para poder utilizar la DLL compilada en C++ en proyectos Visual Basic 6, primero debemos de descargar los archivos DLL cpprest140_2_9.dll y sw-sdk-cpp.dll y colocarlos en la dirección raiz del proyecto Visual Basic 6.

Implementaci&oacute;n
---------
La librería contara con dos servicios principales los que son la Autenticacion y el Timbrado de CFDI.

#### Aunteticaci&oacute;n #####
El servicio de Autenticación es utilizado principalmente para obtener el **token** el cual sera utilizado para poder timbrar nuestro CFDI (xml) ya emitido (sellado), para poder utilizar este servicio es necesario que cuente con un **usuario** y **contraseña** para posteriormente obtenga el token, usted puede utilizar los que estan en este ejemplo para el ambiente de **Pruebas**.
<p align="center">
    <img src="http://developers.sw.com.mx/wp-content/uploads/2017/05/vb6-one.png">
</p>

**Obtener Token
Para poder hacer uso de las funciones de Antenticación de la DLL, debes declarar la referencia de la función en el código VB6.

```vb
    Private Declare Function AuthenticationVB Lib "sw-sdk-cpp.dll" (ByVal Url As String, ByVal User As String, ByVal Pass As String, ByVal Token As String) As Long
```
La respuesta de la petición de la DLL se recibe por referencia, por lo que es necesario crear una variable de tipo String con el nombre de la respuesta "Token" y asegurarnos de reservar el número de espacio de la cadena de la que se espera la respuesta, en este ejemplo, se reservo 1024 espacios.

```vb
    Dim Token As String
    Token = Space$(1024)
```
Despues debes mandar a llamar la función AuthenticationVB que recibe como parametros:

* Url del servicio
* Usuario para obtener.
* Contraseña.
* Token (lleno de espacios para reservar memoria)
* bPoint Dato que regresa la funcion -1 en el caso que haya encontrado un error

```vb
    bPoint = AuthenticationVB(Url, User, Pass, Token)
    MsgBox(Token)
```

#### Timbrar CFDI V1 #####
**TimbrarV1** Recibe el contenido de un **XML** ya emitido (sellado) en formato **String**,    posteriormente si la factura y el token son correctos devuelve el complemento timbre en un string (**TFD**), en caso regresa un response con detalles del error.

<p align="center">
    <img src="http://developers.sw.com.mx/wp-content/uploads/2017/05/vb6-two-design.png">
</p>
* Timbrar

Para poder hacer uso de las funciones de Timbrado de la DLL, debes declarar la referencia de la función en el código VB6.
```vb
    Private Declare Function StampByTokenVB Lib "sw-sdk-cpp.dll" (ByVal Url As String, ByVal Token As String, ByVal xml As String, ByVal tfd As String) As Long

    Private Declare Function StampVB Lib "sw-sdk-cpp.dll" (ByVal Url As String, ByVal User As String, ByVal Pass As String, ByVal xml As String, ByVal tfd As String) As Long

```
La respuesta de la petición de la DLL se recibe por referencia, por lo que es necesario crear una variable de tipo String con el nombre de la respuesta "Tfd" y asegurarnos de reservar el número de espacio de la cadena de la que se espera la respuesta, en este ejemplo, se reservo 1024 espacios.
```vb
    Dim tfd As String
    tfd = Space$(1024)
```
* Timbrar con Token
Para timbrar con Token en Visual Basic 6, se necesitan los siguientes datos:
* Url del servicio
* Token
* Xml (En UTF8)
* bPoint Dato que regresa la funcion -1 en el caso que haya encontrado un error

```vb
    bPoint = StampByTokenVB(Url, Token, xml, tfd)
    MsgBox(tfd)
```
* Timbrar Sin Token
Para timbrar sin Token (utilizando usuario y contraseña) en Visual Basic 6, se necesitan los siguientes datos:
* Url del servicio
* Usuario
* Contraseña
* Xml (En UTF8)
* bPoint Dato que regresa la funcion -1 en el caso que haya encontrado un error

```vb
    bPoint = StampVB(Url, User, Password, xml, tfd)
    MsgBox(tfd)
```
