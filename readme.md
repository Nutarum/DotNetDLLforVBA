# Manual para general DLL desde .NET para usar en VBA sin añadir referencias a windows

## Para lograrlo, utilizaremos la libreria UnmanagedExports

MANUAL BASADO EN LA SIGUIENTE GUIA:  
https://analystcave.com/excel-use-c-sharp-in-excel-vba/

OTROS EJEMPLOS DEL USO DE LA LIBRERIA UnmanagedExports EN:  
https://sites.google.com/site/robertgiesecke/Home/uploads/unmanagedexports#TOC-VB.Net:

### *** Todas las pruebas durante el desarrollo de este manual se han realizado en VISUAL STUDIO 2010 ultimate

#### *** ALGUNAS LIBRERIAS EXTERNAS IMPORTADAS DESDE .NET, NO FUNCIONAN EN VBA

### Generamos un nuevo proyecto:
* file-> new -> project -> Elegir lenguaje -> Class Library -> OK

### Marcamos el proyecto como COM-Visible:
* project -> properties -> application -> assembly information... -> make assembly COM-Visible (marcar)

### Instalamos el instalador de packetes NuGet (Este paso se puede saltar si ya esta instalado)
* tools -> extension manager... -> NuGet package manager (instalar)

### Instalamos la librería UnmanagedExports (Hay que instalarlo en cada proyecto)
* tools -> NuGet package manager -> Package Manager Console	(Se nos abrirá en la parte inferior la consola de NuGet)
* Ejecutar el siguiente comando: Install-Package UnmanagedExports
	
### Cambiamos la plataforma objetivo de la librería:
* build -> configuration manager -> Active solution platform -> new -> type or select the new platform -> x86 -> Ok
	* La configuración por defecto (any cpu) da problemas, en la guía en la que me he basado recomiendan cambiarla por 86
	* Las pruebas han dado error al cambiarla a x64
	
### Programamos las funciones de la librería 
* ejemplos de código más abajo  
	* Importante que el nombre de la funcion .net con la etiqueta [DllExport] coincida con el nombre de la variable de tipo function que se carga de la libreria dll en VBA (CreateTestClass en los ejemplos)

### Generamos el fichero dll
* build -> build solution
	* el fichero se encontrará en la carpeta "nombre proyecto"/"nombre proyecto"/bin/x86/debug (o release)
* Copiamos el fichero DLL en la misma carpeta donde se encuentre nuestro proyecto Access (mdb/mde)

### Añadimos el código necesario en nuestro proyecto Access
* Añadir un modulo en el proyecto Access con un código similar al ejemplo (ejemplo más abajo)

# EJEMPLOS CÓDIGO
En estos ejemplos, encontramos una clase principal, marcada con la etiqueta com visible
en esta clase podremos introducir cualquier código del lenguaje correspondiente sin problemas
y estas funciones podrán ser llamadas desde el código VBA

Después vemos una clase (UnmanagedExports) cuya única función es exportar la clase principal
a nuestro código VBA  
La etiqueta Marshal es necesaria, para que el SO haga compatible el tipo de el objeto

## CODIGO C# 
#### *** SIMPLIFICACIÓN DEL CÓDIGO ORIGINAL DE LA GUIA
```
using RGiesecke.DllExport;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
    [ComVisible(true)]
    public class Class1{
        public string Text;
        public int Numbers;
        public int GetRandomNumber(){
            Random x = new Random();
            return x.Next(100);
        }
    }
    static class UnmanagedExports{
        [DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        static Object CreateTestClass(){
            return new Class1();
        }
    }
```
## CODIGO VBA
```
Declare Function CreateTestClass Lib "ClassLibrary1.dll" () As Object 
Sub TestTheTestClass()
  Dim testClass As Object
  Set testClass = CreateTestClass() 'Creates an instance of Class1
  Debug.Print testClass.GetRandomNumber 'Executes the method
  testClass.Text = "Some text" 'Set the value of the Text property
  testClass.Numbers = 23 'Set the value of the Number property
  Debug.Print testClass.Text
  Debug.Print testClass.Numbers
End Sub
```
## CODIGO VB .NET 
```
Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.IO
Imports RGiesecke.DllExport

<ComVisible(True)>
Public Class Class1
    Public text As String
End Class

Public Module UnmanagedExports
    <DllExport()> _
    Public Function CreateTestClass() As <MarshalAs(UnmanagedType.IDispatch)> Object
        Return New Class1
    End Function
End Module
```
## CODIGO C# 
#### OTRO EJEMPLO CON HILOS 
```
using RGiesecke.DllExport;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.IO;
    [ComVisible(true)]
    public class Class1{
        private static ReaderWriterLockSlim _readWriteLock = new ReaderWriterLockSlim();
        public string Text;
        public int Numbers;
        public int GetRandomNumber(){
            Random x = new Random();            
            return x.Next(100);
        }
        public int startThreads(int numeroHilos)  {
            for (int i = 0; i < numeroHilos; i++)      {
                ThreadStart childref = new ThreadStart(CallToChildThread);
                Thread childThread = new Thread(childref);
                childThread.Start();
            }
                return 0;
        }
        public void CallToChildThread()     {
		writeWaitOnLocked("hola " + DateTime.Now,"asd.txt");
            Thread.Sleep(5000);
            writeWaitOnLocked("adios " + DateTime.Now, "asd.txt");
        }
        public void writeWaitOnLocked(String text, String path){
            // Set Status to Locked
            _readWriteLock.EnterWriteLock();
            try     {
                // Append text to the file
                using (StreamWriter sw = File.AppendText(path))    {
                    sw.WriteLine(text);
                    sw.Close();
                }
            }
            finally  {
                // Release lock
                _readWriteLock.ExitWriteLock();
            }
        }
    }
    static class UnmanagedExports{
        [DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        static Object CreateTestClass(){
            return new Class1();
        }
    }
```

## CODIGO C# 
#### OTRO EJEMPLO PARA ENVIAR PAQUETES UDP 

```
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Net;
using System.Net.Sockets;

using RGiesecke.DllExport;
using System.Runtime.InteropServices;

namespace UdpSender{
    [ComVisible(true)]
    public class Class1{         
        public string send(String ip, String port, String msg){
            Socket sock = new Socket(AddressFamily.InterNetwork, SocketType.Dgram,ProtocolType.Udp);

            IPAddress serverAddr = IPAddress.Parse(ip);

            int portInt = 0;
            if (!Int32.TryParse(port, out portInt)) {
                //si el parse a int no es posible, finalizamos, devolviendo un error
                return "Error: imposible transformar puerto (" + port + ") a numero)";
            }
            IPEndPoint endPoint = new IPEndPoint(serverAddr, portInt);
            byte[] send_buffer = Encoding.ASCII.GetBytes(msg);
            sock.SendTo(send_buffer, endPoint);
            return "";
        }     
    }

    static class UnmanagedExports {
        [DllExport]
        [return: MarshalAs(UnmanagedType.IDispatch)]
        static Object LoadUdpSender() {
            return new Class1();
        }
    }
}
```