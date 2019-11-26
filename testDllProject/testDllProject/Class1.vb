Imports System
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.IO
Imports RGiesecke.DllExport

<ComVisible(True)>
Public Class Class1
    Public Function helloWorld() As String
        Return "HELLO WORLD!"
    End Function
End Class

Public Module UnmanagedExports
    <DllExport()> _
    Public Function CreateTestClass() As <MarshalAs(UnmanagedType.IDispatch)> Object
        Return New Class1
    End Function
End Module
