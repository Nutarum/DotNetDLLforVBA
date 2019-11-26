
using RGiesecke.DllExport;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
[ComVisible(true)]
public class Class1 {

    public String helloWorld() {
        return "Hello World C#!";
    }
}
static class UnmanagedExports {
    [DllExport]
    [return: MarshalAs(UnmanagedType.IDispatch)]
    static Object CreateTestClass() {
        return new Class1();
    }
}