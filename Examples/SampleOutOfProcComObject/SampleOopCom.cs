using OutOfProcComServerTools.Attributes;
using OutOfProcComServerTools.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SampleOutOfProcComObject
{
    [ComVisible(true)]
    [Guid(SampleOopCom.ClassId)]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("SampleOopCom.SampleOopCom")]
    public class SampleOopCom : ReferenceCountedObject, ISampleOopCom
    {
        [ClassId]
        internal const string ClassId = "326BCC35-E0AA-4C8F-AA3E-AB0978EF6E70";
        internal const string InterfaceId = "5577DD6F-B4E0-45EE-A272-96354915BB9C";

        public int AppDomainHash()
        {
            return AppDomain.CurrentDomain.GetHashCode();
        }

        public string HelloWorld(string name)
        {
            return String.Format("Hello, {0}!", name);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComRegisterFunction]
        public static void Register(Type t)
        {
            try
            {
                COMHelper.RegasmRegisterLocalServer(t);
            }
            catch
            {
                throw;
            }
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            COMHelper.RegasmUnregisterLocalServer(t);
        }
    }
}
