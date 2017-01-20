using OutOfProcComServerTools.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SampleOutOfProcComObject
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid(SampleOopCom.InterfaceId)]
    public interface ISampleOopCom : IReferenceCountedObject
    {
        string HelloWorld(string name);
        int AppDomainHash();
    }
}
