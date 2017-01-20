using OutOfProcComServerTools.Attributes;
using OutOfProcComServerTools.Interfaces;
using System;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OutOfProcComServerTools.Classes
{
    public class GenericClassFactory<comType, interfaceType> : IClassFactory where comType : IReferenceCountedObject 
                                                                             where interfaceType : IReferenceCountedObject
    {
        public virtual int CreateInstance(IntPtr pUnkOuter, ref Guid riid,
            out IntPtr ppvObject)
        {
            ppvObject = IntPtr.Zero;

            if (pUnkOuter != IntPtr.Zero)
            {
                // The pUnkOuter parameter was non-NULL and the object does 
                // not support aggregation.
                Marshal.ThrowExceptionForHR(COMNative.CLASS_E_NOAGGREGATION);
            }

            //Find the ClassId attribute
            string clsId = ClassIdAttribute.GetClassIdValue(typeof(comType));

            if (riid == new Guid(clsId) ||
                riid == new Guid(COMNative.IID_IDispatch) ||
                riid == new Guid(COMNative.IID_IUnknown))
            {
                //Get the default constructor
                ConstructorInfo comConstructor = typeof(comType).GetConstructor(Type.EmptyTypes);
                comType comInstance = (comType)comConstructor.Invoke(null);

                // Create the instance of the .NET object
                ppvObject = Marshal.GetComInterfaceForObject(
                    comInstance, typeof(interfaceType));
            }
            else
            {
                // The object that ppvObject points to does not support the 
                // interface identified by riid.
                Marshal.ThrowExceptionForHR(COMNative.E_NOINTERFACE);
            }

            return 0;   // S_OK
        }

        public virtual int LockServer(bool fLock)
        {
            return 0;   // S_OK
        }
    }
}
 