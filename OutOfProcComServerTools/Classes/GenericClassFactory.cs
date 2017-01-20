using OutOfProcComServerTools.Attributes;
using OutOfProcComServerTools.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OutOfProcComServerTools.Classes
{
    public class GenericClassFactory<comType, interfaceType> : IClassFactory where comType : ReferenceCountedObject 
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

            //Find the member decorated with the ClassId attribute.
            MemberFilter clsIdFilter = new MemberFilter((MemberInfo memberInfo, object obj) =>
            {
                var attributes = memberInfo.GetCustomAttributesData();

                return (attributes.Any<CustomAttributeData>(
                        (attr) => attr.Constructor.DeclaringType == typeof(ClassIdAttribute)));
            });

            MemberInfo clsIdField = typeof(comType).FindMembers(MemberTypes.Field | MemberTypes.Property, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, clsIdFilter, null).FirstOrDefault();

            if (clsIdField == null)
                throw new NotImplementedException("The ClassId member is not implemented in class " + typeof(comType).Name + ". Please make sure this class declares a static ClassId string field or property, containing the class's GUID.");

            string clsId = (string)((dynamic)clsIdField)?.GetValue(null);

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
 