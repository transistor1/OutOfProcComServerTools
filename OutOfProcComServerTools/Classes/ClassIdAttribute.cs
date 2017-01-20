using System;
using System.Linq;
using System.Reflection;

namespace OutOfProcComServerTools.Attributes
{
    //This decorator identifies the ClassId field
    public class ClassIdAttribute : Attribute
    {
        public static MemberInfo GetDecoratedMember(Type comType)
        {
            //Find the member decorated with the ClassId attribute.
            MemberFilter clsIdFilter = new MemberFilter((MemberInfo memberInfo, object obj) =>
            {
                var attributes = memberInfo.GetCustomAttributesData();

                return (attributes.Any<CustomAttributeData>(
                        (attr) => attr.Constructor.DeclaringType == typeof(ClassIdAttribute)));
            });

            MemberInfo clsIdField = comType.FindMembers(MemberTypes.Field | MemberTypes.Property, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, clsIdFilter, null).FirstOrDefault();

            return clsIdField;
        }

        public static string GetClassIdValue(Type comType)
        {
            MemberInfo clsIdField = ClassIdAttribute.GetDecoratedMember(comType);

            if (clsIdField == null)
                throw new NotImplementedException("The ClassId member is not implemented in class " + comType.Name + ". Please make sure this class declares a static ClassId string field or property, containing the class's GUID.");

            string clsId = (string)((dynamic)clsIdField)?.GetValue(null);

            return clsId;
        }
    }
}
