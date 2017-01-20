# Out of Process COM Server Tools for .NET

This library aims to ease the implementation of an out-of-process COM server for a .NET COM-visible class.

## Required Knowledge

You should already be familiar with implementing an in-proc COM object.

## Usage

Your project type should be set to build an executable.  Most likely, you will need to explicity specify your project as 32-bit or 64-bit in order to get the executable to load out-of-process.

Add the NuGet package to your repository, and implement your COM-visible class.

Have your COM implement interface IReferenceCountedObject, and derive your COM class from ReferenceCountedObject.

Next, instead of explicitly specifying your CoClass's GUID, you will create a constant field containing the string representation of your class's GUID.  Decorate that field with the `[ClassId]` decorator.

At this point, your COM's code should look similar to the below:

```C#
    [Guid(ComClass1.InterfaceId)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [ComVisible(true)]
    public interface IComClass1 : IReferenceCountedObject
    {
        //Implement your interface
        ...
    }


    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid(ComClass1.ClassId)]
    [ProgId("Com1.ComClass1")]
    public partial class ComClass1 : ReferenceCountedObject, IComClass1
    {
        [ClassId]
        internal const string ClassId = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX";
        internal const string InterfaceId = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX";
        //Implement your class
        ...
    }
```

Add these COM registration functions to your COM object, so that RegAsm will insert the correct registry entries:

```C#
        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComRegisterFunction]
        public static void Register(Type t)
        {
            COMHelper.RegasmRegisterLocalServer(t);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComUnregisterFunction]
        public static void Unregister(Type t)
        {
            COMHelper.RegasmUnregisterLocalServer(t);
        }
```

Now implement a `Main` method in your COM object's project:

```C#
    static void Main(string[] args)
    {
        // Run the out-of-process COM server
        OutOfProcServer.Instance.Run(typeof(ComClass1), typeof(IComClass1));
    }
```

Then, compile and register the executable with RegAsm as you would with an in-proc COM, making sure to register with the `/codebase` option.

## Deriving from IReferenceCountedObject

If you cannot derive your COM from ReferenceCountedObject, you must have it implement IReferenceCountedObject, and add the following `Lock` and `Unlock` commands to your COM class's constructor and destructor.  Your code should look like the implementation of ReferenceCountedObject:

```C#
    public class ReferenceCountedObject : IReferenceCountedObject
    {
        public ReferenceCountedObject()
        {
            // Increment the lock count of objects in the COM server.
            OutOfProcServer.Instance.Lock();
        }

        ~ReferenceCountedObject()
        {
            // Decrement the lock count of objects in the COM server.
            OutOfProcServer.Instance.Unlock();
        }
    }
```

If you do not do this, your COM's lifetime will not be properly managed.

## Other Notes

It might be helpful to include the `Costura.Fody` NuGet package in your COM object's project, or your DLLs may have some trouble being found by the .NET runtime.

## Build Requirements

NuGet extension for Visual Studio (if not pre-installed)

Requires the NuGet Packager Visual Studio add-in.

## Attribution

This project is derived from the "boilerplate" code adapted from [Microsoft's Knowledgebase article 977996](https://support.microsoft.com/en-us/help/977996/how-to-develop-an-out-of-process-com-component-by-using-visual-c-,-visual-c-,-or-visual-basic-.net).  Code provided by Microsoft is subject to their EULA.

Thanks to [Mitch on StackOverflow](http://stackoverflow.com/a/41712885/864414) for pointing me to this solution.
