﻿using OutOfProcComServerTools.Interfaces;
using System;
using System.Threading;

namespace OutOfProcComServerTools.Classes
{
    public class OutOfProcServer
    {
        #region Singleton Pattern

        private Type ClassFactoryType;
        private string ClassId;

        internal OutOfProcServer()
        {
        }

        private static OutOfProcServer _instance = new OutOfProcServer();
        public static OutOfProcServer Instance
        {
            get { return _instance; }
        }

        #endregion


        private object syncRoot = new Object(); // For thread-sync in lock
        private bool _bRunning = false; // Whether the server is running

        // The ID of the thread that runs the message loop
        private uint _nMainThreadID = 0;

        // The lock count (the number of active COM objects) in the server
        private int _nLockCnt = 0;

        // The timer to trigger GC every 5 seconds
        private Timer _gcTimer;

        /// <summary>
        /// The method is call every 5 seconds to GC the managed heap after 
        /// the COM server is started.
        /// </summary>
        /// <param name="stateInfo"></param>
        private static void GarbageCollect(object stateInfo)
        {
            GC.Collect();   // GC
        }

        private uint _cookieSimpleObj;

        /// <summary>
        /// PreMessageLoop is responsible for registering the COM class 
        /// factories for the COM classes to be exposed from the server, and 
        /// initializing the key member variables of the COM server (e.g. 
        /// _nMainThreadID and _nLockCnt).
        /// </summary>
        private void PreMessageLoop()
        {
            /////////////////////////////////////////////////////////////////
            // Register the COM class factories.
            // 

            Guid clsidSimpleObj = new Guid(this.ClassId);

            //Get the default constructor
            var classFactoryConstructor = this.ClassFactoryType.GetConstructor(Type.EmptyTypes);

            IClassFactory newClassFactory = (IClassFactory)classFactoryConstructor.Invoke(null);

            // Register the CSSimpleObject class object
            int hResult = COMNative.CoRegisterClassObject(
                ref clsidSimpleObj,                 // CLSID to be registered
                newClassFactory,   // Class factory
                CLSCTX.LOCAL_SERVER,                // Context to run
                REGCLS.MULTIPLEUSE | REGCLS.SUSPENDED,
                out _cookieSimpleObj);
            if (hResult != 0)
            {
                throw new ApplicationException(
                    "CoRegisterClassObject failed w/err 0x" + hResult.ToString("X"));
            }

            // Register other class objects 
            // ...

            // Inform the SCM about all the registered classes, and begins 
            // letting activation requests into the server process.
            hResult = COMNative.CoResumeClassObjects();
            if (hResult != 0)
            {
                // Revoke the registration of CSSimpleObject on failure
                if (_cookieSimpleObj != 0)
                {
                    COMNative.CoRevokeClassObject(_cookieSimpleObj);
                }

                // Revoke the registration of other classes
                // ...

                throw new ApplicationException(
                    "CoResumeClassObjects failed w/err 0x" + hResult.ToString("X"));
            }


            /////////////////////////////////////////////////////////////////
            // Initialize member variables.
            // 

            // Records the ID of the thread that runs the COM server so that 
            // the server knows where to post the WM_QUIT message to exit the 
            // message loop.
            _nMainThreadID = NativeMethod.GetCurrentThreadId();

            // Records the count of the active COM objects in the server. 
            // When _nLockCnt drops to zero, the server can be shut down.
            _nLockCnt = 0;

            // Start the GC timer to trigger GC every 5 seconds.
            _gcTimer = new Timer(new TimerCallback(GarbageCollect), null,
                5000, 5000);
        }

        /// <summary>
        /// RunMessageLoop runs the standard message loop. The message loop 
        /// quits when it receives the WM_QUIT message.
        /// </summary>
        private void RunMessageLoop()
        {
            MSG msg;
            while (NativeMethod.GetMessage(out msg, IntPtr.Zero, 0, 0))
            {
                NativeMethod.TranslateMessage(ref msg);
                NativeMethod.DispatchMessage(ref msg);
            }
        }

        /// <summary>
        /// PostMessageLoop is called to revoke the registration of the COM 
        /// classes exposed from the server, and perform the cleanups.
        /// </summary>
        private void PostMessageLoop()
        {
            /////////////////////////////////////////////////////////////////
            // Revoke the registration of the COM classes.
            // 

            // Revoke the registration of CSSimpleObject
            if (_cookieSimpleObj != 0)
            {
                COMNative.CoRevokeClassObject(_cookieSimpleObj);
            }

            // Revoke the registration of other classes
            // ...


            /////////////////////////////////////////////////////////////////
            // Perform the cleanup.
            // 

            // Dispose the GC timer.
            if (_gcTimer != null)
            {
                _gcTimer.Dispose();
            }

            // Wait for any threads to finish.
            Thread.Sleep(1000);
        }

        /// <summary>
        /// Convenience method, so that you don't need to provide an explicit IClassFactory implementation.
        /// </summary>
        /// <param name="comClassType">Type value of your COM class</param>
        /// <param name="comClassInterface">Type value of your COM class's interface</param>
        public void Run(Type comClassType, Type comClassInterface)
        {
            //Find the ClassId attribute
            string clsId = ClassIdAttribute.GetClassIdValue(comClassType);

            //Create a generic factory type, so user doesn't have to add
            //boilerplate code to their project
            var factoryType = typeof(GenericClassFactory<,>);
            Type[] typeArgs = { comClassType, comClassInterface };
            var concreteFactoryType = factoryType.MakeGenericType(typeArgs);

            this.Run(clsId, concreteFactoryType);
        }

        /// <summary>
        /// Run the COM server. If the server is running, the function 
        /// returns directly.
        /// </summary>
        /// <remarks>The method is thread-safe.</remarks>
        public void Run(string classId, Type classFactoryType)
        {
            lock (syncRoot) // Ensure thread-safe
            {
                this.ClassId = classId;
                this.ClassFactoryType = classFactoryType;

                // If the server is running, return directly.
                if (_bRunning)
                    return;

                // Indicate that the server is running now.
                _bRunning = true;
            }

            try
            {
                // Call PreMessageLoop to initialize the member variables 
                // and register the class factories.
                PreMessageLoop();

                // Run the message loop.
                RunMessageLoop();

                // Call PostMessageLoop to revoke the registration.
                PostMessageLoop();
            }
            finally
            {
                _bRunning = false;
            }
        }

        /// <summary>
        /// Increase the lock count
        /// </summary>
        /// <returns>The new lock count after the increment</returns>
        /// <remarks>The method is thread-safe.</remarks>
        public int Lock()
        {
            return Interlocked.Increment(ref _nLockCnt);
        }

        /// <summary>
        /// Decrease the lock count. When the lock count drops to zero, post 
        /// the WM_QUIT message to the message loop in the main thread to 
        /// shut down the COM server.
        /// </summary>
        /// <returns>The new lock count after the increment</returns>
        public int Unlock()
        {
            int nRet = Interlocked.Decrement(ref _nLockCnt);

            // If lock drops to zero, attempt to terminate the server.
            if (nRet == 0)
            {
                // Post the WM_QUIT message to the main thread
                NativeMethod.PostThreadMessage(_nMainThreadID,
                    NativeMethod.WM_QUIT, UIntPtr.Zero, IntPtr.Zero);
            }

            return nRet;
        }

        /// <summary>
        /// Get the current lock count.
        /// </summary>
        /// <returns></returns>
        public int GetLockCount()
        {
            return _nLockCnt;
        }
    }
}
