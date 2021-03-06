﻿using OutOfProcComServerTools.Interfaces;
using System.Runtime.InteropServices;

namespace OutOfProcComServerTools.Classes
{
    /// <summary>
    /// Reference counted object base.
    /// </summary>
    [ComVisible(false)]
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
}
