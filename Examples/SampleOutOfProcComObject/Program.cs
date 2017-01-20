using OutOfProcComServerTools.Classes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleOutOfProcComObject
{
    class Program
    {
        static void Main(string[] args)
        {
            OutOfProcServer.Instance.Run(typeof(SampleOopCom), typeof(ISampleOopCom));
        }
    }
}
