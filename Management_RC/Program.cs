using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;


namespace Management_RC
{
    class Program
    {
        static void Main(string[] args)
        {
            Processor ps = new Processor();
            ps.Run();
            Environment.Exit(0);

        }
    }
}
