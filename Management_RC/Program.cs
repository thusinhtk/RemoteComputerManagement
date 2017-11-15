using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;



namespace Management_RC
{
    class Program
    {
        static void Main(string[] args)
        {
            Processor ps = Processor.getInstance();
            ps.Run();
        }

    }
}
