using SmcEslLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace SmcEslSystem
{
    public class EslObject
    {
        public Socket workSocket = null;
        public EslUdpTest.SmcEsl mSmcEsl = new EslUdpTest.SmcEsl(null);
    }
}
