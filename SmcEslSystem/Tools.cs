using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;

namespace SmcEslSystem
{
    public class Tools {
        public Tools()
        {
        }

        public static string ByteArrayToString(byte[] ba)
        {
            StringBuilder stringBuilder = new StringBuilder((int)ba.Length * 2);
            byte[] numArray = ba;
            for (int i = 0; i < (int)numArray.Length; i++)
            {
                stringBuilder.AppendFormat("{0:X2}", numArray[i]);
            }
            return stringBuilder.ToString();
        }

        public static string ConvertBinaryToHex(string strBinary)
        {
            return Convert.ToInt32(strBinary, 2).ToString("x8");
        }

        public static int ConvertHexToInt(string hex)
        {
            return int.Parse(hex, NumberStyles.HexNumber);
        }

        public static string ConvertHexToString(byte[] HexValue)
        {
            return Encoding.UTF8.GetString(HexValue);
        }

        public static string ConvertHexToString(string HexValue)
        {
            string str = "";
            while (HexValue.Length > 0)
            {
                char chr = Convert.ToChar(Convert.ToUInt32(HexValue.Substring(0, 2), 16));
                str = string.Concat(str, chr.ToString());
                HexValue = HexValue.Substring(2, HexValue.Length - 2);
            }
            return str;
        }

        public static string ConvertStringToHex(string text)
        {
            return Tools.ByteArrayToString(Encoding.UTF8.GetBytes(text));
        }

        public static byte[] iCheckSum(byte[] data)
        {
            byte[] numArray = new byte[2];
            int num = 0;
            for (int i = 0; i < (int)data.Length; i++)
            {
                num += data[i];
            }
            byte[] bytes = BitConverter.GetBytes(num);
            Array.Reverse(bytes);
            numArray[0] = bytes[(int)bytes.Length - 2];
            numArray[1] = bytes[(int)bytes.Length - 1];
            return numArray;
        }

        public static string IntToHex(int iValue, int len)
        {
            string str = null;
            if (len == 1)
            {
                str = iValue.ToString("X");
            }
            else if (len == 2)
            {
                str = iValue.ToString("X2");
            }
            else if (len == 3)
            {
                str = iValue.ToString("X3");
            }
            else if (len == 4)
            {
                str = iValue.ToString("X4");
            }
            else if (len == 5)
            {
                str = iValue.ToString("X5");
            }
            else if (len == 6)
            {
                str = iValue.ToString("X6");
            }
            return str;
        }
        private static ManualResetEvent connectDone =
        new ManualResetEvent(false);
        public void SNC_GetAP_Info()
        {

            /* try
             {

                 Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
               //  socket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.Broadcast, 1);
                 //  socket.ReceiveTimeout = 200;
                 //   socket.Connect(new IPEndPoint(IPAddress.Any, 0));
                var  serverFullAddr = new IPEndPoint(IPAddress.Parse("10.10.100.100"),48899);//设置IP，端口 //"10.10.100.100"),8899 OK
                 Console.WriteLine("aaaaaaaaaaaaaaa");
                 var sfasfs = new IPEndPoint(IPAddress.Broadcast,48899);//设置IP，端口
                 //byte[] numArray = new byte[] { 0x61 ,0x64, 0x6D, 0x69, 0x6E, 0x0D, 0x0A };
                 //byte[] numArray = new byte[] { 255, 0, 1, 1, 2};
                 // byte[] numArray =  Encoding.ASCII.GetBytes("www.usr.cn");
               // byte[] numArray = new byte[] { 0x55, 0xFD, 0xAA, 0x00, 0x03, 0x61, 0x00, 0xCC, 0x2D };
                  byte[] numArray = new byte[] { 0x77, 0x77, 0x77, 0x2E, 0x75, 0x73, 0x72, 0x2E, 0x63, 0x6E };
                 socket.ReceiveTimeout = 1000;
                 // byte[] numArray = new byte[] { 0x41, 0x54, 0x2B, 0x4E, 0x45, 0x54, 0x50, 0x0a};
                 for (var a = 0; a < numArray.Length; a++)
                 {
                     Console.WriteLine("numArray" + a + ":" + numArray[a]);
                 }


                 Console.WriteLine("aaaaaaaaaaaaaaa");

                 // Connect to the remote endpoint.  
               //  socket.BeginConnect(serverFullAddr,
                  //   new AsyncCallback(ConnectCallback), socket);
               //  connectDone.WaitOne();

                 //socket.Connect(serverFullAddr);
                  socket.Bind(serverFullAddr);

                 Console.WriteLine("BBBBBBBBBBBBB");
                 // socket.Send(numArray);
                  socket.Send(numArray);
                 byte[] numArray1 = new byte[1024];
                 Console.WriteLine("aaaaaaaaaaa");
                 int num = socket.Receive(numArray1);

                 Console.WriteLine("socket.Receive" + num);

                 byte[] numArray2 = new byte[num];

                 Console.WriteLine("numArray2" + numArray2);
                 Array.Copy(numArray1, 0, numArray2, 0, num);
                 string strss = Tools.ByteArrayToString(numArray2);
                 Console.WriteLine("strss" + strss);


             }
             catch (Exception ee)
             {
                 Console.WriteLine("exception" + ee);
             }*/
              List<AP_Information> aPInformations = new List<AP_Information>();
             byte[] numArray = new byte[] { 255, 1, 1, 2 };
            //byte[] numArray = new byte[] { 255, 0, 1, 1, 2 };
            //byte[] numArray = new byte[] { 0x77, 0x77, 0x77, 0x2E, 0x75, 0x73, 0x72, 0x2E, 0x63, 0x6E };
           // byte[] numArray = new byte[] { 0x61 ,0x64, 0x6D, 0x69, 0x6E, 0x0D, 0x0A };
            IPEndPoint pEndPoint = new IPEndPoint(IPAddress.Broadcast, 1500);
            NetworkInterface[] allNetworkInterfaces = NetworkInterface.GetAllNetworkInterfaces();
            for (int i = 0; i < (int)allNetworkInterfaces.Length; i++)
            {
                NetworkInterface networkInterface = allNetworkInterfaces[i];
                if (networkInterface.NetworkInterfaceType == NetworkInterfaceType.Ethernet && networkInterface.Supports(NetworkInterfaceComponent.IPv4))
                {
                    try
                    {
                        foreach (UnicastIPAddressInformation unicastAddress in networkInterface.GetIPProperties().UnicastAddresses)
                        {
                            if (unicastAddress.Address.AddressFamily != AddressFamily.InterNetwork)
                            {
                                continue;
                            }
                            IPEndPoint BindPoint = new IPEndPoint(unicastAddress.Address, 1500);
                            Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
                            socket.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.Broadcast, 1);
                            socket.ReceiveTimeout = 200;
                            socket.Bind(BindPoint);
                            socket.SendTo(numArray, pEndPoint);
                            byte[] numArray1 = new byte[1024];
                            do
                            {
                                try
                                {

                                    int num = socket.Receive(numArray1);

                                    byte[] numArray2 = new byte[num];
                                    Array.Copy(numArray1, 0, numArray2, 0, num);
                                    Console.WriteLine("strrss"+ numArray2);
                                    if (num == 36)
                                    {
                                        string str = Tools.ByteArrayToString(numArray2);
                                        int num1 = Tools.ConvertHexToInt(str.Substring(10, 2));
                                        int num2 = Tools.ConvertHexToInt(str.Substring(12, 2));
                                        int num3 = Tools.ConvertHexToInt(str.Substring(14, 2));
                                        int num4 = Tools.ConvertHexToInt(str.Substring(16, 2));
                                        string str1 = string.Concat(new object[] { num1, ".", num2, ".", num3, ".", num4 });
                                        string str2 = str.Substring(18, 12);
                                        string str3 = str.Substring(38, 32);
                                        AP_Information aPInformation = new AP_Information()
                                        {
                                            AP_IP = str1,
                                            AP_MAC_Address = str2,
                                            AP_Name = Tools.ConvertHexToString(str3)
                                        };
                                        aPInformations.Add(aPInformation);
                                    }
                                }
                                catch (Exception exception)
                                {
                                    Console.WriteLine("exception" + exception);
                                    break;
                                }
                            }
                            while (socket.ReceiveTimeout != 0);
                            socket.Close();
                        }
                    }
                    catch
                    {
                    }
                }
            }
            Tools.ApScanEventArgs apScanEventArg = new Tools.ApScanEventArgs()
            {
                data = aPInformations
            };
            this.onApScanEvent(this, apScanEventArg);
        }

        public static byte[] StringToByteArray(string hex)
        {
            int length = hex.Length;
            byte[] num = new byte[length / 2];
            for (int i = 0; i < length; i += 2)
            {
                num[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            }
            return num;
        }

        public event EventHandler onApScanEvent;

        public class AP_Information
        {
            public string AP_IP = "";

            public string AP_MAC_Address = "";

            public string AP_Name = "";

            public AP_Information()
            {
            }
        }

        public class ApScanEventArgs : EventArgs
        {
            public List<AP_Information> data;

            public ApScanEventArgs()
            {
            }
        }

        private static void ConnectCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the socket from the state object.  
                Socket client = (Socket)ar.AsyncState;

                // Complete the connection.  
                client.EndConnect(ar);

                Console.WriteLine("Socket connected to {0}",
                    client.RemoteEndPoint.ToString());

                // Signal that the connection has been made.  
                connectDone.Set();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
