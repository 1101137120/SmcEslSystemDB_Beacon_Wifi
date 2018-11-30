using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Diagnostics;
using System.Web.Script.Serialization;
using SmcEslLib;
using static EslUdpTest.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using SmcEslSystem;
using ClassLibrary3;
using ZXing;
using System.Globalization;
using System.Net.Sockets;
using System.Net;
using System.Net.NetworkInformation;

namespace SmcEslSystem
{
    public partial class Form1 : Form
    {

        private Point mouse_offset;
        bool pageOneAll;
        bool beaconAll;
        bool APStart;
        object a = DBNull.Value;
        object b = DBNull.Value;
        object c = DBNull.Value;
        object d = DBNull.Value;
        object easd = DBNull.Value;
        int q = 0;
        int countconnect = 0;
        int rowIndex;
        int rownoinsertcount = 0;
        int rownonullinsert = 1;
        int firstbuildcount = 0;
        //int[] leftmosue;
        int datagridview2no = 0;
        int dataviewcurrentbefore;
        int datagridview3no = 0;
        int picturelabel = 0;
        string clicklabeldel;
        bool resetbeacon = false;
        bool autoMateESL = false;
        bool autoNullESLMate = false;

        bool updateESLingstate = false;
        bool onsaleESLingstate = false;
        bool removeESLingstate = false;
        bool CheckESLOnly = false;
        bool testest = false;
        bool ccctrue = false;
        bool ESLStyleSave = false;
        bool ESLSaleStyleSave = false;
        bool ESLStyleDataChange= false;
        bool ESLSaleStyleDataChange = false;

        int datagridview1curr = 0;
        int datagridview2curr = 0;
        int datagridview3curr = 0;

        string deviceIPData;

        Boolean onetwo;
        Boolean doubletype;
        string ESLFromIP;
        string macaddress;
        string scancodeas;
        string selectIndex;
        string selectSize;
        string dialogtext;
        string headertextall;
        string BeaconDateS;
        string BeaconTimeS;
        string BeaconDateE;
        string editdatagirdcell;



        string BeaconTimeE;
        string openExcelAddress;
        string styleName;
        string styleSaleName;
        string nullMsg;
        private ContextMenu menu = new ContextMenu();
        private ExcelData mExcelData = new ExcelData();

        ElectronicPriceData mElectronicPriceData = new ElectronicPriceData();
        Dictionary<string, EslObject> mDictSocket = new Dictionary<string, EslObject>();

        private static System.Windows.Forms.Timer ConnectBleTimeOut = new System.Windows.Forms.Timer();
        static System.Windows.Forms.Timer DisConnectTimer = new System.Windows.Forms.Timer();
        static System.Windows.Forms.Timer BleWriteTimer = new System.Windows.Forms.Timer();//寫入電子紙，怕沒回馬需要多送
    //    static System.Windows.Forms.Timer ReadTypeTimer = new System.Windows.Forms.Timer();
        private static System.Threading.ManualResetEvent connectDone = new System.Threading.ManualResetEvent(false);

        static Timer COMPORTTimer = new Timer();
        public SerialPort port = new SerialPort();
        private List<byte> packet = new List<byte>();
        Boolean isConnect;
        delegate void Display(Byte[] buffer);// UI讀取用
        
        ASCIIEncoding ascii = new ASCIIEncoding();
        string dataTemp;
        private List<string> MacAddressList = new List<string>();
        private List<Page1> PageList = new List<Page1>();
        private List<Page> BeaconList = new List<Page>();
        private List<OldEslPage> OldEslList = new List<OldEslPage>();
        private List<string> BeaconListNow = new List<string>();
        private List<Page1> SalePageList = new List<Page1>();
        private List<string> BeaconListUpdate = new List<string>();
        private List<Page1> SalePageListUpdate = new List<Page1>();
        private  List<Page> checkESLV = new List<Page>();
        private string[] Bl;
        private List<string> checkaddress = new List<string>();
        List<string> firstbuildlistID = new List<string>();
        List<string> leftmosueESL = new List<string>();
        List<string> ESLFormat = new List<string>();
        List<string> ESL29Format = new List<string>();
        List<string> ESL42Format = new List<string>();
        List<string> ESLFormatUpdate = new List<string>();
        List<string> ESLSaleFormat = new List<string>();
        List<string> ESLSale29Format = new List<string>();
        List<string> ESLSale42Format = new List<string>();
        List<string> ESLSaleFormatUpdate = new List<string>();
        List<List<string>> ESLUpdaateFail = new List<List<string>>();
        List<string> ESLFailData = new List<string>();
        List<string> APList = new List<string>();
        List<string> autoNullESLData = new List<string>();
        List<string> OldRunAPList = new List<string>();
        List<BackPage> backESLList = new List<BackPage>();




        List<int> deldataview2no = new List<int>();
        private string ip = "192.168.1.15";
        private int udpport = 8899;

        bool scan = false;
        bool isRun = false;
        static Timer ConnectTimer = new Timer();
        static Timer CheckConnectTimer = new Timer();
        static Timer CheckBeaconTimer = new Timer();
        static Timer tmr = new Timer();
        static Timer ScanTimer = new Timer();
        static Timer CheckVTimer = new Timer();
        static Timer CheckESLLoadTimer = new Timer();
        static Timer CheckESLStateTimer = new Timer();


        delegate void UIInvoker(string data, string deviceIP);
        delegate void ReceiveDataInvoker(EventArgs e);
        delegate void APScanDataInvoker(EventArgs e);



        EslUdpTest.SmcEsl mSmcEsl;
        UserControl1 mytest;
        bool blank = false;
        bool checkV = false;
        int listcount = 0;
        int totalwritecount = 0;
        int checkconnectcount = 0;
        int totalRows = 0;

        string beaconsales = "";
        string beacondays = "";
        Boolean Runtime = false;
        Boolean down = false;
        Boolean sale = false;
        Boolean immediateUpdate= false;
        Boolean reset = false;
        Boolean saletime = false;
        Boolean checkClick = false;
        Boolean EslStyleChangeUpdate = false;
        Boolean checkESLRSSIClick = false;
        PictureBox pictureBox1 = new PictureBox();
        

        Image originalImage;
        int beacon_index = 0;
        Excel.Application excel;
        Excel.Workbook excelwb;
        //    excel.Application.Workbooks.Add(true);
        Excel.Worksheet mySheet;



        Stopwatch stopwatch = new Stopwatch();//引用stopwatch物件

        public Form1()
        {
            InitializeComponent();
            this.progressBar1.Visible = false;
            pictureBox1.BackColor = Color.White;
            pictureBox1.Size = new Size(212,104);
            pictureBox1.Location = new Point(235, 81);
            panel1.Controls.Add(pictureBox1);
            dataGridView1.Anchor = AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right | AnchorStyles.Top;
            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            //-------------------------------------------------------------------------------------
         /*   foreach (string com in SerialPort.GetPortNames())//取得所有可用的連接埠
            {
                cbbComPort.Items.Add(com);
            }
            if (cbbComPort.Items.Count != 0)
            {
                cbbComPort.SelectedIndex = 0;
            }*/
            isConnect = false;
          //  ConnectStatus.ForeColor = Color.Green;

        //    COMPORTTimer.Interval = 3500;
         //   COMPORTTimer.Tick += new EventHandler(UpCOM);
        //    COMPORTTimer.Start();
            Socket client;

          //  mSmcEsl = new SmcEsl(client);
            mytest = new UserControl1();
            //--------EXCEL仔入
            //excelinputstart();
        /*    mSmcEsl.setUdpClient(ip, udpport);
            mSmcEsl.setBleConnectTimeOut(10 * 1000);
            mSmcEsl.setWriteEslTimeOut(10 * 1000);
            mSmcEsl.onScanDeviceEven += new EventHandler(ScanDeviceEven); //掃描ble
            mSmcEsl.onConnectEslDeviceEven += new EventHandler(ConnectEslDeviceEven);
            mSmcEsl.onDisconnectEslDeviceEven += new EventHandler(DisconnectEslDeviceEven);

            mSmcEsl.onReadDeviceNameEven += new EventHandler(ReadDeviceNameEven);
            mSmcEsl.onWriteDeviceNameEven += new EventHandler(WriteDeviceNameEven);
            mSmcEsl.onSetEslTurnPageTimeEven += new EventHandler(SetEslTurnPageTimeEven);

            mSmcEsl.onWriteEslDataEven += new EventHandler(WriteEslDataEven);
            mSmcEsl.onWriteEslDataFinishEven += new EventHandler(WriteEslDataFinishEven);

            mSmcEsl.onConnectBleTimeOutEven += new EventHandler(ConnectBleTimeOut);
            mSmcEsl.onWriteEslDataTimeOutEven += new EventHandler(WriteEslDataTimeOut); //寫入ESL資料超時

            mSmcEsl.onWriteBeaconEven += new EventHandler(WriteBeaconEven);*/
            //mSmcEsl.DisConnectBleDevice();

            ConnectTimer.Tick += new EventHandler(ConnectBle);
            CheckConnectTimer.Tick += new EventHandler(CheckConnectBle);
            CheckVTimer.Tick += new EventHandler(CheckVConnectBle);
            CheckESLStateTimer.Tick+= new EventHandler(CheckESLState);

            CheckBeaconTimer.Tick += new EventHandler(BeaconCheckUpdate);
            ConnectBleTimeOut.Interval = (30 * 1000);
            ConnectBleTimeOut.Tick += new EventHandler(ConnectBle_TimeOut);
            DisConnectTimer.Tick += new EventHandler(DisConnectBle);
            DisConnectTimer.Interval = 4000;
            BleWriteTimer.Interval = (1 * 1000);
            BleWriteTimer.Tick += new EventHandler(WriteESL_TimeOut);
            //ReadTypeTimer.Interval = (1 * 1000);
          //  ReadTypeTimer.Tick += new EventHandler(ReadTypeTimer_TimeOut);


            CheckESLLoadTimer.Tick += new EventHandler(CheckESLLoad);
            ScanTimer.Tick += new EventHandler(TimerEventProcessor);
            button4.Visible = false;
            radioButton1.Visible = false;
            radioButton2.Visible = false;
            //this.Hide();
            // Application.Run(new Form2());
            //  panel4.BackColor = Color.FromArgb(100,Color.Yellow);

            //  this.pictureBox5.BackColor = Color.Transparent;
            //panel4.Parent = pictureBox4;
            // MyPropertiesGrid property = new MyPropertiesGrid();
            // propertyGrid1.SelectedObject = property;

            /* Console.WriteLine("ExcelInput_Click");
             MacAddressList.Clear();
             BeaconList.Clear();
             PageList.Clear();

             dataGridView2.Columns.Clear();
             DataTable d = new DataTable();
             dataGridView2.DataSource = d;
             DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
             dgvc.Width = 60;
             dgvc.Name = "選取";
             dgvc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
             this.dataGridView1.Columns.Insert(0, dgvc);
             dataGridView1.MouseDown += new MouseEventHandler(dataGridView1_MouseDown);


             openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
             mSmcEsl.DisConnectBleDevice();
             //if (openFileDialog1.ShowDialog() == DialogResult.OK)
             //{
                 //  img1 = ImageDecoder.DecodeFromFile(openFileDialog1.FileName);
                 //MessageBox.Show(openFileDialog1.FileName );
                 string tableName = "[工作表1$]";//在頁簽名稱後加$，再用中括號[]包起來
                 string sql = "select * from " + tableName;//SQL查詢
                 DataTable kk = mExcelData.GetExcelDataTable(@"C:\Users\abby\Desktop\test.xlsx", sql);
                 totalRows = kk.Rows.Count;
                 dataGridView2.DataSource = kk;
                 dgvc = new DataGridViewCheckBoxColumn();
                 dgvc.Width = 50;
                 dgvc.Name = "Beacon選取";
                 this.dataGridView2.Columns.Insert(3, dgvc);


                 this.dataGridView2.RowsDefaultCellStyle.BackColor = Color.Bisque;
                 this.dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;

             //}
             rownoinsertcount = dataGridView1.RowCount;
             Console.WriteLine("this.dataGridView1.ColumnCount" + this.dataGridView1.ColumnCount);
             this.dataGridView2.Columns[1].ReadOnly = true;
             this.dataGridView2.Columns[2].ReadOnly = true;
             this.dataGridView2.Columns[12].ReadOnly = true;
             this.dataGridView2.Columns[13].ReadOnly = true;
             this.dataGridView2.Columns[15].ReadOnly = true;
             this.dataGridView2.Columns[16].ReadOnly = true;
             this.dataGridView2.Columns[17].ReadOnly = true;*/
            //  tmr.Tick += timerHandler;

            /* dataGridView1.ColumnCount = 10;
             DataTable bd = new DataTable();
             // dataGridView1.MouseDown += new MouseEventHandler(dataGridView1_MouseDown);
             thisESLstate = "待機中";
             dataGridView1.Columns[0].Name = "ESLID";
             dataGridView1.Columns[1].Name = "RSSI";
             dataGridView1.Columns[2].Name = "尺寸";
             dataGridView1.Columns[3].Name = "AP";
             dataGridView1.Columns[4].Name = "動作";
             dataGridView1.Columns[5].Name = "變更時間S";
             dataGridView1.Columns[6].Name = "變更時間E";
             dataGridView1.Columns[7].Name = "貨架";
             dataGridView1.Columns[8].Name = "電壓";
             dataGridView1.Columns[9].Name = "套用樣式";*/

        }
        /*     private void timerHandler(object sender, EventArgs e)
             {
                 if (q < Bl.Length)
                 {
                     mSmcEsl.ConnectBleDevice(Bl[q]);
                     Console.WriteLine("MT WRITE");
                     richTextBox1.Text = Bl[q] + "  嘗試連線中請稍候... \r\n";

                     q++;
                 }
                 else {
                     tmr.Stop(); // Manually stop timer, or let run indefinitely
                 }

             }*/



        private void SMCEslReceiveEvent(object sender, EventArgs e)
        {
            ReceiveDataInvoker stc = new ReceiveDataInvoker(ReceiveData);
            this.BeginInvoke(stc, e);
        }
        private void AP_Scan(object sender, EventArgs e)
        {
            APScanDataInvoker stc = new APScanDataInvoker(ApScanReceiveData);
            this.BeginInvoke(stc, e);
        }

        private void ApScanReceiveData(EventArgs e)
        {
          ClearSocket();
          //  dataGridView5.Columns.Clear();

            List<AP_Information> AP = (e as ApScanEventArgs).data;
            /*    dataGridView5.Columns.Clear();
                dataGridView5.ColumnCount = 4;
                dataGridView5.Columns[0].Name = "APName";
                dataGridView5.Columns[1].Name = "IP";
                dataGridView5.Columns[2].Name = "Port";
                dataGridView5.Columns[3].Name = "State";
                dataGridView5.Rows.Add("未指定");*/
            //-----ININ
                 int activeAP=0;
                 int APListData=0;
                 bool apisnull=false;
                 List<Page> lEmp = new List<Page>();
                 if (dataGridView5.Rows.Count <2) {
                     DataTable dt = dataGridView5.DataSource as DataTable;
                     dt.Rows.Add(new object[] {"未指定" });
                 }
            //-----ININ
            /*       foreach (AP_Information mAP_Information in AP)
                   {
                   foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                       {
                       if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == mAP_Information.AP_IP)
                           {
                           apisnull = true;
                           }
                       }
                   if (!apisnull) {
                       DataTable dt = dataGridView5.DataSource as DataTable;
                       dt.Rows.Add(new object[] { mAP_Information.AP_Name, mAP_Information.AP_IP, "8899" });
                       mExcelData.dataGridViewRowCellUpdate(dataGridView5, 1, dataGridView5.Rows.Count-2, false, openExcelAddress, excel, excelwb, mySheet);
                       mExcelData.dataGridViewRowCellUpdate(dataGridView5, 2, dataGridView5.Rows.Count - 2, false, openExcelAddress, excel, excelwb, mySheet);
                       mExcelData.dataGridViewRowCellUpdate(dataGridView5, 3, dataGridView5.Rows.Count-2, false, openExcelAddress, excel, excelwb, mySheet);
                   }
                   apisnull = false;
                   }
               foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
               {

                   apisnull = false;
                   foreach (AP_Information mAP_Information in AP)
                   {
                       Console.WriteLine("mAP_Information.AP_IP"+ mAP_Information.AP_IP+ "dr5.Cells[2].Value.ToString()"+ dr5.Cells[2].Value.ToString());
                       if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mAP_Information.AP_IP)
                       {
                           dr5.Cells[4].Value = "已啟用";
                           activeAP = activeAP + 1;
                           apisnull = true;
                       }

                   }

                  if (!apisnull) {
                       dr5.Cells[4].Value = "";
                       apisnull = false;

                   }


                   APListData = APListData + 1;

               }
               label9.Text = activeAP + "/" + APListData;*/
            // mExcelData.DataGridview5Update(dataGridView5,false,openExcelAddress,excel,excelwb,mySheet);
            //----------ININ
            try {
                foreach (AP_Information mAP_Information in AP)
                {
                    //dataGridView5.Rows.Add(mAP_Information.AP_Name, mAP_Information.AP_IP,"8899","已啟用");
                    richTextBox1.AppendText("IP = " + mAP_Information.AP_IP + " Mac = " + mAP_Information.AP_MAC_Address + " Name = " + mAP_Information.AP_Name);
                    richTextBox1.AppendText("\n");
                    APList.Add(mAP_Information.AP_IP);
                    Console.WriteLine("IP = " + mAP_Information.AP_IP);
                    // ClearSocket();
                    Socket client = null;
                    string ipp = mAP_Information.AP_IP.ToString();
                    client = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);  // TCP
                                                                                                           //client = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp); // UDP
                    IPAddress ipAddress = IPAddress.Parse(ipp);
                    //IPEndPoint remoteEP = new IPEndPoint(ipAddress, 8899);
                    client.ReceiveTimeout = 200;
                    IPEndPoint remoteEP = new IPEndPoint(ipAddress, 8899);
                    client.BeginConnect(remoteEP, new AsyncCallback(ConnectCallback), client);
                  //  connectDone.WaitOne();
                    client = null;
                   // System.Threading.Thread.Sleep(100);
                    //  AP_ListBox.SelectedIndex = 0;
                    // AP_IP_Label.Text = AP_ListBox.SelectedItem.ToString();
                    // AP_ListBox.Items.Add(mAP_Information.AP_IP);
                }
              /*  System.Threading.Thread.Sleep(1000);

                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    kvp.Value.mSmcEsl.startScanBleDevice();

                }*/
            }
            catch (Exception ex) {
                Console.WriteLine("ERROR:"+ex);
            }
                
            //---------ININ
            // datagridview1curr = 2;
            //  aaa(1, false, 0);
        }

        private void ConnectCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the socket from the state object.
                Socket client = (Socket)ar.AsyncState;
                // Complete the connection.
                client.EndConnect(ar);
                Console.WriteLine("client"+ client.RemoteEndPoint.ToString());
                EslUdpTest.SmcEsl aSmcEsl = new EslUdpTest.SmcEsl(client);
                aSmcEsl.onSMCEslReceiveEvent += new EventHandler(SMCEslReceiveEvent); //全資料回傳

                EslObject mEslObject = new EslObject();
                mEslObject.workSocket = client;
                mEslObject.mSmcEsl = aSmcEsl;
                mDictSocket.Add(client.RemoteEndPoint.ToString(), mEslObject);

                bool apisnull = false;
                string[] ipaddress = client.RemoteEndPoint.ToString().Split(':');
                foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                {
                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == ipaddress[0])
                    {
                        apisnull = true;
                    }
                }
                if (!apisnull)
                {
                    DataTable dt = dataGridView5.DataSource as DataTable;
                    dt.Rows.Add(new object[] { ipaddress[0], ipaddress[0], "8899" });
                    mExcelData.dataGridViewRowCellUpdate(dataGridView5, 1, dataGridView5.Rows.Count - 2, false, openExcelAddress, excel, excelwb, mySheet);
                    mExcelData.dataGridViewRowCellUpdate(dataGridView5, 2, dataGridView5.Rows.Count - 2, false, openExcelAddress, excel, excelwb, mySheet);
                    mExcelData.dataGridViewRowCellUpdate(dataGridView5, 3, dataGridView5.Rows.Count - 2, false, openExcelAddress, excel, excelwb, mySheet);
                }
                int activeAP=0;
                int APListData = 0;
                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                {

                        Console.WriteLine("mAP_Information.AP_IP" + ipaddress[0] + "dr5.Cells[2].Value.ToString()" + dr5.Cells[2].Value.ToString());
                        if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == ipaddress[0])
                        {
                        Console.WriteLine("ININ2k6" + ipaddress[0] + "dr5.Cells[2].Value.ToString()" + dr5.Cells[2].Value.ToString());
                        dr5.Cells[4].Value = "已啟用";
                            activeAP = activeAP + 1;
                        }



                    APListData = APListData + 1;

                }
                label9.Text = activeAP + "/" + APListData;
                aSmcEsl.stopScanBleDevice();
                aSmcEsl.DisConnectBleDevice();

             //   tbMessageBox.BeginInvoke(
                //    new RichTextBoxUpdateEventHandler(UpdateRichTextBox), // the method to call back on
                    //new object[] { client.RemoteEndPoint.ToString() + "  連線成功 \n" });
                connectDone.Set();
            }
            catch (Exception e)
            {
            //    tbMessageBox.BeginInvoke(
                 //   new RichTextBoxUpdateEventHandler(UpdateRichTextBox), // the method to call back on
                   // new object[] { "AP 連線失敗，請檢查網路設定是否正確 \n" });
            }
        }

        private void ClearSocket()
        {
            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {
                kvp.Value.mSmcEsl = null;
                kvp.Value.workSocket.Close();
                kvp.Value.workSocket = null;
            }
            mDictSocket.Clear();
        }


      /*  private void UpCOM(Object myObject, EventArgs myEventArgs)
        {
            cbbComPort.Items.Clear();
            cbbComPort.Text = "";
            foreach (string com in SerialPort.GetPortNames())//取得所有可用的連接埠
            {
                cbbComPort.Items.Add(com);
            }
            if (cbbComPort.Items.Count != 0)
            {
                cbbComPort.SelectedIndex = 0;
            }

            // isConnect = false;
            ConnectStatus.ForeColor = Color.Green;

            if (port.IsOpen)
            {
                isConnect = true;
            }
            else
            {
                ConnectStatus.Text = "Connect Fail !";
                ConnectStatus.ForeColor = Color.Green;
                isConnect = false;
            }
        }*/


        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            Console.WriteLine("dataGridView1_MouseDown");
            DataGridView dgv = sender as DataGridView;
            DataGridViewRow dr;
            int col = dgv.HitTest(e.X, e.Y).ColumnIndex;
            int row = dgv.HitTest(e.X, e.Y).RowIndex;
            // string BindESL;


            //允許用戶添加行時，最後一行為未實際添加的行，因此不需考慮彈出菜單
            if (row < 0 || (dgv.AllowUserToAddRows && row == dgv.Rows.Count - 1))
            {
                return;
            }
        
            
            /*    else
                {
                    //取消所選栏位
                    for (int i = 0; i < dgv.Rows.Count; i++)
                    {
                        dgv.Rows[i].Selected = false;
                        for (int j = 0; j < dgv.Columns.Count; j++)
                        {
                            dgv.Rows[i].Cells[j].Selected = false;
                        }
                    }
                    //选中当前鼠标所在的行
                    dgv.Rows[row].Selected = true;
               // BindESL = dgv.Rows[row].Cells[12].Value.ToString();
                }*/
            if (e.Button == MouseButtons.Right)//按下右鍵
            {
                menu.Show(dataGridView1, new Point(e.X, e.Y));//顯示右鍵選單
                //建立選單
                ContextMenuStrip contextMenuStrip = new ContextMenuStrip();

                ////分隔线
                //contextMenuStrip.Items.Add(new ToolStripSeparator());
                dgv.EndEdit();
                ToolStripMenuItem tsmiRemoveCurrentRow = new ToolStripMenuItem("顯示圖片");
                tsmiRemoveCurrentRow.Click += (obj, arg) =>
                {
                    // dgv.Rows.RemoveAt(row);
                    dr = dgv.Rows[row];

                    /* Console.WriteLine("條碼:" + dr.Cells[3].Value);
                     Console.WriteLine("品名:" + dr.Cells[4].Value);
                     Console.WriteLine("品牌:" + dr.Cells[5].Value);
                     Console.WriteLine("規格:" + dr.Cells[6].Value);
                     Console.WriteLine("價格:" + dr.Cells[7].Value);
                     Console.WriteLine("特價:" + dr.Cells[8].Value);
                     Console.WriteLine("Web:"  + dr.Cells[9].Value);

                     Console.WriteLine("主要促銷:" + dr.Cells[13].Value);
                     Console.WriteLine("活動日期:" + dr.Cells[14].Value);
                     Console.WriteLine("相關文宣:" + dr.Cells[15].Value);*/


                    //字體   品名  品牌  規格  價格  特價  條碼  Qr
                    Bitmap bmp = mElectronicPriceData.setPage1("Calibri", dr.Cells[6].Value.ToString(), dr.Cells[7].Value.ToString(),
                        dr.Cells[8].Value.ToString(), dr.Cells[9].Value.ToString(), dr.Cells[10].Value.ToString(),
                        dr.Cells[5].Value.ToString(), dr.Cells[11].Value.ToString(), dr.Cells[1].Value.ToString(), headertextall, ESLFormat);

                    pictureBoxPage1.Image = bmp;


                };
                contextMenuStrip.Items.Add(tsmiRemoveCurrentRow);


                ToolStripMenuItem tsmiRemoveAll = new ToolStripMenuItem("取消");
                tsmiRemoveAll.Click += (obj, arg) =>
                {
                    // dgv.Rows.Clear();
                };
                contextMenuStrip.Items.Add(tsmiRemoveAll);

                contextMenuStrip.Show(dgv, new Point(e.X, e.Y));
            }
            /*   if (e.Button == MouseButtons.Left)//按下左鍵
               {



               }*/
        }


        private void dataGridView5_MouseDown(object sender, MouseEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
           // DataGridViewRow dr;
            int col = dgv.HitTest(e.X, e.Y).ColumnIndex;
            int row = dgv.HitTest(e.X, e.Y).RowIndex;
            // string BindESL;


            //允許用戶添加行時，最後一行為未實際添加的行，因此不需考慮彈出菜單
            if (row < 0 || (dgv.AllowUserToAddRows && row == dgv.Rows.Count - 1))
            {
                return;
            }
            if (e.Button == MouseButtons.Right)//按下右鍵
            {
                menu.Show(dataGridView5, new Point(e.X, e.Y));//顯示右鍵選單
                //建立選單
                ContextMenuStrip contextMenuStrip = new ContextMenuStrip();

                ////分隔线
                //contextMenuStrip.Items.Add(new ToolStripSeparator());
                dgv.EndEdit();

                List<int> deldataview5no = new List<int>();
                ToolStripMenuItem tsmiRemoveAll = new ToolStripMenuItem("刪除");
                tsmiRemoveAll.Click += (obj, arg) =>
                {
                    if (dgv.Rows[row].Cells[1].Value.ToString() == "未指定")
                    {
                        MessageBox.Show("預設類別無法刪除");
                    }
                    else
                    {
                       
                        dgv.Rows.RemoveAt(row);
                        // dgv.Rows.Clear();
                        deldataview5no.Add(row + 2);

                     /*   if (comboBox1.Items.Count + 2 != dataGridView5.RowCount)
                        {
                            comboBox1.Items.Clear();
                            foreach (DataGridViewRow dr in dgv.Rows)
                            {
                                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() != "")
                                    comboBox1.Items.Add(dr.Cells[2].Value.ToString());
                            }
                        }*/
                        mExcelData.dataviewdel(dataGridView4, deldataview5no, "工作表4", openExcelAddress,excel,excelwb,mySheet);
                    }
                    
                };
                contextMenuStrip.Items.Add(tsmiRemoveAll);

                contextMenuStrip.Show(dgv, new Point(e.X, e.Y));
            }
            /*   if (e.Button == MouseButtons.Left)//按下左鍵
               {



               }*/
        }



        private void DataGridView1_CellValidated(object sender, EventArgs e)
        {
            // Update the labels to reflect changes to the selection.
            MessageBox.Show("Cannot delete Starting Balance row!");
        }


        #region Button
    /*    private void Connect_COM_Button_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Connect_COM_Button_Click");
            if (port.IsOpen)
            {
                port.Close();
                ConnectStatus.Text = "Connect Fail !";
                ConnectStatus.ForeColor = Color.Green;
                isConnect = false;
            }

            if (!port.IsOpen)
            {
                try
                {
                    port.PortName = cbbComPort.Text;
                    this.port.BaudRate = 9600;
                    this.port.Parity = Parity.None;       // Parity = none
                    this.port.StopBits = StopBits.One;    // stop bits = one
                    this.port.DataBits = 8;               // data bits = 8

                    // 設定 PORT 接收事件
                    port.DataReceived += new SerialDataReceivedEventHandler(port1_DataReceived);

                    // 打開 PORT
                    port.Open();
                }
                catch (Exception ex)
                {
                    port.Close();
                    MessageBox.Show("串口出問題請重新啟動程式");
                }
            }
            if (port.IsOpen == true)
            {
                ConnectStatus.Text = "Connect OK !";
                ConnectStatus.ForeColor = Color.Green;
                isConnect = true;
            }
            else預設類別無法刪除
            {
                ConnectStatus.Text = "Connect Fail !";
                ConnectStatus.ForeColor = Color.Green;
                isConnect = false;
            }
        }*/


        private void SendData_Click(object sender, EventArgs e)
        {

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }



        /*    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {

                kvp.Value.mSmcEsl.stopScanBleDevice();
            }*/
            if (!testest)
            {
                UpdateESLDen.Text = "0";
                updateESLper.Text = "0";
                /*    if (styleName==null) {
                        if (dataGridView2.Rows.Count > 1)
                        {
                            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                            {
                                Console.WriteLine(dr.Cells[1].RowIndex + dr.Cells[1].Value.ToString());

                                if (dr.Cells[1].RowIndex==0)
                                {

                                    for (int i = 0; i < dr.Cells.Count; i++)
                                    {

                                        if (dr.Cells[i].Value!=null&&dr.Cells[i].Value.ToString() != "")
                                        {
                                            Console.WriteLine("HGEEGE");
                                            if (i == 1)
                                            {
                                                styleName = dr.Cells[1].Value.ToString();
                                            }

                                            if (i != 0 && i != 1)
                                            {

                                                ESLFormat.Add(dr.Cells[i].Value.ToString());

                                                Console.WriteLine(dr.Cells[i].Value.ToString());
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }*/



                updateESLingstate = true;

                Console.WriteLine("SendData_Click");
                // UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                rowIndex = 0;
                nullMsg = null;
                PageList.Clear();
                listcount = 0;
                string eslVState = null;

                datagridview2no = 0;
                string eslNotMateAP = null;
                string eslAPNoSetMsg = null;
                dataGridView1.ClearSelection();
                // mSmcEsl.DisConnectBleDevice();
                
                if (dataGridView1.RowCount < 2)
                {
                    MessageBox.Show("請先載入資料表");
                    dataGridView1.Enabled = true;
                    return;
                }
                if (!APStart)
                {
                    MessageBox.Show("請先連接AP");
                    dataGridView1.Enabled = true;
                    return;
                }



                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    /*   if(dr.Cells[1].Value!=null)
                       {
                           if (dr.Cells[1].Value.ToString() != "")
                           {
                               if (dr.Cells[1].Style.ForeColor == Color.Black)
                               {
                                   dr.Cells[0].Value = true;
                                   dr.Cells[0].ReadOnly = false;
                               }
                           }
                           else
                           {
                               dr.Cells[0].Value = false;
                               dr.Cells[0].ReadOnly = true;
                           }
                       }*/
                    Console.WriteLine("ForeColor" + dr.Cells[1].Style.ForeColor.Name.ToString());
                    Console.WriteLine("ForeColor" + dr.Cells[1].Style.ForeColor.Name);
                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "" && dr.Cells[1].Style.ForeColor != Color.Black && dr.Cells[1].Style.ForeColor.Name != "0")
                    {
                        if (dr.Cells[1].Value.ToString() != "")
                        {

                            dr.Selected = false;
                            int aaaa = dr.Cells[1].Value.ToString().Length;
                            Console.WriteLine("aaaa" + dr.Cells[1].Value.ToString().Length);
                            if (aaaa > 14)
                            {


                                string[] drrow = dr.Cells[1].Value.ToString().Split(',');

                                for (int bb = 0; bb < drrow.Length; bb++)
                                {
                                    /*  foreach (DataGridViewRow drnew in dataGridView1.Rows)
                                      {
                                          int drdrnewlg = drnew.Cells[12].Value.ToString().Length;
                                          Console.WriteLine("cccccccc");
                                          if (drnew.Cells[12].Value.ToString() != null) {
                                          if (drnew.Cells[12].Value.ToString().Contains(drrow[bb]))
                                          {
                                                //  Console.WriteLine("drnew.Cells[12].Value.ToString()"+ drnew.Cells[12].Value.ToString());

                                              if (drnew.Cells[12].Value.ToString().Length > 14)
                                              {
                                                      string[] drnewrow1 = drnew.Cells[12].Value.ToString().Split(',');
                                                      string[] drnewrow2 = drnew.Cells[13].Value.ToString().Split(',');
                                                      string[] drnewrow3 = drnew.Cells[14].Value.ToString().Split(',');
                                                      string[] drnewrow4 = drnew.Cells[15].Value.ToString().Split(',');
                                                      string[] drnewrow5 = drnew.Cells[16].Value.ToString().Split(',');
                                                      drnew.Cells[12].Value = DBNull.Value;
                                                      drnew.Cells[13].Value = DBNull.Value;
                                                      drnew.Cells[14].Value = DBNull.Value;
                                                      drnew.Cells[15].Value = 0;
                                                      drnew.Cells[16].Value = DBNull.Value;
                                                      for (int aa=0;aa< drnewrow1.Length;aa++) {
                                                          if (drrow[bb] == drnewrow1[aa])
                                                          {

                                                              if (dr.Cells[12].Value.ToString() == null)
                                                              {
                                                                  if (drnew.Cells[12].Value != null)
                                                                      a = drnewrow1[aa];
                                                                  if (drnew.Cells[13].Value != null)
                                                                      b = drnewrow2[aa];
                                                                  if (drnew.Cells[14].Value != null)
                                                                      c = drnewrow3[aa];
                                                                  if (drnew.Cells[15].Value != null)
                                                                      d = drnewrow4[aa];

                                                                  easd = drnewrow5[aa];
                                                              }
                                                              else
                                                              {

                                                                  if (drnew.Cells[12].Value != null)
                                                                      a = a + "," + drnewrow1[aa];
                                                                  if (drnew.Cells[13].Value != null)
                                                                      b = b + "," + drnewrow2[aa];

                                                                  if (drnew.Cells[14].Value != null)
                                                                      c = c + "," + drnewrow3[aa];

                                                                  if (drnew.Cells[15].Value != null)
                                                                      d = d + "," + drnewrow4[aa];


                                                                  easd = easd + "," + drnewrow5[aa];
                                                              }
                                                          }
                                                          else {
                                                              drnew.Cells[12].Value = drnew.Cells[12].Value+ drnewrow1[aa];
                                                              drnew.Cells[13].Value = drnew.Cells[13].Value+ drnewrow2[aa];
                                                              drnew.Cells[14].Value = drnew.Cells[14].Value+ drnewrow3[aa];
                                                              drnew.Cells[15].Value = drnew.Cells[15].Value+ drnewrow4[aa];
                                                              drnew.Cells[16].Value = drnew.Cells[16].Value+ drnewrow5[aa];

                                                          }
                                                      }

                                                  }
                                              else
                                              {
                                                  if (dr.Cells[12].Value.ToString() == null)
                                                  {
                                                      if (drnew.Cells[12].Value != null)
                                                          a = drnew.Cells[12].Value;
                                                      if (drnew.Cells[13].Value != null)
                                                          b = drnew.Cells[13].Value;
                                                      if (drnew.Cells[14].Value != null)
                                                          c = drnew.Cells[14].Value;
                                                      if (drnew.Cells[15].Value != null)
                                                          d = drnew.Cells[15].Value;

                                                      easd = drnew.Cells[16].Value;
                                                  }
                                                  else
                                                  {

                                                      if (drnew.Cells[12].Value != null)
                                                          a = a+"," + drnew.Cells[12].Value;
                                                      if (drnew.Cells[13].Value != null)
                                                          b = b+"," + drnew.Cells[13].Value;

                                                          if (drnew.Cells[14].Value != null)
                                                          c = c+"," + drnew.Cells[14].Value;

                                                          if (drnew.Cells[15].Value != null)
                                                          d = d+"," + drnew.Cells[15].Value;


                                                          easd = easd+"," + drnew.Cells[16].Value;
                                                  }

                                                  drnew.Cells[12].Value = DBNull.Value;
                                                  drnew.Cells[13].Value = DBNull.Value;
                                                  drnew.Cells[14].Value = DBNull.Value;
                                                  drnew.Cells[15].Value = 0;
                                                  drnew.Cells[16].Value = DBNull.Value;
                                              }
                                              // drnew.Cells[12].Value = String.Empty;
                                              //  drnew.Cells[13].Value = String.Empty;
                                          }
                                      }

                                      }*/
                                    Page1 mPageC = new Page1();


                                   
                                    Console.WriteLine("MT WRITE QQW" + drrow[bb]);
                                    UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                    mPageC.no = (dr.Index + 1).ToString();
                                    mPageC.BleAddress = drrow[bb];
                                    mPageC.barcode = dr.Cells[5].Value.ToString();
                                    mPageC.product_name = dr.Cells[6].Value.ToString();
                                    mPageC.Brand = dr.Cells[7].Value.ToString();
                                    mPageC.specification = dr.Cells[8].Value.ToString();
                                    mPageC.price = dr.Cells[9].Value.ToString();
                                    mPageC.Special_offer = dr.Cells[10].Value.ToString();
                                    mPageC.Web = dr.Cells[11].Value.ToString();
                                    mPageC.onsale = dr.Cells[15].Value.ToString();
                                    mPageC.TimerConnect = new System.Windows.Forms.Timer();
                                    mPageC.TimerConnect.Interval = (30 * 1000);
                                    mPageC.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                    mPageC.TimerSeconds = new Stopwatch();
                                    mPageC.actionName = "Bind";
                                    if (mPageC.onsale == "V")
                                        mPageC.ProductStyle = styleSaleName;
                                    else
                                        mPageC.ProductStyle = styleName;

                                    mPageC.HeadertextALL = headertextall;
                                    mPageC.usingAddress = drrow[bb];
                                    foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                    {
                                        if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == drrow[bb])
                                        {
                                            if (drAP.Cells[8].Value.ToString() == "")
                                            {
                                                if (eslNotMateAP == null)
                                                    eslNotMateAP = drrow[bb];
                                                else
                                                    eslNotMateAP = eslNotMateAP + "," + drrow[bb];
                                                // MessageBox.Show("請先配對ESL IP");
                                                //break;
                                            }
                                            if (drAP.Cells[5].Value.ToString() != "" && Convert.ToDouble(drAP.Cells[5].Value) < 2.85)
                                            {
                                                if (eslVState == null)
                                                    eslVState = drrow[bb];
                                                else
                                                    eslVState = eslVState + "," + drrow[bb];
                                                // MessageBox.Show("請先配對ESL IP");
                                                //break;
                                            }

                                            mPageC.APLink = drAP.Cells[8].Value.ToString();
                                            break;
                                        }
                                    }
                                    foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                    {
                                        if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPageC.APLink)
                                        {
                                            if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                            {
                                                PageList.Add(mPageC);
                                            }
                                            else
                                            {
                                                if (eslAPNoSetMsg == null)
                                                    eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                                if (eslAPNoSetMsg != null)
                                                    eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                //  MessageBox.Show(" " + ((DataRowView)(dr.DataBoundItem)).Row.ItemArray[0] + " 被選取了！");
                                //字體   品名  品牌  規格  價格  特價  條碼  Qr
                                Console.WriteLine("AAAAAAAAAA");
                                Page1 mPage = new Page1();
                                if (dr.Cells[1].Value.ToString() == "")
                                {
                                    mPage.no = (rownoinsertcount + rownonullinsert).ToString();
                                    rownonullinsert++;
                                }
                                else
                                {
                                    mPage.no = (dr.Index + 1).ToString();

                                }
                                UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                mPage.BleAddress = dr.Cells[1].Value.ToString();
                                mPage.barcode = dr.Cells[5].Value.ToString();
                                mPage.product_name = dr.Cells[6].Value.ToString();
                                mPage.Brand = dr.Cells[7].Value.ToString();
                                mPage.specification = dr.Cells[8].Value.ToString();
                                mPage.price = dr.Cells[9].Value.ToString();
                                mPage.Special_offer = dr.Cells[10].Value.ToString();
                                mPage.Web = dr.Cells[11].Value.ToString();
                                mPage.onsale = dr.Cells[15].Value.ToString();
                                mPage.TimerConnect = new System.Windows.Forms.Timer();
                                mPage.TimerConnect.Interval = (30 * 1000);
                                mPage.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                mPage.TimerSeconds = new Stopwatch();
                                mPage.actionName = "Bind";
                                if (mPage.onsale == "V")
                                    mPage.ProductStyle = styleSaleName;
                                else
                                    mPage.ProductStyle = styleName;
                                mPage.HeadertextALL = headertextall;
                                mPage.usingAddress = dr.Cells[1].Value.ToString();
                                foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                {
                                    if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == dr.Cells[1].Value.ToString())
                                    {
                                        if (drAP.Cells[8].Value.ToString() == "")
                                        {
                                            if (eslNotMateAP == null)
                                                eslNotMateAP = mPage.usingAddress;
                                            else
                                                eslNotMateAP = eslNotMateAP + "," + mPage.usingAddress;
                                            // MessageBox.Show("請先配對ESL IP");
                                            //break;
                                        }
                                        if (drAP.Cells[5].Value.ToString() != "" && Convert.ToDouble(drAP.Cells[5].Value) < 2.85)
                                        {
                                            if (eslVState == null)
                                                eslVState = mPage.usingAddress;
                                            else
                                                eslVState = eslVState + "," + mPage.usingAddress;
                                            // MessageBox.Show("請先配對ESL IP");
                                            //break;
                                        }
                                        mPage.APLink = drAP.Cells[8].Value.ToString();


                                        break;
                                    }
                                }
                                dr.Cells[16].Value = "V";
                                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                {
                                    if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPage.APLink)
                                    {
                                        if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                        {
                                            PageList.Add(mPage);
                                        }
                                        else
                                        {
                                            if (eslAPNoSetMsg == null)
                                                eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                            if (eslAPNoSetMsg != null)
                                                eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                        }
                                    }
                                }

                                Console.WriteLine("BBBBBBBBBB");
                            }


                            /*    foreach (DataGridViewRow drnew in dataGridView1.Rows)
                                {
                                    Console.WriteLine("cccccccc");
                                    if (dr.Cells[2].Value.Equals(drnew.Cells[12].Value))
                                    {

                                        if (drnew.Cells[12].Value != null)
                                            a = drnew.Cells[12].Value;
                                        if (drnew.Cells[13].Value != null)
                                            b = drnew.Cells[13].Value;
                                        if (drnew.Cells[14].Value != null)
                                            c = drnew.Cells[14].Value;
                                        if (drnew.Cells[15].Value != null)
                                            d = drnew.Cells[15].Value;
                                        easd = drnew.Cells[16].Value;
                                        drnew.Cells[12].Value = DBNull.Value;
                                        drnew.Cells[13].Value = DBNull.Value;
                                        drnew.Cells[14].Value = DBNull.Value;
                                        //drnew.Cells[15].Value = DBNull.Value;
                                        drnew.Cells[16].Value = "X";


                                        // drnew.Cells[12].Value = String.Empty;
                                        //  drnew.Cells[13].Value = String.Empty;
                                    }

                                }
                                /* if (dr.Cells[12].Value.ToString().Length > 1)
                                 {
                                     dr.Cells[12].Value = dr.Cells[12].Value + "," + dr.Cells[2].Value;
                                 }
                                 else {
                                     dr.Cells[12].Value = dr.Cells[2].Value;
                                 }
                                 if (dr.Cells[13].Value.ToString().Length > 1)
                                 {
                                     dr.Cells[13].Value = dr.Cells[13].Value.ToString() + b;
                                 }
                                 else
                                 {
                                     dr.Cells[13].Value = b;
                                 }
                                 if (dr.Cells[14].Value.ToString().Length > 1)
                                 {
                                     dr.Cells[14].Value = dr.Cells[14].Value.ToString() + c;
                                 }
                                 else
                                 {
                                     dr.Cells[14].Value = c;
                                 }
                                 if (dr.Cells[15].Value.ToString().Length > 1)
                                 {
                                     dr.Cells[15].Value = dr.Cells[15].Value.ToString() + d;
                                 }
                                 else
                                 {
                                     dr.Cells[15].Value = d;
                                 }
                                 if (dr.Cells[16].Value.ToString().Length > 1)
                                 {
                                     dr.Cells[16].Value = dr.Cells[16].Value.ToString() + easd;
                                 }
                                 else
                                 {
                                     dr.Cells[16].Value = easd;
                                 }
                                dr.Cells[12].Value = dr.Cells[2].Value;
                                dr.Cells[13].Value = b;
                                dr.Cells[14].Value = c;
                                dr.Cells[15].Value = d;
                                //dr.Cells[16].Value = easd;
                                dr.Cells[16].Value = "V";
                                //dr.Cells[17].Value = DateTime.Now.ToString();
                                foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                                {
                                    if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString() == dr.Cells[12].Value.ToString())
                                    {
                                        dr4.Cells[6].Value = dr.Cells[6].Value;
                                    }
                                }*/


                            //break;
                        }
                        else
                        {
                            if (nullMsg == null)
                            {
                                nullMsg = dr.Cells[6].Value.ToString();
                            }
                            else
                            {
                                nullMsg = nullMsg + "," + dr.Cells[6].Value.ToString();
                            }
                        }
                    }
                }



                if (nullMsg != null)
                {
                    //MessageBox.Show("勾選"+ nullMsg + "未綁定電子標籤");
                    PageList.Clear();
                    dataGridView1.Enabled = true;
                    nullMsg = null;
                    datagridview1curr = 2;
                    aaa(1, false, 0);
                    DialogResult dialogResult = MessageBox.Show("勾選" + nullMsg + "未綁定電子標籤" + "\r\n" + "是否繼續綁定", "未綁定", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }


                if (eslNotMateAP != null)
                {
                    //MessageBox.Show("勾選"+ nullMsg + "未綁定電子標籤");
                    PageList.Clear();
                    dataGridView1.Enabled = true;
                    nullMsg = null;
                    datagridview1curr = 2;
                    aaa(1, false, 0);
                    DialogResult dialogResult = MessageBox.Show(eslNotMateAP + "未配對AP請自動配對" + "\r\n" + "是否繼續執行", "未配對ESL", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }

                if (eslVState != null)
                {
                    //MessageBox.Show("勾選"+ nullMsg + "未綁定電子標籤");
                    PageList.Clear();
                    dataGridView1.Enabled = true;
                    nullMsg = null;
                    datagridview1curr = 2;
                    aaa(1, false, 0);
                    DialogResult dialogResult = MessageBox.Show(eslVState + "電壓未達2.85V" + "\r\n" + "是否繼續執行", "電壓", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }



                Console.WriteLine("WHYWHYYYYYYYYYYYYYYYYYYY" + PageList.Count);


                if (eslAPNoSetMsg != null)
                {
                    //  tt = 1;
                    // MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用");

                    dataGridView1.Enabled = true;
                    DialogResult dialogResult = MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用" + "\r\n" + "是否繼續執行", "AP未啟用", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                    else if (PageList.Count == 1)
                    {
                        return;
                    }

                }


                if (PageList.Count == 1 && PageList[0].APLink == null)
                {

                }

                if (PageList.Count > 0)
                {

                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {

                        kvp.Value.mSmcEsl.stopScanBleDevice();
                    }

                    System.Threading.Thread.Sleep(1000);
                    //int tt = 0;
                    dataGridView1.Enabled = false;
                    testest = true;
                    onlockedbutton(testest);
                    //pictureBox4.Visible = true;
                    ProgressBarVisible(PageList.Count);
                    UpdateESLDen.Text = PageList.Count.ToString();
                
                    List<string> RunAPList = new List<string>();


                    List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).ToList();
                    foreach (Page1 p in list)
                    {
                        RunAPList.Add(p.APLink);

                    }


                    stopwatch.Reset();
                    stopwatch.Start();





                    for (int a = 0; a < RunAPList.Count; a++)
                    {
                        for (int i = 0; i < PageList.Count; i++)
                        {
                            if (PageList[i].APLink == RunAPList[a])
                            {
                                Page1 mPage1 = PageList[i];
                                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                                {

                                    if (kvp.Key.Contains(mPage1.APLink))
                                    {


                                        int Blcount = mPage1.BleAddress.Length;
                                  /*      Bitmap bmp;
                                        if (mPage1.onsale == "V")
                                        {
                                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                              mPage1.specification, mPage1.price, mPage1.Special_offer,
                                               mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                        }
                                        else
                                        {
                                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                              mPage1.specification, mPage1.price, mPage1.Special_offer,
                                               mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                        }*/

                                        int numVal = Convert.ToInt32(mPage1.no) - 1;
                                        Console.WriteLine("mPage1.no" + mPage1.no);
                                        dataGridView1.Rows[numVal].Cells[0].Selected = true;
                                        aaa(datagridview1curr, true, numVal);
                                        dataGridView1.Rows[numVal].Cells[17].Value = DateTime.Now.ToString();
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, numVal, false, openExcelAddress, excel, excelwb, mySheet);
                                     //   pictureBoxPage1.Image = bmp;

                                        Console.WriteLine("ININ");
                                        deviceIPData = mPage1.APLink;
                                       // ConnectBleTimeOut.Start();
                                        kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.BleAddress);

                                      /*  mPage1.TimerConnect = new System.Windows.Forms.Timer();
                                        mPage1.TimerConnect.Interval = (30 * 1000);
                                        mPage1.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                        mPage1.TimerSeconds = new Stopwatch();*/
                                        mPage1.TimerConnect.Start();
                                        mPage1.TimerSeconds.Start();
                                        //    kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                        // kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.BleAddress,0);
                                        // kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.BleAddress);
                                        //System.Threading.Thread.Sleep(1000);
                                        //    EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                        //   mSmcEsl.UpdataESLDataFromBuffer(mPage1.BleAddress, 0, 3,0);
                                        richTextBox1.Text = richTextBox1.Text + PageList[i].usingAddress + "  嘗試連線中請稍候... \r\n";
                                    //    System.Threading.Thread.Sleep(1000);
                                    }
                                }
                                break;
                            }
                        }
                    }
                    //  mSmcEsl.TransformImageToData(bmp);
                    //  mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                    macaddress = PageList[listcount].BleAddress;
                    // richTextBox1.Text = mPage1.BleAddress + "  嘗試連線中請稍候... \r\n";


                }
                rownoinsertcount = dataGridView1.RowCount;
                rownonullinsert = 1;
            }
            else
            {
                MessageBox.Show("ESL更新中請稍後", "更新中");
            }
        }





        private void ExcelInput_Click(object sender, EventArgs e)
        {
            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {

                kvp.Value.mSmcEsl.stopScanBleDevice();
            }
            UpdateESLDen.Text = "0";
            updateESLper.Text = "0";
            Console.WriteLine("ExcelInput_Click");
            MacAddressList.Clear();
            BeaconList.Clear();
            PageList.Clear();
            testest = true;
            dataGridView1.Columns.Clear();
            DataTable d = new DataTable();
            dataGridView1.DataSource = d;

            dataGridView2.Columns.Clear();
            dataGridView2.DataSource = d;

            dataGridView4.Columns.Clear();
            dataGridView4.DataSource = d;

            dataGridView5.Columns.Clear();
            dataGridView5.DataSource = d;

            dataGridView7.Columns.Clear();
            dataGridView7.DataSource = d;

            dataGridView7.Columns.Clear();
            dataGridView7.DataSource = d;

            dataGridView8.Columns.Clear();

            DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
            dgvc.Width = 60;
            dgvc.Name = "ESL綁定";
            dgvc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            DataGridViewColumn dgvc2 = new DataGridViewCheckBoxColumn();
            dgvc2.Width = 60;
            dgvc2.Name = "選取";
            dgvc2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            DataGridViewColumn dgvc7 = new DataGridViewCheckBoxColumn();
            dgvc7.Width = 60;
            dgvc7.Name = "選取";
            dgvc7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewColumn dgvc3 = new DataGridViewCheckBoxColumn();
            dgvc3.Width = 60;
            dgvc3.Name = "選取";
            dgvc3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewColumn dgvc4 = new DataGridViewCheckBoxColumn();
            dgvc4.Width = 60;
            dgvc4.Name = "選取";
            dgvc4.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewColumn dgvc8 = new DataGridViewCheckBoxColumn();
            dgvc8.Width = 40;
            dgvc8.Name = "選取";
            dgvc8.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;


            this.dataGridView1.Columns.Insert(0, dgvc);
            dataGridView8.ColumnCount = 1;
            this.dataGridView2.Columns.Insert(0, dgvc2);
            this.dataGridView4.Columns.Insert(0, dgvc3);
            this.dataGridView5.Columns.Insert(0, dgvc4);
            this.dataGridView7.Columns.Insert(0, dgvc7);
            this.dataGridView8.Columns.Insert(0, dgvc8);
            dataGridView8.Columns[1].Name = "至入選項";
            dataGridView1.MouseDown += new MouseEventHandler(dataGridView1_MouseDown);
            dataGridView5.MouseDown += new MouseEventHandler(dataGridView5_MouseDown);
            //mSmcEsl.stopScanBleDevice();
            datagridview1curr = 0;
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            //  mSmcEsl.DisConnectBleDevice();
            if (excelwb != null)
            {
                excelwb.Save();
                mySheet = null;
                excelwb.Close();
                excelwb = null;
                excel.Quit();
                excel = null;
            }
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //  img1 = ImageDecoder.DecodeFromFile(openFileDialog1.FileName);
                //MessageBox.Show(openFileDialog1.FileName );
                string tableName = "[工作表1$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql = "select * from " + tableName;//SQL查詢
                excel = new Excel.Application();
                excelwb = excel.Workbooks.Open(@openFileDialog1.FileName);
               // excel.Application.Workbooks.Add(true);
                mySheet = new Excel.Worksheet();

                openExcelAddress = openFileDialog1.FileName;
                DataTable kk = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql);
                string tableName2 = "[工作表2$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql2 = "select * from " + tableName2+"WHERE 版型類型=0";//SQL查詢
                DataTable kk2 = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql2);
                string sql5 = "select * from " + tableName2 + "WHERE 版型類型=1";//SQL查詢
                DataTable kk5 = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql5);
                string tableName3 = "[工作表3$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql3 = "select * from " + tableName3;//SQL查詢
                DataTable kk3 = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql3);
                string tableName4 = "[工作表4$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql4 = "select * from " + tableName4;//SQL查詢
                DataTable kk4 = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql4);
                string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                //using (StreamWriter sw = File.CreateText(path)) { }
                string today = DateTime.Now.ToString("yyyyMMdd");
                // filename =filename;
              /* string  filepath = exeDir + @"\" + today + "esldemoV2";
                DataTable datalist = mExcelData.GetExcelDataTable(filepath, sql);
                if (File.Exists(filepath))
                {
                    UpdateESLDen.Text = "0";
                    updateESLper.Text = "0";
                }
                else {
                    foreach (DataRow dr in datalist.Rows) {
                        
                        Console.WriteLine("OPOPOP" + dr[11].ToString());
                        if (dr[11].ToString() == "更新成功") {
                            updateESLper.Text = (Convert.ToInt32(updateESLper.Text) + 1).ToString();
                        }
                    }
                    UpdateESLDen.Text = datalist.Rows.Count.ToString();
                }
                    Console.WriteLine("datalist"+ datalist.Rows.Count);*/
               // Console.WriteLine("headertextall" + kk3.Rows.Count+","+kk3.Columns.Count);
                dataGridView2.DataSource = kk2;
                dataGridView4.DataSource = kk3;
                dataGridView5.DataSource = kk4;
                dataGridView1.DataSource = kk;
                dataGridView7.DataSource = kk5;
           /*     foreach (DataRow dr in kk5.Rows)
                {
                    foreach (DataColumn dc in kk5.Columns)
                    {
                        Console.WriteLine("ssssssssssssss" + dr[dc].ToString());
                       
                    }
               }*/

                BindESL.Text = "0";
                if (dataGridView4.Rows.Count > 1)
                {
                    foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                    {
                        if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() != "未綁定")
                            BindESL.Text = (Convert.ToInt32(BindESL.Text) + 1).ToString();
                    }
                }
                /*  foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                  {
                      if (dr.Cells[1].Value != null && dr.Cells[2].Value != null)
                      {

                          if (dr.Cells[1].Value.ToString() == "192.168.1.15") {
                              mSmcEsl.setUdpClient(dr.Cells[1].Value.ToString(), Convert.ToInt32(dr.Cells[2].Value));
                              mSmcEsl.onScanDeviceEven += new EventHandler(ScanDeviceEven); //掃描ble
                          }

                          if (dr.Cells[1].Value.ToString() == "192.168.1.16")
                          {
                              mSmcEsl.setUdpClient(dr.Cells[1].Value.ToString(), Convert.ToInt32(dr.Cells[2].Value));
                              mSmcEsl.onScanDeviceEven += new EventHandler(ScanDeviceEven1); //掃描ble
                          }


                      }

                  }*/
                //  mSmcEsl.setUdpClient("192.168.1.16", udpport);
                //  mSmcEsl.onScanDeviceEven += new EventHandler(ScanDeviceEven); //掃描ble
                DataGridViewImageColumn columnImage = new DataGridViewImageColumn();
                columnImage.DefaultCellStyle.NullValue = null;
              /*  dgvc3.Width = 60;
                dgvc3.Name = "狀態";
                dgvc3.DefaultCellStyle.NullValue = null;
                dgvc3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;*/
                this.dataGridView1.Columns.Insert(2, columnImage);

                CountESLAll.Text = kk3.Rows.Count.ToString();
                productAll.Text = kk.Rows.Count.ToString();
                dgvc = new DataGridViewCheckBoxColumn();
                dgvc.Width = 50;
                dgvc.Name = "商品選取";
                this.dataGridView1.Columns.Insert(3, dgvc);
             /*   comboBox1.Items.Clear();
                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows) {
                    if(dr5.Cells[2].Value!=null&&dr5.Cells[2].Value.ToString()!="")
                    comboBox1.Items.Add(dr5.Cells[2].Value.ToString());
                }*/


                foreach (DataGridViewColumn column in this.dataGridView1.Columns)
                {
                    //  Console.WriteLine("----------------------------------------");
                    if(column.Index!=0&& column.Index != 2 && column.Index != 3 && column.Index != 4 && column.Index != 5 && column.Index < 12)
                    dataGridView8.Rows.Add(false, column.HeaderText);

                }

                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    //  Console.WriteLine("----------------------------------------");



                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                    {
                        dr.Cells[0].Value = true;
                        dr.DefaultCellStyle.ForeColor = Color.Black;
                    }
                    else
                    {
                        dr.Cells[0].ReadOnly = true;
                        dr.DefaultCellStyle.ForeColor = Color.Gray;
                    }
                       
                    
                       

                }

                dataGridView8.Rows.Add(false, "文字方塊");
                dataGridView8.Rows.Add(false, "文字方塊");
                dataGridView8.Rows.Add(false, "文字方块");
                this.dataGridView4.RowsDefaultCellStyle.BackColor = Color.Bisque;
                this.dataGridView4.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                this.dataGridView5.RowsDefaultCellStyle.BackColor = Color.Bisque;
                this.dataGridView5.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView5.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                this.dataGridView7.RowsDefaultCellStyle.BackColor = Color.Bisque;
                this.dataGridView7.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView7.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                this.dataGridView2.RowsDefaultCellStyle.BackColor = Color.Bisque;
                this.dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                
                this.dataGridView1.RowsDefaultCellStyle.BackColor = Color.Bisque;
                this.dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                this.dataGridView8.RowsDefaultCellStyle.BackColor = Color.Bisque;
                this.dataGridView8.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView8.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                rownoinsertcount = dataGridView1.RowCount;
                Console.WriteLine("this.dataGridView1.ColumnCount" + this.dataGridView1.ColumnCount);
               // this.dataGridView1.Columns[1].ReadOnly = true;
                this.dataGridView1.Columns[2].ReadOnly = true;
                //this.dataGridView1.Columns[2].ReadOnly = true;
                this.dataGridView1.Columns[12].ReadOnly = true;
                this.dataGridView1.Columns[13].ReadOnly = true;
                this.dataGridView1.Columns[14].ReadOnly = true;
                this.dataGridView1.Columns[15].ReadOnly = true;
                this.dataGridView1.Columns[16].ReadOnly = true;
                this.dataGridView1.Columns[17].ReadOnly = true;
                

                this.dataGridView1.Columns[12].Visible = false;
                this.dataGridView1.Columns[14].Visible = false;
                this.dataGridView1.Columns[15].Visible = false;
                this.dataGridView1.Columns[16].Visible = false;
                this.dataGridView1.Columns[17].Visible = false;
                this.dataGridView1.Columns[18].Visible = false;

                if (dataGridView4.Rows.Count > 1) {
                this.dataGridView4.Columns[2].ReadOnly = true;
                this.dataGridView4.Columns[3].ReadOnly = true;
                this.dataGridView4.Columns[4].ReadOnly = true;
                this.dataGridView4.Columns[5].ReadOnly = true;
                this.dataGridView4.Columns[6].ReadOnly = true;
                this.dataGridView4.Columns[8].ReadOnly = true;
                }


                this.dataGridView2.Columns[1].ReadOnly = true;
                this.dataGridView2.Columns[2].ReadOnly = true;
                this.dataGridView2.Columns[0].Width = 20;
                this.dataGridView2.Columns[1].Width = 79;
                this.dataGridView2.Columns[2].Width = 20;
                this.dataGridView7.Columns[1].ReadOnly = true;
                this.dataGridView7.Columns[2].ReadOnly = true;
                this.dataGridView7.Columns[0].Width = 20;
                this.dataGridView7.Columns[1].Width = 79;
                this.dataGridView7.Columns[2].Width = 20;


                this.dataGridView5.Columns[1].ReadOnly = true;
                this.dataGridView5.Columns[2].ReadOnly = true;
                this.dataGridView5.Columns[3].ReadOnly = true;
                this.dataGridView5.Columns[4].ReadOnly = true;
                this.dataGridView5.Columns[3].Visible = false;

                this.dataGridView1.Columns[6].Frozen = true;
                CheckBeaconTimer.Interval = 3500;
                CheckBeaconTimer.Start();
                //this.dataGridView1.Columns[12].Visible = false;
                //this.dataGridView1.Columns[18].Visible = false;
                for (int i = 0; i < this.dataGridView1.ColumnCount; i++)
                {
                    if (i != 0)
                        headertextall = headertextall + ",";
                    headertextall = headertextall + this.dataGridView1.Columns[i].Name;
                }

              //  Console.WriteLine("headertextall" + headertextall);
                for (int ee = 0; ee < this.dataGridView2.ColumnCount; ee++)
                {
                    if (ee != 0 && ee != 1 && ee != 2)
                        this.dataGridView2.Columns[ee].Visible = false;
                }

                for (int ee = 0; ee < this.dataGridView7.ColumnCount; ee++)
                {
                    if (ee != 0 && ee != 1 && ee != 2)
                        this.dataGridView7.Columns[ee].Visible = false;
                }

                if (dataGridView4.Rows.Count > 0) { 
                foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                {
                    foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                    {
                          //  Console.WriteLine("OUOU" + dr5.Cells[5].Value);
                            if (dr5.Cells[5].Value != null && dr5.Cells[5].Value.ToString() == "") {
                               
                                dr5.Cells[5].Value = 0;
                          //      Console.WriteLine("ININ" + dr5.Cells[5].Value);
                            }
                            

                        if (dr5.Cells[2].Value != null && dr4.Cells[8].Value != null && dr5.Cells[2].Value.ToString() == dr4.Cells[8].Value.ToString())
                        {
                                dr5.Cells[5].Value = Convert.ToInt32(dr5.Cells[5].Value)+1;
                        }
                    }
                }
                }


                foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
                {
                    if (dr2.Cells[2].Value != null)
                    {
                        if(dr2.Cells[2].Value.ToString() == "V") {
                      //      Console.WriteLine("HGEEGE"+ dr2.Cells[2].Value.ToString());
                            for (int i = 0; i < dr2.Cells.Count; i++)
                        {
                    //            Console.WriteLine(i+"dr2.Cells[i].Value" + dr2.Cells[i].Value);
                                if (i != 0) {
                           
                                         //  Console.WriteLine("HGEEGE");
                                if (i == 1)
                                {
                                    styleName = dr2.Cells[1].Value.ToString();
                                }
                                if (i != 0 && i != 1 && i != 2)
                                {
                                        if (dr2.Cells[i].Value != null && dr2.Cells[i].Value.ToString() != "")
                                        {
                                            if(dr2.Cells[4].Value.ToString()=="1")
                                                ESL29Format.Add(dr2.Cells[i].Value.ToString());
                                            else if (dr2.Cells[4].Value.ToString() == "2")
                                                ESL42Format.Add(dr2.Cells[i].Value.ToString());
                                            else 
                                                ESLFormat.Add(dr2.Cells[i].Value.ToString());
                                            // Console.WriteLine("ESLFormat" + dr2.Cells[i].Value.ToString());
                                        }
                                        else
                                        {
                                            if(i<dataGridView2.ColumnCount)
                                                if (dr2.Cells[i -1].Value.ToString() != ""&&dr2.Cells[4].Value.ToString() == "0")
                                                    ESLFormat.Add(dr2.Cells[i].Value.ToString());
                                                else if (dr2.Cells[i - 1].Value.ToString() != "" && dr2.Cells[4].Value.ToString() == "1")
                                                    ESL29Format.Add(dr2.Cells[i].Value.ToString());
                                                else if (dr2.Cells[i - 1].Value.ToString() != "" && dr2.Cells[4].Value.ToString() == "2")
                                                    ESL42Format.Add(dr2.Cells[i].Value.ToString());
                                        }
                                    }
                            
                                }
                            }
                        }

                    }
                }


                foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
                {
                    if (dr7.Cells[2].Value != null)
                    {

                        if (dr7.Cells[2].Value.ToString() == "V")
                        {
                            for (int i = 0; i < dr7.Cells.Count; i++)
                            {
                                if (i != 0)
                                {
                                    
                                        //           Console.WriteLine("HGEEGE");
                                        if (i == 1)
                                        {
                                            styleSaleName = dr7.Cells[1].Value.ToString();
                                        }
                                        if (i != 0 && i != 1 && i != 2)
                                        {
                                        if (dr7.Cells[i].Value != null && dr7.Cells[i].Value.ToString() != "")
                                        {
                                            if (dr7.Cells[4].Value.ToString() == "1")
                                                ESLSale29Format.Add(dr7.Cells[i].Value.ToString());
                                            else if (dr7.Cells[4].Value.ToString() == "2")
                                                ESLSale42Format.Add(dr7.Cells[i].Value.ToString());
                                            else
                                                ESLSaleFormat.Add(dr7.Cells[i].Value.ToString());
                                            //   Console.WriteLine("ESLSaleFormat" + dr7.Cells[i].Value.ToString());
                                        }
                                        else
                                        {
                                            if (i < dataGridView7.ColumnCount)
                                                if (dr7.Cells[i - 1].Value.ToString() != "" && dr7.Cells[4].Value.ToString() == "0")
                                                    ESLSaleFormat.Add(dr7.Cells[i].Value.ToString());
                                                else if (dr7.Cells[i - 1].Value.ToString() != "" && dr7.Cells[4].Value.ToString() == "1")
                                                    ESLSale29Format.Add(dr7.Cells[i].Value.ToString());
                                                else if (dr7.Cells[i - 1].Value.ToString() != "" && dr7.Cells[4].Value.ToString() == "2")
                                                    ESLSale42Format.Add(dr7.Cells[i].Value.ToString());
                                        }
                                    }

                                }
                            }
                        }

                    }
                }


           

  
            testest = false;

            APStart = true;
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }
               EslUdpTest.Tools tool = new EslUdpTest.Tools();
           // Tools tool = new Tools();
               tool.onApScanEvent += new EventHandler(AP_Scan);
               tool.SNC_GetAP_Info();


            //  Console.WriteLine("---");
            setLocalTime();
            //datagridview1curr = 2;
            datagridview1curr = 2;
            }

        }


        private void PageOneSelectAll_Click(object sender, EventArgs e)
        {
            Console.WriteLine("PageOneSelectAll_Click");
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != String.Empty)
                {
                    if (pageOneAll == false)
                    {
                        dr.Cells[3].Value = true;
                    }
                    else
                    {
                        dr.Cells[3].Value = false;
                    }
                }
            }
            pageOneAll = !pageOneAll;
        }



        private void ExportToExcel_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = @"imgone.xlsx";
            saveFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                mExcelData.ExportDataGridview(dataGridView1, dataGridView2, dataGridView4, dataGridView5, dataGridView7, true, saveFileDialog1.FileName);

            }

        }



        #endregion Button


        public static string ByteArrayToString(byte[] ba)
        {
            Console.WriteLine("ByteArrayToString");
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            foreach (byte b in ba)
                hex.AppendFormat("{0:x2}", b);
            return hex.ToString();
        }

        //接收UART資料
        private void port1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            Console.WriteLine("port1_DataReceived");
            while (port.BytesToRead != 0)
            {
                packet.Add((byte)port.ReadByte());
            }
            byte[] bArrary = packet.ToArray();
            if (bArrary.Length < 1)
            {
                packet.Clear();
                return;
            }

            bArrary = packet.ToArray();
            if (bArrary[bArrary.Length - 1] == (byte)0x0d)
            {
                byte[] rarray = new byte[bArrary.Length - 1];
                Array.Copy(bArrary, 0, rarray, 0, rarray.Length);
                Display d = new Display(DisplayTextString);
                try
                {
                    this.Invoke(d, new Object[] { rarray });
                }
                catch (Exception ex) { }
            }
        }

        //取得資料並顯示
        private void DisplayTextString(byte[] RX)
        {
            Console.WriteLine("DisplayTextString");
            packet.Clear();
            String decoded = ascii.GetString(RX);
            Console.WriteLine("Decoded string: '{0}'", decoded);
            richTextBox1.Text = decoded;

            decimal number3 = 0;
            Boolean canConvert = decimal.TryParse(decoded, out number3);

            if (onetwo == false)
            {
                onetwo = true;

                rowIndex = 0;
                if (canConvert == true)//仙圖商品條碼
                {
                    Console.WriteLine("@@@@@@@@@  ");
                    doubletype = false;
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {

                        if (dr.Cells[5].Value.ToString() == decoded)
                        {
                            MacAddressList.Add(decoded);
                            Console.WriteLine("@@@@   " + rowIndex);
                            break;
                        }
                        rowIndex++;
                    }
                }
                else//電子標籤ID
                {
                    doubletype = true;
                    dataTemp = decoded;
                    Console.WriteLine("XXXXX  ");
                    //dataGridView1[1, rowIndex].Value = decoded;
                }
            }
            else
            {
                Console.WriteLine("dfffffffffffffffff");

                if (canConvert == true && doubletype == true)//先墊子標籤在商品條碼
                {
                    Console.WriteLine("zzzzzzzzzzzzzz");

                    Boolean maccheck = false;
                    foreach (string mac in MacAddressList)
                    {
                        Console.WriteLine("tttttttttt");
                        if (mac.Equals(dataTemp))
                        {
                            Console.WriteLine("ssssssdd");
                            maccheck = true;
                        }
                    }
                    rowIndex = 0;
                    if (maccheck == false)
                    {
                        Console.WriteLine("aaaaaaaaa");
                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                        {
                            if (dr.Cells[5].Value.ToString() == decoded)
                            {
                                //dataGridView1[1, rowIndex].Value = dataTemp;

                                Console.WriteLine("sddddd");
                                //----------------------------------
                                if (dr.Cells[1].Value.ToString().Length > 11)
                                {
                                    dataGridView1[1, rowIndex].Value = dataGridView1[1, rowIndex].Value + "," + dataTemp;
                                }
                                else
                                {
                                    Console.WriteLine("qqwwe");
                                    dataGridView1[1, rowIndex].Value = dataTemp;
                                    Console.WriteLine("YOYOYOOY");
                                }
                                MacAddressList.Add(dataTemp);
                                break;
                            }

                            rowIndex++;

                        }

                    }
                    else
                    {
                        Console.WriteLine("xxxxxxxxxxxx");
                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                        {

                            if (dr.Cells[1].Value.ToString().Contains(',' + dataTemp))
                            {
                                int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + dataTemp);
                                dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                break;
                            }
                            if (dr.Cells[1].Value.ToString().Contains(dataTemp + ','))
                            {

                                int changeaddr = dr.Cells[1].Value.ToString().IndexOf(dataTemp + ',');
                                dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                break;
                            }

                            if (dr.Cells[1].Value.ToString().Contains(dataTemp))
                            {

                                int changeaddr = dr.Cells[1].Value.ToString().IndexOf(dataTemp);
                                dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                                break;
                            }

                        }
                        string aaa = MacAddressList[MacAddressList.Count - 1];
                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                        {
                            if (dr.Cells[5].Value.ToString() == decoded)
                            {
                                if (dr.Cells[1].Value.ToString().Length > 1)
                                {
                                    dataGridView1[1, rowIndex].Value = dataGridView1[1, rowIndex].Value + "," + dataTemp;
                                }
                                else
                                {
                                    dataGridView1[1, rowIndex].Value = dataTemp;
                                }
                                break;
                            }
                            rowIndex++;

                        }

                        MacAddressList.Add(decoded);
                    }
                }
                if (canConvert == false && doubletype == false)//先商品條碼在墊子標籤
                {
                   // Console.WriteLine("ssssssssssssss");
                    Boolean maccheck = false;
                    foreach (string mac in MacAddressList)
                    {
                        if (mac.Equals(decoded))
                        {
                            maccheck = true;
                        }
                    }
                    if (maccheck == false)
                    {

                        string aaa = MacAddressList[MacAddressList.Count - 1];

                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                        {
                            if (dr.Cells[5].Value.ToString() == aaa)
                            {
                                if (dr.Cells[1].Value.ToString().Length > 1)
                                {
                                    dataGridView1[1, rowIndex].Value = dataGridView1[1, rowIndex].Value + "," + decoded;
                                }
                                else
                                {
                                    dataGridView1[1, rowIndex].Value = decoded;
                                }
                                break;
                            }

                        }

                        MacAddressList.Add(decoded);

                    }
                    else
                    {
                        Console.WriteLine("dddddddddddddddd");
                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                        {

                            if (dr.Cells[1].Value.ToString().Contains(',' + decoded))
                            {
                                int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + decoded);
                                dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                break;
                            }
                            if (dr.Cells[1].Value.ToString().Contains(decoded + ','))
                            {

                                int changeaddr = dr.Cells[1].Value.ToString().IndexOf(decoded + ',');
                                dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                break;
                            }

                            if (dr.Cells[1].Value.ToString().Contains(decoded))
                            {

                                int changeaddr = dr.Cells[1].Value.ToString().IndexOf(decoded);
                                dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                                break;
                            }

                        }
                        string aaa = MacAddressList[MacAddressList.Count - 1];
                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                        {
                            if (dr.Cells[5].Value.ToString() == aaa)
                            {
                                if (dr.Cells[1].Value.ToString().Length > 1)
                                {
                                    dataGridView1[1, rowIndex].Value = dataGridView1[1, rowIndex].Value + "," + decoded;
                                }
                                else
                                {
                                    dataGridView1[1, rowIndex].Value = decoded;
                                }
                                break;
                            }

                        }

                        MacAddressList.Add(decoded);

                    }

                }


                onetwo = false;
            }


        }


        //----------------------------------------------------------------------------------------
        private void BeaconSelectAll_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[4].Value != null)
                {
                    if (beaconAll == false)
                    {
                        dr.Cells[3].Value = true;
                    }
                    else
                    {
                        dr.Cells[3].Value = false;
                    }
                }
            }
            beaconAll = !beaconAll;
        }

        private void BeaconStartUP_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }

            if (!APStart)
            {
                MessageBox.Show("請先連接AP");
                dataGridView1.Enabled = true;
                return;
            }

            /*  BeaconList.Clear();
              foreach (DataGridViewRow dr in this.dataGridView1.Rows)
              {
                  if (dr.Cells[3].Value != null && (bool)dr.Cells[3].Value)
                  {
                      BeaconList.Add(dr.Cells[5].Value.ToString());
                  }
              }
              beacon_index = 0;*/
            string value = "Document 1"; 
            if (InputBox("DateTimePicker", "推播時間設定", "開始時間:", ref value) == DialogResult.OK)
            {
                // System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
                // messageBoxCS.AppendFormat("{0} = {1}", "ClickedItem", e.ClickedItem);
                // messageBoxCS.AppendLine();t
                //MessageBox.Show(messageBoxCS.ToString(), "ItemClicked Event");
                List<string> sss = new List<string>();
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    if (dr.Cells[3].Value!=null&&(bool)dr.Cells[3].Value) {
                        dr.Cells[21].Value = BeaconTimeS;
                        dr.Cells[22].Value = BeaconTimeE;
                        dr.Cells[23].Value = beaconsales;
                        sss.Add(dr.Cells[5].Value.ToString());
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 21, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 22, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 23, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                    }
                }
                Console.WriteLine("BeaconTimeS:"+ BeaconTimeS+ "BeaconTimeE" + BeaconTimeE);
                beacon_data_set(sss, BeaconTimeS, BeaconTimeE, beaconsales);




              //  if (BeaconList.Count > 0)
              //     mSmcEsl.WriteBeaconData("ESL143AP01", BeaconList[beacon_index], false);
            }

            


            //BeaconList
        }
        //推播Beacon倫尋檢查
        private void BeaconCheckUpdate(object sender, EventArgs e)
            {
            Console.WriteLine(testest + "Beacon Check Not IN"+ APStart);

            if (!testest&&APStart)
            { 
           Console.WriteLine("Beacon Check");
           // mSmcEsl.stopScanBleDevice();
            BeaconList.Clear();
            beacon_index = 0;
            BeaconListUpdate.Clear();

                string saletimemsg="";
                string beaconmsg="";
                List<string> nullbeacon = new List<string>();
                //beaconAll CHECK=============================
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    if (dr.Cells[21].Value != null && dr.Cells[22].Value != null && dr.Cells[21].Value.ToString() != "" && dr.Cells[22].Value.ToString() != "")
                    {

                        string format = "yyyy/MM/dd HH:mm:ss";
                        string start = Convert.ToDateTime(dr.Cells[21].Value).ToString("yyyy/MM/dd HH:mm:ss");
                        string end = Convert.ToDateTime(dr.Cells[22].Value).ToString("yyyy/MM/dd HH:mm:ss");
                        DateTime strDate = DateTime.ParseExact(start, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                        DateTime endDate = DateTime.ParseExact(end, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                        Console.WriteLine("start"+ start+ "endDate"+ endDate+ DateTime.Now);
                     
                        if (DateTime.Compare(strDate, DateTime.Now) < 0 && DateTime.Compare(endDate, DateTime.Now) > 0)
                        {
                            Page mPage = new Page();
                            Console.WriteLine("----------------1111");
                            foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                            {
                                if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString() !=""&& dr.Cells[1].Value.ToString().Contains(dr4.Cells[1].Value.ToString()))
                                {
                                    Console.WriteLine(dr.Cells[1].Value.ToString()+"----------------222222"+ dr4.Cells[1].Value.ToString());
                                    if (dr.Cells[1].Value.ToString().Length > 13)
                                    {
                                        string[] BeaconProductAll = dr.Cells[1].Value.ToString().Split(',');
                                        Console.WriteLine("----------------3333333333");
                                        for (int i = 0; i < BeaconProductAll.Length; i++)
                                        {
                                            Console.WriteLine("----------------444444");
                                            if (dr4.Cells[1].Value.ToString() == BeaconProductAll[i])
                                            {
                                                mPage.BeaconProduct = dr.Cells[5].Value.ToString();
                                                mPage.ProductName = dr.Cells[6].Value.ToString();
                                                mPage.SBeaconTime = Convert.ToDateTime(dr.Cells[21].Value.ToString());
                                                mPage.EBeaconTime = Convert.ToDateTime(dr.Cells[22].Value.ToString());
                                             /*   TimeSpan ts = mPage.EBeaconTime - mPage.SBeaconTime;
                                                double days = ts.TotalDays;
                                                if (days < 10)
                                                    mPage.salesDay = "0" + Convert.ToInt32(days).ToString();
                                                else
                                                    mPage.salesDay = Convert.ToInt32(days).ToString();

                                                mPage.Comment = dr.Cells[23].Value.ToString();*/
                                                mPage.APID = dr4.Cells[8].Value.ToString();
                                                Console.WriteLine("----------------OKOK");
                                                BeaconList.Add(mPage);
                                                BeaconListUpdate.Add(mPage.BeaconProduct+ mPage.ProductName + mPage.SBeaconTime + mPage.EBeaconTime + mPage.Comment);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("BeaconListnji3.3"+ dr4.Cells[1].Value.ToString());
                                        mPage.BeaconProduct = dr.Cells[5].Value.ToString();
                                        mPage.ProductName = dr.Cells[6].Value.ToString();
                                        mPage.SBeaconTime = Convert.ToDateTime(dr.Cells[21].Value.ToString());
                                        mPage.EBeaconTime = Convert.ToDateTime(dr.Cells[22].Value.ToString());
                                      /*  mPage.Comment = dr.Cells[23].Value.ToString();
                                        TimeSpan ts =  mPage.EBeaconTime - mPage.SBeaconTime;
                                        double days = ts.TotalDays;
                                        if (days < 10)
                                            mPage.salesDay = "0" + Convert.ToInt32(days).ToString();
                                        else
                                            mPage.salesDay = Convert.ToInt32(days).ToString();
                                        */
                                        mPage.APID = dr4.Cells[8].Value.ToString();
                                        BeaconList.Add(mPage);
                                        BeaconListUpdate.Add(mPage.BeaconProduct + mPage.ProductName + mPage.SBeaconTime + mPage.EBeaconTime + mPage.Comment);
                                    }
                                }
                            }
                            Console.WriteLine("Beacon 加入");
                            Console.WriteLine("CH 加入" + BeaconListNow.Count() + BeaconList.Count());
                        }
                        else {
                            //CheckBeaconTimer.Stop();
                           
                            beaconmsg = beaconmsg + dr.Cells[6].Value + "促銷推播已到期" + dr.Cells[21].Value + "-" + dr.Cells[22].Value + "\r\n";
                            //    if (result == DialogResult.OK) {
                            // Do something
                            dr.Cells[21].Value = DBNull.Value;
                            dr.Cells[22].Value = DBNull.Value;
                            dr.Cells[23].Value = DBNull.Value;
                            nullbeacon.Add(dr.Cells[5].Value.ToString());
                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 21, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 22, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 23, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                            productState(dr);
                            //CheckBeaconTimer.Start();
                            //}

                        }
                    }

                }

                if (nullbeacon.Count!=0)
                    beacon_data_set(nullbeacon, "", "", "");

                if (beaconmsg != "") {
                    CheckBeaconTimer.Stop();
                    DialogResult result = MessageBox.Show(beaconmsg, "Beacon訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    if (result == DialogResult.OK) {
                        CheckBeaconTimer.Start();
                    }
                }
                if (!BeaconListUpdate.SequenceEqual(BeaconListNow))
                {
                    Console.WriteLine("Beacon 更新");

                    if (BeaconList.Count > 0)
                    {
                        //setLocalTime(BeaconList[beacon_index].APID);
                        testest = true;
                        foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                        {
                            if (kvp.Key.Contains(BeaconList[beacon_index].APID))
                            {
                                setBeaconTime(BeaconList[beacon_index].APID);
                              //  System.Threading.Thread.Sleep(100);
                                Console.WriteLine("ESL143AP01"+ BeaconList[beacon_index].BeaconProduct+BeaconList[beacon_index].Comment+ BeaconList[beacon_index].salesDay);
                                if(BeaconList.Count==1)
                                    kvp.Value.mSmcEsl.WriteBeaconData("ESL143AP01", BeaconList[beacon_index].BeaconProduct, true);
                                else
                                    kvp.Value.mSmcEsl.WriteBeaconData("ESL143AP01", BeaconList[beacon_index].BeaconProduct, false);
                            }
                        }

                    }
                    else
                    {
                        Page mPage = new Page();
                        mPage.BeaconProduct = "0000000000000";
                        BeaconList.Add(mPage);

                        foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                        {
                                kvp.Value.mSmcEsl.setBeaconTime(18, 12, 31, 23, 59, 99, 12, 31, 23, 59);
                                kvp.Value.mSmcEsl.WriteBeaconData("ESL143AP01", BeaconList[beacon_index].BeaconProduct, true);
                        }
                    }

                    BeaconListNow.Clear();
                    BeaconListNow.AddRange(BeaconListUpdate);

                }



                //SALE ALL CHECK===========================================
                // PageList.Clear();
                //  listcount = 0;
               SalePageListUpdate.Clear();
                Console.WriteLine("BeaconPP"+ PageList.Count);
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    if (dr.Cells[19].Value != null && dr.Cells[20].Value != null && dr.Cells[19].Value.ToString() != "" && dr.Cells[20].Value.ToString() != "")
                    {
                        Console.WriteLine("199119" + dr.Cells[6].Value.ToString());
                        Console.WriteLine("199119"+dr.Cells[19].Value.ToString());
                        string format = "yyyy/MM/dd HH:mm:ss";
                        string start = Convert.ToDateTime(dr.Cells[19].Value).ToString("yyyy/MM/dd HH:mm:ss");
                        string end = Convert.ToDateTime(dr.Cells[20].Value).ToString("yyyy/MM/dd HH:mm:ss");
                        DateTime strDate = DateTime.ParseExact(start, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                        DateTime endDate = DateTime.ParseExact(end, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                        if (dr.Cells[1].Value.ToString().Length > 13)
                            {

                                string[] usingAddressSplit = dr.Cells[1].Value.ToString().Split(',');
                                string Special_offer;
                                string onSale="";
                                string updateStyle = "";
                            
                                for (int i = 0; i < usingAddressSplit.Length; i++)
                                {
                                    Page1 mPageA = new Page1();
                                    mPageA.no = (dr.Index + 1).ToString();
                                    mPageA.BleAddress = dr.Cells[1].Value.ToString();
                                    mPageA.barcode = dr.Cells[5].Value.ToString();
                                    mPageA.product_name = dr.Cells[6].Value.ToString();
                                    mPageA.Brand = dr.Cells[7].Value.ToString();
                                    mPageA.specification = dr.Cells[8].Value.ToString();
                                    mPageA.price = dr.Cells[9].Value.ToString();
                                    mPageA.Web = dr.Cells[11].Value.ToString();
                                    mPageA.usingAddress = usingAddressSplit[i];
                                    mPageA.HeadertextALL = headertextall;
                                    mPageA.Special_offer = dr.Cells[10].Value.ToString();
                                    mPageA.onsale = onSale;
                                    mPageA.onSaleTimeS = dr.Cells[19].Value.ToString();
                                    mPageA.TimerConnect = new System.Windows.Forms.Timer();
                                    mPageA.TimerConnect.Interval = (30 * 1000);
                                    mPageA.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                    mPageA.TimerSeconds = new Stopwatch();
                                    mPageA.onSaleTimeE = dr.Cells[20].Value.ToString();
                                    mPageA.actionName = "saletime";
                                Console.WriteLine(" mPageA.ProductStyle " + mPageA.ProductStyle);
                                foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                    {
                                        if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == usingAddressSplit[i])
                                        {
                                            mPageA.APLink = drAP.Cells[8].Value.ToString();
                                            break;
                                        }
                                    }
                                if (DateTime.Compare(strDate, DateTime.Now) < 0 && DateTime.Compare(endDate, DateTime.Now) > 0)
                                {

                                    if(dr.Cells[15].Value!=null&& dr.Cells[15].Value.ToString()== "X") { 
                                    dr.Cells[15].Value = onSale = "V";
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    Special_offer = dr.Cells[10].Value.ToString();
                                    updateStyle = styleSaleName;
                                     mPageA.ProductStyle = updateStyle;
                                        SalePageListUpdate.Add(mPageA);
                                    }
                                }
                                else
                                {
                                    if (dr.Cells[15].Value != null && dr.Cells[15].Value.ToString() == "V")
                                    {
                                        dr.Cells[15].Value = onSale = "X";
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        Special_offer = dr.Cells[10].Value.ToString();
                                        updateStyle = styleName;
                                        mPageA.ProductStyle = updateStyle;
                                        SalePageListUpdate.Add(mPageA);
                                    }
                                }
                                //  PageList.Add(mPageA);
                                

                            }
                            }
                            else
                            {
                                Console.WriteLine("dr.Cells[6].Value.ToString()" + dr.Cells[6].Value.ToString());
                                Page1 mPageC = new Page1();
                                mPageC.no = (dr.Index + 1).ToString();
                                mPageC.BleAddress = dr.Cells[1].Value.ToString();
                                mPageC.barcode = dr.Cells[5].Value.ToString();
                                mPageC.product_name = dr.Cells[6].Value.ToString();
                                mPageC.Brand = dr.Cells[7].Value.ToString();
                                mPageC.specification = dr.Cells[8].Value.ToString();
                                mPageC.price = dr.Cells[9].Value.ToString();
                               
                                mPageC.Web = dr.Cells[11].Value.ToString();
                                mPageC.usingAddress = dr.Cells[1].Value.ToString();
                                mPageC.HeadertextALL = headertextall;
                                mPageC.TimerConnect = new System.Windows.Forms.Timer();
                                mPageC.TimerConnect.Interval = (30 * 1000);
                                mPageC.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                mPageC.TimerSeconds = new Stopwatch();
                            //mPageC.Special_offer = dr.Cells[10].Value.ToString();
                            string updateStyle="";
                            
                               
                                mPageC.onsale = dr.Cells[15].Value.ToString();
                                mPageC.onSaleTimeS = dr.Cells[19].Value.ToString();
                                mPageC.onSaleTimeE = dr.Cells[20].Value.ToString();
                                mPageC.actionName = "saletime";
                            Console.WriteLine(" mPageC.ProductStyle " + mPageC.ProductStyle);
                                foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                {
                                    if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == dr.Cells[1].Value.ToString())
                                    {
                                        mPageC.APLink = drAP.Cells[8].Value.ToString();
                                        break;
                                    }
                                }
                            if (DateTime.Compare(strDate, DateTime.Now) < 0 && DateTime.Compare(endDate, DateTime.Now) > 0)
                            {
                                dr.Cells[15].Value = mPageC.onsale = "V";
                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                mPageC.Special_offer = dr.Cells[10].Value.ToString();
                                updateStyle = styleSaleName;
                                mPageC.ProductStyle = updateStyle;
                                SalePageListUpdate.Add(mPageC);
                            }
                            else
                            {
                                dr.Cells[15].Value = mPageC.onsale = "X";
                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                mPageC.Special_offer = dr.Cells[10].Value.ToString();
                                updateStyle = styleName;
                                mPageC.ProductStyle = updateStyle;
                                SalePageListUpdate.Add(mPageC);
                            }
                            
                        }
                    }
                     }

               

                // ------------------明天 初始修改

                if (SalePageListUpdate.Count !=0)
                {

                    ProgressBarVisible(PageList.Count);
                    List<string> RunAPList = new List<string>();
                    List<Page1> list = SalePageListUpdate.GroupBy(a => a.APLink).Select(g => g.First()).ToList();
                    foreach (Page1 p in list)
                    {
                        RunAPList.Add(p.APLink);

                    }

                    foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                    {
                        if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == SalePageListUpdate[0].APLink)
                        {
                            if (dr5.Cells[4].Value.ToString() == "")
                            {
                                if (down || sale)
                                {
                                    
                                   MessageBox.Show(SalePageListUpdate[0].usingAddress + "該ESL綁定AP未啟用");
                                }
                                else
                                {
                                    
                                   MessageBox.Show(SalePageListUpdate[0].BleAddress + "該ESL綁定AP未啟用");
                                }
                                return;
                            }
                        }

                    }

                   
                 //   sale = true;
                  
                    Boolean assalepage= true;
                    //mSmcEsl.TransformImageToData(bmp);
                    Console.WriteLine("Count" + SalePageListUpdate.Count+"k"+ SalePageList.Count+"p"+PageList.Count);

                    if (SalePageListUpdate.Count != SalePageList.Count)
                    {
                        assalepage = false;
                    }
                    else
                    {
                        if (SalePageList.Count != 0) { 
                        for (int i = 0; i < SalePageListUpdate.Count; i++) {

                            if (SalePageListUpdate[i].onsale != SalePageList[i].onsale || SalePageListUpdate[i].onSaleTimeS != SalePageList[i].onSaleTimeS || SalePageListUpdate[i].onSaleTimeE != SalePageList[i].onSaleTimeE || SalePageListUpdate[i].product_name != SalePageList[i].product_name || SalePageListUpdate[i].barcode != SalePageList[i].barcode || SalePageListUpdate[i].price != SalePageList[i].price || SalePageListUpdate[i].Special_offer != SalePageList[i].Special_offer || SalePageListUpdate[i].specification != SalePageList[i].specification)
                            {
                                    Console.WriteLine("LOOK"+ SalePageListUpdate+"and"+ SalePageList);
                                    assalepage = false;
                            }
                        }
                        }
                    }
                    if (!assalepage)
                    {
                        foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                        {

                            kvp.Value.mSmcEsl.stopScanBleDevice();
                        }
                        testest = true;
                        onlockedbutton(testest);
                        saletime = true;
                        /*for (int i = 0; i < SalePageListUpdate.Count; i++)
                        {
                            PageList.Add(SalePageListUpdate[i]);
                        }*/

                            

                        //SalePageList.Clear();
                     /*   for (int i = 0; i < SalePageListUpdate.Count; i++)
                        {
                            Page1 mPageA = new Page1();
                            mPageA.no = SalePageListUpdate[i].no;
                            mPageA.BleAddress = SalePageListUpdate[i].BleAddress;
                            mPageA.barcode = SalePageListUpdate[i].barcode;
                            mPageA.product_name = SalePageListUpdate[i].product_name;
                            mPageA.Brand = SalePageListUpdate[i].Brand;
                            mPageA.specification = SalePageListUpdate[i].specification;
                            mPageA.price = SalePageListUpdate[i].price;
                            mPageA.Web = SalePageListUpdate[i].Web;
                            mPageA.usingAddress = SalePageListUpdate[i].usingAddress;
                            mPageA.HeadertextALL = SalePageListUpdate[i].HeadertextALL;
                            mPageA.Special_offer = SalePageListUpdate[i].Special_offer;
                            mPageA.onsale = SalePageListUpdate[i].onsale;
                            mPageA.onSaleTimeS = SalePageListUpdate[i].onSaleTimeS;
                            mPageA.ProductStyle = SalePageListUpdate[i].ProductStyle;
                            mPageA.onSaleTimeE = SalePageListUpdate[i].onSaleTimeE;
                            mPageA.actionName = SalePageListUpdate[i].actionName;
                            mPageA.APLink = SalePageListUpdate[i].APLink;
                            PageList.Add(mPageA);
                            SalePageList.Add(mPageA);
                        }*/

                        //PageList.AddRange(SalePageListUpdate);
                        PageList = SalePageListUpdate;
                        listcount = 0;
                        stopwatch.Reset();
                        stopwatch.Start();
                        Console.WriteLine("不依樣近來更新");
                        ProgressBarVisible(PageList.Count);
                        for (int a = 0; a < RunAPList.Count; a++)
                        {
                            for (int i = 0; i < PageList.Count; i++)
                            {
                                Console.WriteLine("(PageList[i]" + PageList[i].APLink + "RunAPList[a]" + RunAPList[a]);
                                if (PageList[i].APLink == RunAPList[a])
                                {
                                    Console.WriteLine("PageList[i].APLink" + PageList[i].APLink);
                                    Page1 mPage1 = PageList[i];
                                    if (mPage1.usingAddress != "")
                                    {
                                        // int Blcount = mPage1.BleAddress.Length;
                                        string format = "yyyy/MM/dd HH:mm:ss";
                                        string starta = Convert.ToDateTime(mPage1.onSaleTimeS).ToString("yyyy/MM/dd HH:mm:ss");
                                        string enda = Convert.ToDateTime(mPage1.onSaleTimeE).ToString("yyyy/MM/dd HH:mm:ss");
                                        DateTime strDatea = DateTime.ParseExact(starta, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                                        DateTime endDatea = DateTime.ParseExact(enda, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                                        Bitmap bmp;
                                        if (DateTime.Compare(strDatea, DateTime.Now) < 0 && DateTime.Compare(endDatea, DateTime.Now) > 0)
                                        {
                                             bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                             mPage1.specification, mPage1.price, mPage1.Special_offer,
                                             mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat); 

                                        }
                                        else
                                        {
                                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                            mPage1.specification, mPage1.price, mPage1.Special_offer,
                                            mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                            //CheckBeaconTimer.Stop();
                                            saletimemsg = saletimemsg + PageList[i].product_name + "特價已到期" + PageList[i].onSaleTimeS + "-" + PageList[i].onSaleTimeE + "\r\n";
                                            //if (result == DialogResult.OK)
                                            //{
                                            // Do something
                                            //1/31
                                            foreach (DataGridViewRow dr in this.dataGridView1.Rows) {
                                                if (dr.Cells[5].Value!=null&&PageList[i].barcode == dr.Cells[5].Value.ToString())
                                                {
                                                    dr.Cells[19].Value = DBNull.Value;
                                                    dr.Cells[20].Value = DBNull.Value;
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                }
                                            }
                                            //CheckBeaconTimer.Start();
                                            //}
                                        }

                                        foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                                        {

                                                kvp.Value.mSmcEsl.stopScanBleDevice();

                                        }
                                        foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                                        {

                                            if (kvp.Key.Contains(mPage1.APLink))
                                            {



                                                int numVal = Convert.ToInt32(mPage1.no) - 1;
                                                Console.WriteLine("mPage1.no" + mPage1.no);
                                                Console.WriteLine("mPage1.APLink" + mPage1.APLink);
                                                //dataGridView1.ClearSelection();
                                                dataGridView1.Rows[numVal].Cells[0].Selected = true;
                                                aaa(datagridview1curr, true, numVal);
                                                dataGridView1.Rows[numVal].Cells[17].Value = DateTime.Now.ToString();
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, numVal, false, openExcelAddress, excel, excelwb, mySheet);
                                                pictureBoxPage1.Image = bmp;

                                                //Console.WriteLine("ININ");
                                                // kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                                //  kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress,0);
                                                deviceIPData = mPage1.APLink;
                                                //ConnectBleTimeOut.Start();
                                                kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.usingAddress);
                                             /*   mPage1.TimerConnect = new System.Windows.Forms.Timer();
                                                mPage1.TimerConnect.Interval = (30 * 1000);
                                                mPage1.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                                mPage1.TimerSeconds = new Stopwatch();*/
                                                mPage1.TimerSeconds.Start();
                                                mPage1.TimerConnect.Start();
                                                //  System.Threading.Thread.Sleep(100);
                                                //      SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                                //    mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3);
                                                //  richTextBox1.Text = mPage1.usingAddress + "  嘗試連線中請稍候... \r\n";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("該商品" + mPage1.product_name + "未裝置電子標籤");
                                        dataGridView1.Enabled = true;
                                    }
                                    break;
                                }
                            }
                        }
                        SalePageList.Clear();
                        /*  for (int i=0;i< SalePageListUpdate.Count;i++) {
                              SalePageList.Add(SalePageListUpdate[i]);
                          }*/

                     /*   for (int i = 0; i < SalePageListUpdate.Count; i++)
                        {
                            Page1 mPageA = new Page1();
                            mPageA.no = SalePageListUpdate[i].no;
                            mPageA.BleAddress = SalePageListUpdate[i].BleAddress;
                            mPageA.barcode = SalePageListUpdate[i].barcode;
                            mPageA.product_name = SalePageListUpdate[i].product_name;
                            mPageA.Brand = SalePageListUpdate[i].Brand;
                            mPageA.specification = SalePageListUpdate[i].specification;
                            mPageA.price = SalePageListUpdate[i].price;
                            mPageA.Web = SalePageListUpdate[i].Web;
                            mPageA.usingAddress = SalePageListUpdate[i].usingAddress;
                            mPageA.HeadertextALL = SalePageListUpdate[i].HeadertextALL;
                            mPageA.Special_offer = SalePageListUpdate[i].Special_offer;
                            mPageA.onsale = SalePageListUpdate[i].onsale;
                            mPageA.onSaleTimeS = SalePageListUpdate[i].onSaleTimeS;
                            mPageA.ProductStyle = SalePageListUpdate[i].ProductStyle;
                            mPageA.onSaleTimeE = SalePageListUpdate[i].onSaleTimeE;
                            mPageA.actionName = SalePageListUpdate[i].actionName;
                            mPageA.APLink = SalePageListUpdate[i].APLink;
                            SalePageList.Add(mPageA);
                        }*/
                        SalePageList.AddRange(SalePageListUpdate);


                    }
                   
                    // dataGridView3.Rows[datagridview2no].Cells[4].Value = "連線中";

                    //  mSmcEsl.TransformImageToData(bmp);
                    
                    // mSmcEsl.ConnectBleDevice(mPage1.usingAddress);

                    // mSmcEsl.WriteESLData(mPage1.usingAddress);
                    macaddress = PageList[listcount].usingAddress;
                    string sub = Environment.CurrentDirectory;
                    Console.WriteLine("sub" + sub);

                    // dataGridView3.Rows[0].Cells[4].Value = "連線中";
                    

                }
                if (saletimemsg != "")
                {
               //     CheckBeaconTimer.Stop();
              //      DialogResult result = MessageBox.Show(saletimemsg, "Beacon訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
               /*     if (result == DialogResult.OK)
                    {
                        CheckBeaconTimer.Start();
                    }*/

                }

            }
        }

        public void ProgressBarVisible(int Size)
        {
            Console.WriteLine("ProgressBarVisible");
            Size = Size * 10;
            this.progressBar1.Visible = true; //顯示進度條
            progressBar1.Maximum = Size;//設置最大長度值
            progressBar1.Value = 0;//設置當前值
            progressBar1.Step = 10;//設置沒次增長多少
            //progressBar1.Increment(listcount);
        }
        // int j = 0;
        private void ReceiveData(EventArgs e)
        {


            int msgId = (e as EslUdpTest.SmcEsl.SMCEslReceiveEventArgs).msgId;
            bool status = (e as EslUdpTest.SmcEsl.SMCEslReceiveEventArgs).status;
            string deviceIP = (e as EslUdpTest.SmcEsl.SMCEslReceiveEventArgs).apIP;
            string data = (e as EslUdpTest.SmcEsl.SMCEslReceiveEventArgs).data;
            double battery = (e as EslUdpTest.SmcEsl.SMCEslReceiveEventArgs).battery;
            string[] IP = deviceIP.Split(':');
            ESLFromIP = IP[0];
            string str_data = "";
            Console.WriteLine("UpdateUI"+ data);
            /*   if (data.Equals("連線成功"))
               {

                   countconnect = 0;
                   this.progressBar1.Visible = true;
                   updateESLper.Text= (Convert.ToUInt32(updateESLper.Text)+1).ToString();

                   Console.WriteLine("-----------------------");
                   richTextBox1.Text = "連線成功";
                   if (CheckESLOnly) {
                       APESLState.Text = "連線成功";
                       CheckESLOnly = false;
                       mSmcEsl.DisConnectBleDevice();
                   }
                   if (checkaddress.Count <= 0)
                   {
                       ConnectTimer.Stop();
                       if (down || sale)
                       {
                           mSmcEsl.WriteESLData(PageList[listcount].usingAddress);
                       }
                       else
                       {
                           mSmcEsl.WriteESLData(PageList[listcount].BleAddress);
                       }
                   }
                   else
                   {

                       foreach (DataGridViewRow dr in dataGridView1.Rows)
                       {
                           Console.WriteLine("dr.Cells[2].Value.ToString()");
                           if (
                           .Value != null)
                           {
                               if (dr.Cells[12].Value.ToString().Contains(checkaddress[checkconnectcount]) && dr.Cells[4].Style.BackColor != Color.Red)
                               {
                                   dr.Cells[4].Style.BackColor = Color.Green;
                                   dr.Cells[4].Value = DateTime.Now.ToString();
                               }
                           }
                       }
                       checkconnectcount++;
                       if (checkconnectcount >= checkaddress.Count)
                       {


                           checkconnectcount = 0;
                           checkaddress.Clear();
                           richTextBox1.Text = "輪尋完畢" + "\r\n" + richTextBox1.Text;
                       }
                       else
                       {
                           CheckConnectTimer.Interval = 3000;
                           CheckConnectTimer.Start();
                       }

                   }


               }
               else if (data.Equals("連線失敗"))
               {
                   Console.WriteLine("~~~~~~~~~~");
                   Runtime = true;
                   if (countconnect < 5)
                   {
                       if (CheckESLOnly)
                       {
                           Console.WriteLine("CheckESLOnly");
                           APESLState.Text = "重新嘗試" + countconnect + "次";
                           //1/2新年第一天上工摟~~~
                           ConnectTimer.Interval = 3000;
                           ConnectTimer.Start();
                           countconnect++;

                       }
                       else {
                           Console.WriteLine("ERROR" + countconnect);
                           ConnectTimer.Interval = 3000;
                           ConnectTimer.Start();
                           countconnect++;
                       }

                   }
                   else
                   {

                           countconnect = 0;

                       foreach (DataGridViewRow dr in dataGridView1.Rows)
                           {

                               if (dr.Cells[2].Value != null)
                               {
                                   if (dr.Cells[2].Value.ToString() == macaddress)
                                   {
                                   ESLFailData.Clear();
                                       dr.Cells[4].Style.BackColor = Color.Red;
                                       dr.Cells[4].Value = DateTime.Now.ToString();
                                       PageList[listcount].UpdateState = "更新失敗";
                                       PageList[listcount].UpdateTime = DateTime.Now.ToString();
                                       int failcount = ESLUpdaateFail.Count;
                                   Console.WriteLine("dr.Cells[2].Value.ToString()"+ dr.Cells[2].Value.ToString() + failcount);
                                   ESLFailData.Add(PageList[listcount].BleAddress);
                                   ESLFailData.Add(DateTime.Now.ToString());
                                   ESLFailData.Add("連線失敗");
                                   ESLUpdaateFail.Add(ESLFailData);
                                   // Page1 mPage1 = PageList[listcount];
                                   foreach (DataGridViewRow dr3 in dataGridView3.Rows)
                                       {
                                       Console.WriteLine("1111111111"+ dr3.Cells[0].Value.ToString());
                                       if (dr3.Cells[0].Value!=null&&dr3.Cells[0].Value.ToString().Contains(macaddress))
                                           {

                                           dr3.Cells[4].Value = "連線失敗";
                                           dr3.Cells[6].Value = DateTime.Now.ToString();

                                           break;
                                           }
                                       }

                                   foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                   {
                                       Console.WriteLine("22222222222" + dr4.Cells[1].Value.ToString());
                                       if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString().Contains(macaddress))
                                       {

                                           dr4.Cells[2].Style.BackColor = Color.Red;
                                           dr4.Cells[2].Value = DateTime.Now.ToString();
                                           break;
                                       }
                                   }


                               }
                               }
                           }
                           if (listcount + 1 < PageList.Count)
                           {
                           UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                           Console.WriteLine("3333333333");
                           listcount++;
                               Console.WriteLine("error-----------");
                               Page1 mPage1 = PageList[listcount];
                               Bitmap bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                         mPage1.specification, mPage1.price, mPage1.Special_offer,
                                          mPage1.barcode, mPage1.Web, mPage1.HeadertextALL, ESLFormat);

                               pictureBoxPage1.Image = bmp;
                               mSmcEsl.TransformImageToData(bmp);

                               int numVal = Convert.ToInt32(mPage1.no) - 1;
                               button4.Visible = true;
                           dataGridView1.Rows[numVal].Selected = true;
                               aaa(datagridview1curr, true, numVal);
                               ConnectTimer.Interval = 3000;
                               ConnectTimer.Start();

                           }
                           else
                           {
                               down = false;
                               sale = false;
                               stopwatch.Stop();//碼錶停止
                               TimeSpan ts = stopwatch.Elapsed;
                               string elapsedTime = String.Format("{0:00} 分 {1:00} 秒 {2:000} ms",
                              ts.Minutes, ts.Seconds,
                              ts.Milliseconds);
                               dataGridView1.Enabled = true;
                               removeESLingstate = false;
                               onsaleESLingstate = false;
                               updateESLingstate = false;
                               datagridview1curr = 1;
                               //mSmcEsl.DisConnectBleDevice();
                               ConnectTimer.Stop();
                               int ESLBindCount = 0;
                               foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                               {

                               if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() != "未綁定") {
                                   ESLBindCount = ESLBindCount + 1;
                                   BindESL.Text = ESLBindCount.ToString();

                               }

                               }
                               mExcelData.UpdateDataList(false,"esldemoV2.xlsx", PageList);
                               testest = false;
                           mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                           mExcelData.DataGridviewSave(dataGridView1, false, openExcelAddress);
                               MessageBox.Show("全部更新完成  \r\n" + elapsedTime);
                           }
                   }
                   richTextBox1.Text = "連線失敗" + "\r\n" + richTextBox1.Text;
               }
               else if (data.Equals("Beacon更新成功"))
               {

                   if (!resetbeacon)
                   {
                       richTextBox1.Text = "Beacon更新成功" + "\r\n" + richTextBox1.Text;
                       if (beacon_index + 1 < BeaconList.Count - 1)
                       {
                           beacon_index++;
                           mSmcEsl.WriteBeaconData("ESL143AP01", BeaconList[beacon_index], false);
                       }
                       else if (beacon_index + 1 < BeaconList.Count)
                       {
                           beacon_index++;
                           mSmcEsl.WriteBeaconData("ESL143AP01", BeaconList[beacon_index], true);

                       }
                       else
                       {
                           MessageBox.Show("Beacon 全部更新完成");
                          // BeaconListNow.Clear();
                          // BeaconListNow = BeaconList;
                       }
                   }
                   else
                   {
                       Console.WriteLine("RESET CHECKED");
                       resetbeacon = false;
                   }
               }
               else if (data.Equals("Beacon更新失敗"))
               {
                   richTextBox1.Text = "Beacon更新失敗" + "\r\n" + richTextBox1.Text;

               }

               else if (data.Equals("斷線失敗"))
               {
                   if (isRun) {
                       ConnectTimer.Interval = 3000;
                       ConnectTimer.Start();
                   }
               }
               else if (data.Equals("斷線成功"))
               {
                   richTextBox1.Text = "斷線成功" + macaddress + "\r\n" + richTextBox1.Text;
                   //  Console.WriteLine("macaddress" + macaddress);
                   //if(!CheckESLOnly)
                   foreach (DataGridViewRow dr in dataGridView1.Rows)
                   {

                       Console.WriteLine("dr.Cells[2].Value.ToString()");
                       if (down || sale)
                       {
                           if (dr.Cells[12].Value != null)
                           {
                               if (dr.Cells[12].Value.ToString().Contains(macaddress))
                               {
                                   Console.WriteLine("macaddress" + macaddress);
                                   dr.Cells[4].Style.BackColor = Color.Green;
                                   dr.Cells[4].Value = DateTime.Now.ToString();
                                   dr.Cells[18].Value = DateTime.Now.ToString();
                                   PageList[listcount].UpdateState = "更新成功";
                                   PageList[listcount].UpdateTime = DateTime.Now.ToString();
                                   Page1 mPage1 = PageList[listcount];
                                   foreach (DataGridViewRow dr3 in dataGridView3.Rows)
                                   {
                                       if (dr3.Cells[0].Value.ToString().Contains(macaddress)) {
                                           dr3.Cells[4].Value = "已完成";
                                           dr3.Cells[6].Value = DateTime.Now.ToString();
                                          // UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                           break;
                                       }
                                   }

                                   foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                   {
                                       Console.WriteLine("jjjjjjj"+ dr4.Cells[1].Value.ToString() + macaddress);
                                       if (dr4.Cells[1].Value.ToString().Contains(macaddress))
                                       {
                                           Console.WriteLine("ININJ");
                                           //dr4.Cells[2].Value = DateTime.Now.ToString();
                                           // dr4.Cells[2].Style.BackColor = Color.Green;
                                           dr4.Cells[6].Value = dr.Cells[6].Value;
                                           break;
                                       }
                                   }
                               }
                           }
                       }
                       else
                       {
                           if (dr.Cells[2].Value != null)
                           {
                               if (dr.Cells[2].Value.ToString().Contains(macaddress))
                               {
                                   Console.WriteLine("macaddress" + macaddress);
                                   dr.Cells[4].Style.BackColor = Color.Green;
                                   dr.Cells[4].Value = DateTime.Now.ToString();
                                   dr.Cells[18].Value = DateTime.Now.ToString();
                                 //  UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                  PageList[listcount].UpdateState = "更新成功";
                                   PageList[listcount].UpdateTime = DateTime.Now.ToString();
                                   foreach (DataGridViewRow dr3 in dataGridView3.Rows)
                                   {
                                       if (dr3.Cells[0].Value.ToString().Contains(macaddress))
                                       {
                                           dr3.Cells[4].Value = "已完成";
                                           dr3.Cells[6].Value = DateTime.Now.ToString();
                                          /// UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                           break;
                                       }
                                   }

                                   foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                   {
                                       Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + macaddress);
                                       if (dr4.Cells[1].Value != null&& dr4.Cells[1].Value.ToString().Contains(macaddress))
                                       {
                                           Console.WriteLine("AA" + dr4.Cells[1].Value.ToString());
                                           dr4.Cells[2].Value = DateTime.Now.ToString();
                                           dr4.Cells[2].Style.BackColor = Color.Green;
                                           dr4.Cells[6].Value = dr.Cells[6].Value;
                                           break;

                                       }
                                   }

                               }

                           }
                       }



                   }

                   if (listcount + 1 < PageList.Count)
                   {
                       UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                       listcount++;
                       Page1 mPage1 = PageList[listcount];
                       if (down)
                       {

                           Bitmap bmp = mElectronicPriceData.writeIDimage(PageList[listcount].usingAddress.ToString());
                           mSmcEsl.TransformImageToData(bmp);
                           pictureBoxPage1.Image = bmp;
                           //mSmcEsl.WriteESLData(PageList[listcount].BleAddress.ToString());
                       }
                       else
                       {

                           Bitmap bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                      mPage1.barcode, mPage1.Web, mPage1.HeadertextALL, ESLFormat);

                           pictureBoxPage1.Image = bmp;
                           mSmcEsl.TransformImageToData(bmp);
                       }


                       int numVal = Convert.ToInt32(mPage1.no) - 1;
                       dataGridView1.ClearSelection();
                       dataGridView1.Rows[numVal].Cells[0].Selected = true;
                       aaa(datagridview1curr, true, numVal);


                       ConnectTimer.Interval = 3000;
                       ConnectTimer.Start();

                   }
                   else
                   {
                       Console.WriteLine("87877878787");
                       down = false;
                       sale = false;
                       stopwatch.Stop();//碼錶停止
                       TimeSpan ts = stopwatch.Elapsed;
                       string elapsedTime = String.Format("{0:00} 分 {1:00} 秒 {2:000} ms",
                      ts.Minutes, ts.Seconds,
                      ts.Milliseconds);
                       dataGridView1.Enabled = true;
                       removeESLingstate = false;
                       onsaleESLingstate = false;
                       updateESLingstate = false;
                       testest = false;
                       // mSmcEsl.DisConnectBleDevice();
                       //ConnectTimer.Stop();
                       datagridview1curr = 1;

                       int ESLBindCount = 0;
                       foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                       {

                           if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() != "未綁定")
                           {
                               ESLBindCount = ESLBindCount + 1;
                               BindESL.Text = ESLBindCount.ToString();

                           }
                       }
                           mExcelData.UpdateDataList(false, "esldemoV2.xlsx", PageList);
                       mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                       mExcelData.DataGridviewSave(dataGridView1, false, openExcelAddress);
                       ConnectTimer.Stop();
                       MessageBox.Show("全部更新完成  \r\n" + elapsedTime);
                   }
               }

               else if (data.Equals("ConnectBleTimeOut"))
               {
                   this.progressBar1.Visible = false; //隱藏進度條

                   // leconnect.Text = "連線超時...";
                   //  leconnect.ForeColor = Color.Red;
                   mSmcEsl.DisConnectBleDevice();
                   ConnectTimer.Interval = 3000;
                   ConnectTimer.Start();

               }

               else if (data.Equals("ESL_TimeOut"))
               {
                   this.progressBar1.Visible = false; //隱藏進度條
                                                      // if(listcount > 0)
                                                      //       listcount--;
                   mSmcEsl.DisConnectBleDevice();

                   Console.Write("Ble List index : " + listcount + Environment.NewLine);
               }
               else if (data.Equals("資料寫入完成完成完成完成完成完成完成完成"))
               {
                   progressBar1.Value = 0;
                   this.progressBar1.Visible = false; //顯示進度條
                   mSmcEsl.DisConnectBleDevice();






               }*/

            //掃描
            if (msgId == EslUdpTest.SmcEsl.msg_ScanDevice)
            {
                UpdateUI_Scan(data, deviceIP, battery);
            }

            // 藍牙連線
            else if (msgId == EslUdpTest.SmcEsl.msg_ConnectEslDevice)
            {
                Console.WriteLine("--------------連線中---------");
                if (status)
                {
                    str_data = "連線成功";
                    //    tbMessageBox.SelectionColor = Color.FromArgb(60, 119, 119);

                    countconnect = 0;
                   updateESLper.Text= (Convert.ToUInt32(updateESLper.Text)+1).ToString();

                   Console.WriteLine("-----------------------");
               /*    if (CheckESLOnly) {
                       //APESLState.Text = "連線成功";
                       CheckESLOnly = false;
                       mSmcEsl.DisConnectBleDevice();
                   }*/

                    ConnectTimer.Stop();
                    //ConnectBleTimeOut.Stop();    
                  //  System.Threading.Thread.Sleep(1000);
                    Page1 mPage1 = new Page1();
                    Console.WriteLine("msg_ConnectEslDevicedeviceIP:" + deviceIP);
                    for (int i = 0; i < PageList.Count; i++)
                    {

                        Console.WriteLine("msg_ConnectEslDevicedeviceIP:" + PageList[i].BleAddress+" "+PageList[i].APLink+" "+PageList[i].UpdateState);
                        if ((PageList[i].APLink+":8899") == deviceIP && PageList[i].UpdateState == null)
                        {
                            
                            mPage1 = PageList[i];
                            PageList[i].TimerSeconds.Stop();
                            PageList[i].TimerConnect.Stop();
                            break;
                        }
                        if (i == PageList.Count - 1)
                        {
                            List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).Where(p => p.UpdateState == null).ToList();
                            foreach (Page1 p in list)
                            {
                                OldRunAPList.Add(p.APLink);

                            }
                        }

                    }


                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {
                        if (kvp.Key.Contains(mPage1.APLink))
                        {
                            Console.WriteLine("ReadType" + mPage1.APLink);
                            kvp.Value.mSmcEsl.ReadBleDeviceName();
                        }
                    }

                    //     richTextBox1.Text = "連線成功"+"\r\n"+ richTextBox1.Text;
                    //  richTextBox1.ForeColor = Color.Blue;
                    /*    foreach (DataGridViewRow dr in this.dataGridView4.Rows) {
                            if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == checkESLV[listcount].ESLID) {

                                dr.Cells[0].Value = false;
                            }
                        }
                        foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                        {
                            if (kvp.Key.Contains(checkESLV[listcount].APID))
                            {

                                kvp.Value.mSmcEsl.ReadEslBattery();
                                kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                checkESLV[listcount].ESLID
                                System.Threading.Thread.Sleep(100);
                                kvp.Value.mSmcEsl.DisConnectBleDevice();
                            }
                        }*/
                    /*  Bitmap bmp;
                      foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                      {
                          if (kvp.Key.Contains(PageList[listcount].APLink))
                          {

                              if (PageList[listcount].onsale == "V")
                              {
                                  bmp = mElectronicPriceData.setPage1("Calibri", PageList[listcount].product_name, PageList[listcount].Brand,
                                          PageList[listcount].specification, PageList[listcount].price, PageList[listcount].Special_offer,
                                             PageList[listcount].barcode, PageList[listcount].Web, PageList[listcount].usingAddress, PageList[listcount].HeadertextALL, ESLSaleFormat);
                              }
                              else
                              {
                                  bmp = mElectronicPriceData.setPage1("Calibri", PageList[listcount].product_name, PageList[listcount].Brand,
                                          PageList[listcount].specification, PageList[listcount].price, PageList[listcount].Special_offer,
                                             PageList[listcount].barcode, PageList[listcount].Web, PageList[listcount].usingAddress, PageList[listcount].HeadertextALL, ESLFormat);
                              }

                              kvp.Value.mSmcEsl.TransformImageToData(bmp);
                              //  kvp.Value.mSmcEsl.WriteESLDataWithBle2(PageList[listcount].BleAddress);
                              //kvp.Value.mSmcEsl.WriteESLDataWithBle2("FFFFFFFF");
                              kvp.Value.mSmcEsl.WriteESLDataWithBle();
                              //  System.Threading.Thread.Sleep(100);
                              //  kvp.Value.mSmcEsl.DisConnectBleDevice();
                          }
                      }
                      */
                    //listcount++;

                    // DisConnectTimer.Stop();

                    // ConnectTimer.Interval = 4000;
                    // ConnectTimer.Start();
                    //  }
                }
            }
            // 藍牙斷線
            else if (msgId == EslUdpTest.SmcEsl.msg_DisconnectEslDevice)
            {
                //ConnectBleTimeOut.Stop();
                Console.WriteLine("msg_DisconnectEslDevice");
                if (status)
                {
                    str_data = "斷線成功";
                //    this.progressBar1.Visible = false; //隱藏進度條
                   // stopwatch.Reset();
                  //  stopwatch.Start();

                    DisConnectTimer.Stop();
                   System.Threading.Thread.Sleep(1000);
                    // richTextBox1.Text = "斷線成功" + "\r\n" + richTextBox1.Text;
                    //  richTextBox1.ForeColor = Color.Red;
                    //dataGridView1.Rows[1].Selected = true;
                    // aaa(1, false, 0);
                    //  if (continuewrite)
                    //  {
                    data = "斷線失敗";
              /*      if (checkV) {
                        Console.WriteLine("ISRUN");
                        CheckVTimer.Interval = 1000;
                        CheckVTimer.Start();
                    }*/
                    if (testest)
                    {
 


                        Console.WriteLine("listcount" + listcount + " PageList.Count" + PageList.Count);

                        if (listcount + 1 < PageList.Count)
                        {
                            Console.WriteLine("deviceFFFGGGGGGGG" + deviceIP);
                            listcount++;
                            bleConnect(deviceIP);

                        }
                        else
                        {

                            if (sale)
                            {
                                bool TT = false;
                                foreach (DataGridViewRow dr in dataGridView1.Rows)
                                {
                                    if (dr.Cells[5].Style.ForeColor == Color.Red || dr.Cells[6].Style.ForeColor == Color.Red || dr.Cells[7].Style.ForeColor == Color.Red || dr.Cells[8].Style.ForeColor == Color.Red || dr.Cells[9].Style.ForeColor == Color.Red || dr.Cells[10].Style.ForeColor == Color.Red || dr.Cells[11].Style.ForeColor == Color.Red)
                                    {
                                        TT = true;
                                    }
                                }
                                if (!TT)
                                {
                                    button2.Enabled = false;
                                    button2.BackColor = Color.Gray;

                                }


                            }

                            if (EslStyleChangeUpdate)
                            {
                                button19.BackColor = Color.Gray;
                                button19.Enabled = false;
                            }

                            /*    if (!down && !sale && !reset && !saletime)
                                {
                                    foreach (DataGridViewRow dr in dataGridView1.Rows)
                                    {
                                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                                        {
                                            dr.Cells[0].Value = true;
                                            dr.Cells[0].ReadOnly = false;

                                        }
                                        else
                                        {
                                            dr.Cells[0].Value = false;
                                            dr.Cells[0].ReadOnly = true;
                                        }
                                        dr.Cells[1].Value = dr.Cells[12].Value;
                                    }

                                }*/
                            testest = false;
                            onlockedbutton(testest);
                            checkClick = false;
                            down = false;
                            OldRunAPList.Clear();
                            backESLList.Clear();
                            Console.WriteLine("87877878787");
                            progressBar1.Visible = false;
                            down = false;
                            sale = false;
                            reset = false;
                            saletime = false;
                            immediateUpdate = false;
                            listcount = 0;
                            EslStyleChangeUpdate = false;

                            pictureBoxPage1.Image = null;
                            richTextBox1.Text = "";
                            stopwatch.Stop();//碼錶停止
                            TimeSpan tsa = stopwatch.Elapsed;
                            string elapsedTimea = String.Format("{0:00} 分 {1:00} 秒 {2:000} ms",
                            tsa.Minutes, tsa.Seconds,
                            tsa.Milliseconds);
                            dataGridView1.Enabled = true;
                            removeESLingstate = false;
                            onsaleESLingstate = false;
                            updateESLingstate = false;
                            ESLStyleDataChange = false;
                            ESLSaleStyleDataChange = false;
                            PageList.Clear();
                            testest = false;

                            //pictureBox4.Visible = false;
                            //   checkClick = false;
                            //  OldRunAPList.RemoveAll(it => true);
                            OldRunAPList.Clear();
                            // mSmcEsl.DisConnectBleDevice();
                            //ConnectTimer.Stop();
                            CheckBeaconTimer.Start();
                            int ESLBindCount = 0;
                            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                            {

                                if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() != "未綁定")
                                {
                                    ESLBindCount = ESLBindCount + 1;
                                    BindESL.Text = ESLBindCount.ToString();

                                }
                            }


                            ConnectTimer.Stop();
                            if (!radioButton1.Checked)
                            {
                                MessageBox.Show("全部更新完成  \r\n" + elapsedTimea);
                            }
                            //    mExcelData.UpdateDataList(false, "esldemoV2.xlsx", PageList);
                            //      mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                            //     mExcelData.DataGridviewSave(dataGridView1, false, openExcelAddress);

                            datagridview1curr = 2;
                            aaa(1, false, 0);

                        }
                        totalwritecount++;
                        // ConnectTimer.Interval = 2000;
                        //ConnectTimer.Start();
                        //  }
                    }


                }
                else
                {
                    //richTextBox1.Text = "斷線失敗";
                    str_data = "斷線失敗";
                  
                }
            }

            else if (msgId == EslUdpTest.SmcEsl.msg_ReadEslType)
            {
                Console.WriteLine(" SmcEsl.msg_ReadEslType" + data);

                
                string EslSize =  data.Substring(6, 2);
                if (status)
                {
                
                    if (data.Substring(6, 2).Equals("00"))
                    {
                        str_data = "2.13吋";
                    }
                    else if (data.Substring(6, 2).Equals("01"))
                    {
                        str_data = "2.9吋";
                    }
                    else if (data.Substring(6, 2).Equals("02"))
                    {
                        str_data = "4.2吋";
                    }

                    writeESLdataBySzie(deviceIP,EslSize,true);
               

                }
                else
                {
                    str_data = "讀取尺寸失敗";
                }

            }

            // 取得藍牙名稱
            else if (msgId == EslUdpTest.SmcEsl.msg_ReadEslName)
            {
                str_data = "Device Name :" + data;
                Console.WriteLine("msg_ReadEslName:" + data);
                if (!data.Contains("ESL-0003"))
                {
                    Console.WriteLine("NOC---"+ data);
                    writeESLdataBySzie(deviceIP, "00", false);
                }
                else
                {
                    Console.WriteLine("C---" + data);
                    Page1 mPage1 = new Page1();
                    for (int i = 0; i < PageList.Count; i++)
                    {
                        if ((PageList[i].APLink + ":8899") == deviceIP && PageList[i].UpdateState == null)
                        {

                            mPage1 = PageList[i];
                            break;
                        }
                        if (i == PageList.Count - 1)
                        {
                            List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).Where(p => p.UpdateState == null).ToList();
                            foreach (Page1 p in list)
                            {
                                OldRunAPList.Add(p.APLink);

                            }
                        }

                    }

                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {
                        if (kvp.Key.Contains(mPage1.APLink))
                        {
                            Console.WriteLine("ReadType" + mPage1.APLink);
                            kvp.Value.mSmcEsl.ReadEslType();
                        }
                    }
                }


            }
            // 寫入設備名稱
            else if (msgId == EslUdpTest.SmcEsl.msg_WriteEslName)
            {
                if (status)
                {
                    str_data = "Esl 名稱更新成功";
                }
                else
                {
                    str_data = "Esl 名稱更新失敗";
                }
            }
            // 寫入ESL資料
            else if (msgId == EslUdpTest.SmcEsl.msg_WriteEslData)
            {
                if (status)
                {
                    Console.WriteLine("資料寫入成功");
                    //progressBar1.Value += progressBar1.Step;
                   // str_data = "資料寫入成功";
                }
                else
                {
                    str_data = "資料寫入失敗";
                }
            }

            // 寫入ESL資料
            else if (msgId == EslUdpTest.SmcEsl.msg_WriteEslData2)
            {
                if (status)
                {
                    Console.WriteLine("資料寫入成功");
                    //progressBar1.Value += progressBar1.Step;
                  //  str_data = "資料寫入成功";
                }
                else
                {
                    str_data = "資料寫入失敗";
                }
            }
            // 寫入ESL資料，全部寫完
            else if (msgId == EslUdpTest.SmcEsl.msg_WriteEslDataFinish)
            {
                string reCount = (e as EslUdpTest.SmcEsl.SMCEslReceiveEventArgs).Re;
                Console.WriteLine(listcount+"---------msg_WriteEslDataFinish---------"+ reCount);

             //   this.progressBar1.Visible = false;
                str_data = "全部資料寫入完成";
                
                str_data = writeESLsuccess(deviceIP, str_data);
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    if (kvp.Key.Contains(deviceIP.Split(':')[0]))
                    {
                        kvp.Value.mSmcEsl.DisConnectBleDevice();
                    }
                }
            }


            // 寫入ESL資料，全部寫完
            else if (msgId == EslUdpTest.SmcEsl.msg_WriteEslDataFinish2)
            {
                string reCount = (e as EslUdpTest.SmcEsl.SMCEslReceiveEventArgs).Re;
                Console.WriteLine(listcount + "---------msg_WriteEslDataFinish2---------" + reCount);
                BleWriteTimer.Stop();
                str_data = "全部資料寫入完成";
                //  this.progressBar1.Visible = false;
                str_data = writeESLsuccess(deviceIP, str_data);
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    if (kvp.Key.Contains(deviceIP.Split(':')[0]))
                    {
                        kvp.Value.mSmcEsl.DisConnectBleDevice();
                    }
                }

  
            }
            // 寫入AP Beacon Data
            else if (msgId == EslUdpTest.SmcEsl.msg_WriteBeacon)
            {
                Console.WriteLine("WWWWWWWWWWWWWWWWWWWWWWW");
                if (!resetbeacon)
                {
                    Console.WriteLine(BeaconList.Count+"成功次數"+ beacon_index);
                    str_data = "Beacon更新成功"+ BeaconList[beacon_index].BeaconProduct;
                    richTextBox1.Text = richTextBox1.Text+"Beacon更新成功" + "\r\n" ;
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {
                        if (dr.Cells[5].Value != null && dr.Cells[5].Value.ToString() == BeaconList[beacon_index].BeaconProduct)
                        {
                            
                            productState(dr);
                        }
                    }
                    if (beacon_index + 1 < BeaconList.Count - 1)
                    {
                        beacon_index++;
                        //setLocalTime(BeaconList[beacon_index].APID);
                        Console.WriteLine("還有Beacon");

                        foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                        {
                            if (kvp.Key.Contains(BeaconList[beacon_index].APID))
                            {
                                setBeaconTime(BeaconList[beacon_index].APID);
                                //System.Threading.Thread.Sleep(100);
                                kvp.Value.mSmcEsl.WriteBeaconData("ESL143AP01", BeaconList[beacon_index].BeaconProduct, false);
                                
                            }
                        }
                        
                    }
                    else if (beacon_index + 1 < BeaconList.Count)
                    {
                        beacon_index++;
                       // setLocalTime(BeaconList[beacon_index].APID);
                          Console.WriteLine("lastBeacon");
                        foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                        {
                            if (kvp.Key.Contains(BeaconList[beacon_index].APID))
                            {
                                setBeaconTime(BeaconList[beacon_index].APID);
                              //  System.Threading.Thread.Sleep(100);
                                kvp.Value.mSmcEsl.WriteBeaconData("ESL143AP01", BeaconList[beacon_index].BeaconProduct, true);
                            }
                        }
                        

                    }
                    else
                    {
                        MessageBox.Show("Beacon 全部更新完成");
                        testest = false;
                        datagridview1curr = 2;
                        aaa(1, false, 0);
                        // BeaconListNow.Clear();
                        // BeaconListNow = BeaconList;
                    }
                }
                else
                {
                    Console.WriteLine("RESET CHECKED");
                    resetbeacon = false;
                }

                /*   if (status)
                   {
                       str_data = "Beacon更新成功";
                       // tbBeaconCount
                       BeaconIndex++;
                       string EID = "000000000000";

                       if (BeaconIndex < 10)
                       {
                           EID = "00000000000" + BeaconIndex;
                       }
                       else if (BeaconIndex > 10 && BeaconIndex < 100)
                       {
                           EID = "0000000000" + BeaconIndex;
                       }
                       else
                       {
                           EID = "000000000" + BeaconIndex;
                       }

                       if (BeaconIndex < int.Parse(tbBeaconCount.Text.ToString()) - 1)
                       {
                           //mSmcEsl.WriteBeaconData("ESL143AP01", EID, false);
                           foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                           {
                               if (kvp.Key.Contains(AP_IP_Label.Text))
                               {
                                   kvp.Value.mSmcEsl.WriteBeaconData("ESL143AP01", EID, false);
                               }
                           }
                       }
                       else if (BeaconIndex == int.Parse(tbBeaconCount.Text.ToString()) - 1)
                       {
                           // mSmcEsl.WriteBeaconData("ESL143AP01", EID, true);
                           foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                           {
                               if (kvp.Key.Contains(AP_IP_Label.Text))
                               {
                                   kvp.Value.mSmcEsl.WriteBeaconData("ESL143AP01", EID, true);
                               }
                           }
                       }
                   }
                   else
                   {
                       str_data = "Beacon更新失敗";
                   }*/
            }

            // ---------  ESL  版本 -------
            else if (msgId == EslUdpTest.SmcEsl.msg_ReadEslVersion)
            {
                if (status)
                {
                    str_data = "ESL 版本 = " + data;
                }
                else
                {
                    str_data = "ESL 版本讀取錯誤";
                }
            }
            // ---------  ESL  電壓 -------
            else if (msgId == EslUdpTest.SmcEsl.msg_ReadEslBattery)
            {
                if (status)
                {
                    str_data = "Esl電池電壓 = " + data + " V";
                    foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                    {
                        if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString() == macaddress)
                        {
                            dr4.Cells[5].Value = data;
                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 5, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                        }
                    }
                }
                else
                {
                    str_data = "Esl電池電壓讀取失敗";
                }
            }
            // ---------  ESL  製造資料 -------
            else if (msgId == EslUdpTest.SmcEsl.msg_ReadManufactureData)
            {
                if (status)
                {
                    str_data = "製造資料 = " + data;
                }
                else
                {
                    str_data = "製造資料讀取錯誤";
                }
            }
            // ---------  ESL  版本 -------
            else if (msgId == EslUdpTest.SmcEsl.msg_WriteManufactureData)
            {
                if (status)
                {
                    str_data = "ESL 版本寫入成功";
                   

                }
                else
                {
                    str_data = "ESL 版本寫入失敗";
                }
            }

            // ---------  AP  寫入Buffer -------
            else if (msgId == EslUdpTest.SmcEsl.msg_WriteESLDataBuffer)
            {
                if (status)
                {
                   // str_data = "寫入 AP Buffer 完成";
                    Console.WriteLine("寫入 AP Buffer 完成" + deviceIP);
                    for (int i = 0; i < PageList.Count; i++) {
                        Console.WriteLine(PageList[i].APLink + "=" + deviceIP+"and"+ PageList[i].UpdateState);
                        if (PageList[i].APLink+":8899" == deviceIP && PageList[i].UpdateState == null) {
                            Page1 mPage1 = PageList[i];
                            str_data = " 寫入AP Buffer 完成"+ mPage1.usingAddress+"連線";
                            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                            {
                                Console.WriteLine("mPage1.APLink" + mPage1.APLink);
                                Console.WriteLine("mPage1.usingAddress" + mPage1.usingAddress);
                                if (kvp.Key.Contains(mPage1.APLink))
                                {

                                    EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                    mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3,0);
                                    //Console.WriteLine("WWWWWWWWWWWWWWWWWWWWWWWTTTTTTTTT");
                                }
                            }
                            break;
                        }
                    }
                    
                }
                else
                {
                    str_data = "寫入 AP Buffer失敗";
                    Console.Write(str_data);
                }
            }

            // ---------  AP  更新ESL -------
            else if (msgId == EslUdpTest.SmcEsl.msg_UpdataESLDataFromBuffer)
            {
                if (status)
                {
                    //ConnectBleTimeOut.Stop();
                    //str_data = "AP 更新 ESL 完成";
                    updateESLper.Text = (Convert.ToUInt32(updateESLper.Text) + 1).ToString();
                    //  richTextBox1.Text = "斷線成功" + macaddress + "\r\n" + richTextBox1.Text;


                    Console.WriteLine("AP 更新 PageList.Count" + PageList.Count);

                    for (int i = 0; i < PageList.Count; i++)
                    {
                      //  Console.WriteLine(i+"AP 更新 ESL" + PageList[i].usingAddress + deviceIP);
                        if (PageList[i].APLink + ":8899" == deviceIP&& PageList[i].UpdateState==null)
                        {
                            str_data = "AP 更新" + PageList[i].usingAddress + "完成";
                            Console.WriteLine(i + "AP 更新 ESL" + PageList[i].usingAddress + deviceIP);
                            if (PageList[i].actionName=="reset")// 地3業還原
                            {
                                PageList[i].UpdateState = "更新成功";
                                PageList[i].UpdateTime = DateTime.Now.ToString();
                                foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                {
                                    //    Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + macaddress);
                                    if (dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                    {
                                        //  Console.WriteLine("ININJ");
                                        dr4.Cells[2].Value = DateTime.Now.ToString();
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        dr4.Cells[2].Style.BackColor = Color.Green;
                                        dataGridView4.Rows[dr4.Index].Cells[0].Selected = false;
                                        dr4.Cells[0].Value = false;
                                        dr4.Cells[3].Value = "";
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        dr4.Cells[6].Value = "未绑定";
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        //    dr4.Cells[6].Value = dr.Cells[6].Value;
                                        break;
                                    }
                                }

                                foreach (DataGridViewRow dr in dataGridView1.Rows) {

                                    if (dr.Cells[1].Value != null&&dr.Cells[1].Value.ToString()!="")
                                    {
                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                    {
                                        Console.WriteLine("=========================");
                                        if (dr.Cells[1].Value.ToString().Contains(',' + PageList[i].usingAddress))
                                        {
                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + PageList[i].usingAddress);
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                            dr.Cells[12].Value = dr.Cells[1].Value;
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        }
                                        if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress + ','))
                                        {

                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress + ',');
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                            dr.Cells[12].Value = dr.Cells[1].Value;
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        }

                                        if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                        {

                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress);
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                                            dr.Cells[12].Value = dr.Cells[1].Value;
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        }

                                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString()=="")
                                        {
                                            MessageBox.Show(dr.Cells[6].Value.ToString()+"無綁定ESL自動下架");
                                            dr.Cells[0].ReadOnly = true;
                                            dr.Cells[0].Value = false;
                                            dr.Cells[13].Value = "";
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.DefaultCellStyle.ForeColor = Color.Gray;
                                        }
                                            break;
                                        }
                                    }
                                    Console.WriteLine("=OOOOOOOOOO===");
                                }
                            }
                            else
                            {//第一頁功能

                                PageList[i].UpdateState = "更新成功";
                                PageList[i].UpdateTime = DateTime.Now.ToString();
                                List<string> nullbeacon = new List<string>();
                                foreach (DataGridViewRow dr in dataGridView1.Rows)
                                {

                                   // Console.WriteLine("dr.Cells[1].Value.ToString()");
                                    if (PageList[i].actionName == "down" || PageList[i].actionName == "sale" || PageList[i].actionName == "saletime" || PageList[i].actionName == "EslStyleChangeUpdate")
                                    {



                                        if (dr.Cells[1].Value != null)
                                        {
                                            if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                            {
                                                //// Console.WriteLine("macaddress" + macaddress);
                                                // dr.Cells[4].Style.BackColor = Color.Green;
                                                //  dr.Cells[4].Value = DateTime.Now.ToString();
                                                // Console.WriteLine("macaddress" + dr.Cells[6].Value);
                                                // dr.Cells[18].Value = DateTime.Now.ToString();

                                                dr.Cells[4].Style.BackColor = Color.Green;
                                                dr.Cells[4].Value = DateTime.Now.ToString();
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1,4, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                
                                               /* PageList[i].UpdateState = "更新成功";
                                                PageList[i].UpdateTime = DateTime.Now.ToString();*/
                                                dataGridView1.Rows[dr.Index].Cells[0].Selected = false;
                                                Page1 mPage1 = PageList[i];
                                                if (PageList[i].actionName == "down")
                                                {

                                                    foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                                    {
                                                       //  Console.WriteLine(dr4.Cells[1].Value.ToString()+"jjjjjjj" + PageList[i].usingAddress);
                                                        if (dr4.Cells[1].Value!=null&&dr4.Cells[1].Value.ToString()==PageList[i].usingAddress)
                                                        {
//Console.WriteLine("ININJ");
                                                            dr4.Cells[2].Value = DateTime.Now.ToString();
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            dr4.Cells[2].Style.BackColor = Color.Green;
                                                            dataGridView4.Rows[dr4.Index].Cells[0].Selected = false;
                                                            dr4.Cells[0].Value = false;
                                                            dr4.Cells[3].Value = "";
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            dr4.Cells[6].Value = "未綁定";
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            //    dr4.Cells[6].Value = dr.Cells[6].Value;
                                                            break;
                                                        }
                                                    }




                                                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                                                        {
                                                            if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                            {
                                                                Console.WriteLine(PageList[i].usingAddress+"=========================" + dr.Cells[1].Value.ToString());
                                                                if (dr.Cells[1].Value.ToString().Contains(',' + PageList[i].usingAddress))
                                                                {
                                                                    int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + PageList[i].usingAddress);
                                                                    dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                                                    dr.Cells[12].Value = dr.Cells[1].Value;
                                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                }
                                                                if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress + ','))
                                                                {

                                                                    int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress + ',');
                                                                    dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                                                    dr.Cells[12].Value = dr.Cells[1].Value;
                                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                }

                                                                if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                                {

                                                                    int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress);
                                                                    dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                                                                    dr.Cells[12].Value = dr.Cells[1].Value;
                                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                }

                                                                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == "")
                                                                {
                                                                    dr.Cells[0].ReadOnly = true;
                                                                    dr.Cells[0].Value = false;
                                                                    dr.Cells[2].Value = DBNull.Value;
                                                                    dr.Cells[13].Value = "";
                                                                    dr.Cells[19].Value = DBNull.Value;
                                                                    dr.Cells[20].Value = DBNull.Value;
                                                                    dr.Cells[21].Value = DBNull.Value;
                                                                    dr.Cells[22].Value = DBNull.Value;
                                                                dr.Cells[23].Value = DBNull.Value;

                                                                nullbeacon.Add(dr.Cells[5].Value.ToString());
                                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 21, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 22, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 23, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                                dr.DefaultCellStyle.ForeColor = Color.Gray;
                                                                }

                                                            }
                                                        }

                                                    // dr.DefaultCellStyle.ForeColor = Color.Gray;
                                                    dr.Cells[15].Value = "X";
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr.Cells[16].Value = "X";
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 16, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    // dr.Cells[1].Value = "";
                                                    //  mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr.Cells[18].Value = DateTime.Now.ToString();
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr.Cells[13].Value = "";
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                   // dr.Cells[12].Value = "";
                                                  //  mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                   // dr.Cells[0].ReadOnly = true;
                                                    //dr.Cells[0].Value = false;

                                                }

                                             /*   if (immediateUpdate)
                                                {
                                                    dr.Cells[13].Value = PageList[i].ProductStyle;
                                                    Console.WriteLine(PageList[i].product_name + "PageList[i].ProductStyle" + PageList[i].ProductStyle);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    productState(dr);
                                                    dr.Cells[18].Value = DateTime.Now.ToString();
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);

                                                    foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                                    {
                                                        Console.WriteLine(dr4.Cells[1].Value.ToString() + "jjjjjjj" + PageList[i].usingAddress);
                                                        if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString() == PageList[i].usingAddress)
                                                        {
                                                            Console.WriteLine("ININJ");
                                                            dr4.Cells[2].Value = DateTime.Now.ToString();
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            dr4.Cells[2].Style.BackColor = Color.Green;
                                                            dr4.Cells[3].Value = dr.Cells[13].Value;
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            dr4.Cells[6].Value = dr.Cells[6].Value;
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            break;
                                                        }
                                                    }

                                                }*/

                                                
                                                if (PageList[i].actionName == "sale")
                                                {
                                                    if(PageList[i].onsale=="V")
                                                        dr.Cells[13].Value = styleSaleName;
                                                    else
                                                        dr.Cells[13].Value = styleName;
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 5, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 6, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 7, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 8, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 9, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 10, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 11, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr.Cells[18].Value = DateTime.Now.ToString();
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    
                                                    dr.Cells[5].Style.ForeColor = Color.Empty;
                                                    dr.Cells[6].Style.ForeColor = Color.Empty;
                                                    dr.Cells[7].Style.ForeColor = Color.Empty;
                                                    dr.Cells[8].Style.ForeColor = Color.Empty;
                                                    dr.Cells[9].Style.ForeColor = Color.Empty;
                                                    dr.Cells[10].Style.ForeColor = Color.Empty;
                                                    dr.Cells[11].Style.ForeColor = Color.Empty;
                                                    dr.DefaultCellStyle.ForeColor = Color.Black;
                                                    foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                                    {
                                                       // Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + PageList[i].usingAddress);
                                                        if (dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                        {
                                                            Console.WriteLine("ININJ");
                                                            dr4.Cells[2].Value = DateTime.Now.ToString();
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            dr4.Cells[2].Style.BackColor = Color.Green;
                                                            dr4.Cells[3].Value = styleName;
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            dr4.Cells[6].Value = dr.Cells[6].Value;
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            break;
                                                        }
                                                    }

                                                }
                       
                                                if (PageList[i].actionName == "saletime")
                                                {
                                                    dr.Cells[13].Value = PageList[i].ProductStyle;
                                                    Console.WriteLine(PageList[i].product_name+ "PageList[i].ProductStyle"+ PageList[i].ProductStyle);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    productState(dr);
                                                    dr.Cells[18].Value = DateTime.Now.ToString();
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr.Cells[3].Value = false;



                                                    foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                                    {
                                                    //    Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + PageList[i].usingAddress);
                                                        if (dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                        {
                                                         //   Console.WriteLine("ININJ");
                                                            dr4.Cells[2].Value = DateTime.Now.ToString();
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            dr4.Cells[2].Style.BackColor = Color.Green;
                                                            dr4.Cells[3].Value = dr.Cells[13].Value;
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            dr4.Cells[6].Value = dr.Cells[6].Value;
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            break;
                                                        }
                                                    }
                                                }

                                                if (PageList[i].actionName == "EslStyleChangeUpdate")
                                                {
                                                    dr.Cells[13].Value = PageList[i].ProductStyle;
                                                    Console.WriteLine(PageList[i].product_name + "PageList[i].ProductStyle" + PageList[i].ProductStyle);
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr.Cells[18].Value = DateTime.Now.ToString();
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                                    {
                                                     //   Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + PageList[i].usingAddress);
                                                        if (dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                        {
                                                        //    Console.WriteLine("ININJ");
                                                            dr4.Cells[2].Value = DateTime.Now.ToString();
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            dr4.Cells[2].Style.BackColor = Color.Green;
                                                            dr4.Cells[3].Value = dr.Cells[13].Value;
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            dr4.Cells[6].Value = dr.Cells[6].Value;
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                            break;
                                                        }
                                                    }
                                                }


                                                foreach (DataGridViewRow dr3 in dataGridView3.Rows)
                                                {
                                                    if (dr3.Cells[0].Value.ToString().Contains(PageList[i].usingAddress))
                                                    {
                                                        dr3.Cells[4].Value = "已完成";
                                                        dr3.Cells[6].Value = DateTime.Now.ToString();
                                                        // UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                                        break;
                                                    }
                                                }

                                             /*   foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                                {
                                                    Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + PageList[i].usingAddress);
                                                    if (dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                    {
                                                        Console.WriteLine("ININJ");
                                                        dr4.Cells[2].Value = DateTime.Now.ToString();
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4,2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        dr4.Cells[2].Style.BackColor = Color.Green;
                                                        dr4.Cells[3].Value = styleName;
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        dr4.Cells[6].Value = dr.Cells[6].Value;
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        break;
                                                    }
                                                }*/
                                            }
                                        }



                                    }
                                    else
                                    {
                                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString().Contains(PageList[i].BleAddress))
                                        {
                                            foreach (DataGridViewRow drnew in dataGridView1.Rows)
                                            {
                                                if(drnew.Cells[12].Value!=null&& drnew.Cells[12].Value.ToString()!="")
                                                {
                                                    if (drnew.Cells[12].Value.ToString().Contains(PageList[i].BleAddress))
                                                    {
                                                        Console.WriteLine("cccccccc" + drnew.Cells[12].Value);
                                                        if (drnew.Cells[12].Value != null)
                                                            a = drnew.Cells[12].Value;
                                                        if (drnew.Cells[13].Value != null)
                                                            b = drnew.Cells[13].Value;
                                                        if (drnew.Cells[14].Value != null)
                                                            c = drnew.Cells[14].Value;
                                                        if (drnew.Cells[15].Value != null)
                                                            d = drnew.Cells[15].Value;
                                                        easd = drnew.Cells[16].Value;
                                                        //  drnew.Cells[12].Value = DBNull.Value;
                                                        if (drnew.Cells[12].Value.ToString().Contains(',' + PageList[i].usingAddress))
                                                        {
                                                            int changeaddr = drnew.Cells[12].Value.ToString().IndexOf(',' + PageList[i].usingAddress);
                                                            drnew.Cells[12].Value = drnew.Cells[12].Value.ToString().Remove(changeaddr, 13);
                                                            

                                                        }
                                                        if (drnew.Cells[12].Value.ToString().Contains(PageList[i].usingAddress + ','))
                                                        {

                                                            int changeaddr = drnew.Cells[12].Value.ToString().IndexOf(PageList[i].usingAddress + ',');
                                                            drnew.Cells[12].Value = drnew.Cells[12].Value.ToString().Remove(changeaddr, 13);

                                                        }

                                                        if (drnew.Cells[12].Value.ToString().Contains(PageList[i].usingAddress))
                                                        {

                                                            int changeaddr = drnew.Cells[12].Value.ToString().IndexOf(PageList[i].usingAddress);
                                                            drnew.Cells[12].Value = drnew.Cells[12].Value.ToString().Remove(changeaddr, 12);

                                                        }
                                                        if (drnew.Cells[12].Value.ToString().Length == 0)
                                                        {
                                                            drnew.DefaultCellStyle.ForeColor = Color.Gray;
                                                            drnew.Cells[0].Value = false;
                                                            drnew.Cells[0].ReadOnly = true;
                                                        }
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        drnew.Cells[13].Value = DBNull.Value;
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        drnew.Cells[14].Value = DBNull.Value;
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 14, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        if (drnew.Cells[12].Value != null && drnew.Cells[12].Value.ToString() == "")
                                                        {
                                                            drnew.Cells[16].Value = "X";
                                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 16, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);

                                                           // drnew.Cells[15].Value = "X";
                                                       //     mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        }
                                                        break;
                                                    }
                                                   
                                                }

                                            }

                                            Console.WriteLine("b" + b.ToString());
                                            Console.WriteLine("c" + c.ToString());
                                            Console.WriteLine("d" + d.ToString());
                                           // dr.Cells[12].Value = dr.Cells[1].Value;
                                            if (dr.Cells[12].Value.ToString().Length>0)
                                                dr.Cells[12].Value = dr.Cells[12].Value.ToString() + "," + PageList[i].usingAddress;
                                            else
                                                dr.Cells[12].Value = PageList[i].usingAddress;

                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.Cells[13].Value = b;
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.Cells[14].Value = c;
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 14, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                           /* if (dr.Cells[12].Value != null && dr.Cells[12].Value.ToString().Length == 12) {
                                                dr.Cells[15].Value = "X";
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            }*/
                                            dr.Cells[16].Value = "V";
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 16, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);

                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 5, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 6, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 7, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 8, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 9, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 10, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 11, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.DefaultCellStyle.ForeColor = Color.Black;
                                            dr.Cells[1].Style.ForeColor = Color.Black;
                                            dr.Cells[0].ReadOnly = false;
                                            dr.Cells[0].Value = true;
                                            /*   foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                                               {
                                                   if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString() == dr.Cells[12].Value.ToString())
                                                   {
                                                       dr4.Cells[6].Value = dr.Cells[6].Value;
                                                   }
                                               }*/
                                        }
                                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                                        {
                                            Console.WriteLine("macaddress" + dr.Cells[1].Value.ToString());
                                            if (dr.Cells[1].Value.ToString().Contains(PageList[i].BleAddress))
                                            {
                                                Console.WriteLine("macaddress" + PageList[i].BleAddress);
                                                dr.Cells[4].Style.BackColor = Color.Green;
                                                dr.Cells[4].Value = DateTime.Now.ToString();
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 4, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                dr.Cells[18].Value = DateTime.Now.ToString();
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                dr.Cells[13].Value = PageList[i].ProductStyle;
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                if (dr.Index % 2 == 1)
                                                {
                                                    dr.DefaultCellStyle.BackColor = Color.Beige;
                                                }
                                                else
                                                {
                                                    dr.DefaultCellStyle.BackColor = Color.Bisque;
                                                }
                                                //  UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                                PageList[i].UpdateState = "更新成功";
                                                PageList[i].UpdateTime = DateTime.Now.ToString();
                                                foreach (DataGridViewRow dr3 in dataGridView3.Rows)
                                                {
                                                    if (dr3.Cells[0].Value.ToString().Contains(PageList[i].BleAddress))
                                                    {
                                                        dr3.Cells[4].Value = "已完成";
                                                        dr3.Cells[6].Value = DateTime.Now.ToString();
                                                        /// UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                                        break;
                                                    }
                                                }

                                                foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                                {
                                                   // Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + PageList[i].BleAddress);
                                                    if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString().Contains(PageList[i].BleAddress))
                                                    {
                                                       // Console.WriteLine("AA" + dr4.Cells[1].Value.ToString());
                                                        dr4.Cells[2].Value = DateTime.Now.ToString();
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        dr4.Cells[2].Style.BackColor = Color.Green;
                                                        dr4.Cells[3].Value = styleName;
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        dr4.Cells[6].Value = dr.Cells[6].Value;
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        break;

                                                    }
                                                }

                                            }

                                        }
                                    }



                                }

                                if (nullbeacon.Count != 0)
                                    beacon_data_set(nullbeacon, "", "", "");

                            }
                            deviceIPData = PageList[i].APLink;
                           // PageList.RemoveAt(i);
                            Console.WriteLine("aaaaaaa");
                            break;
                        }
                    }


                    Console.WriteLine("listcount" + listcount+ " PageList.Count" + PageList.Count);

                    if (listcount + 1 < PageList.Count)
                    {
                        Console.WriteLine("deviceFFFGGGGGGGG"+ deviceIP);
                       listcount++;
                        bufferConnect(deviceIP);

                    }
                    else
                    {

                        if (sale)
                        {
                            bool TT = false;
                            foreach (DataGridViewRow dr in dataGridView1.Rows)
                            {
                                if (dr.Cells[5].Style.ForeColor == Color.Red || dr.Cells[6].Style.ForeColor == Color.Red || dr.Cells[7].Style.ForeColor == Color.Red || dr.Cells[8].Style.ForeColor == Color.Red || dr.Cells[9].Style.ForeColor == Color.Red || dr.Cells[10].Style.ForeColor == Color.Red || dr.Cells[11].Style.ForeColor == Color.Red)
                                {
                                    TT = true;
                                }
                            }
                            if (!TT)
                            {
                                button2.Enabled = false;
                                button2.BackColor = Color.Gray;

                            }
                              
                           
                        }

                        if (EslStyleChangeUpdate)
                        {
                            button19.BackColor = Color.Gray;
                            button19.Enabled = false;
                        }

                    /*    if (!down && !sale && !reset && !saletime)
                        {
                            foreach (DataGridViewRow dr in dataGridView1.Rows)
                            {
                                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                                {
                                    dr.Cells[0].Value = true;
                                    dr.Cells[0].ReadOnly = false;
                                    
                                }
                                else
                                {
                                    dr.Cells[0].Value = false;
                                    dr.Cells[0].ReadOnly = true;
                                }
                                dr.Cells[1].Value = dr.Cells[12].Value;
                            }

                        }*/
                            testest = false;
                            onlockedbutton(testest);
                            checkClick = false;
                            down = false;
                            OldRunAPList.Clear();
                            backESLList.Clear();
                            Console.WriteLine("87877878787");
                            progressBar1.Visible = false;
                            down = false;
                            sale = false;
                            reset = false;
                            saletime = false;
                            immediateUpdate = false;
                            listcount = 0;
                            EslStyleChangeUpdate = false;

                            pictureBoxPage1.Image = null;
                            richTextBox1.Text = "";
                            stopwatch.Stop();//碼錶停止
                            TimeSpan ts = stopwatch.Elapsed;
                            string elapsedTime = String.Format("{0:00} 分 {1:00} 秒 {2:000} ms",
                            ts.Minutes, ts.Seconds,
                            ts.Milliseconds);
                            dataGridView1.Enabled = true;
                            removeESLingstate = false;
                            onsaleESLingstate = false;
                            updateESLingstate = false;
                            ESLStyleDataChange = false;
                            ESLSaleStyleDataChange = false;
                            PageList.Clear();
                            testest = false;

                            //pictureBox4.Visible = false;
                        //   checkClick = false;
                        //  OldRunAPList.RemoveAll(it => true);
                            OldRunAPList.Clear();
                            // mSmcEsl.DisConnectBleDevice();
                            //ConnectTimer.Stop();
                            CheckBeaconTimer.Start();
                            int ESLBindCount = 0;
                            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                            {

                                if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() != "未綁定")
                                {
                                    ESLBindCount = ESLBindCount + 1;
                                    BindESL.Text = ESLBindCount.ToString();

                                }
                            }


                            ConnectTimer.Stop();
                        if (!radioButton1.Checked)
                        {
                            MessageBox.Show("全部更新完成  \r\n" + elapsedTime);
                        }
                            //    mExcelData.UpdateDataList(false, "esldemoV2.xlsx", PageList);
                            //      mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                            //     mExcelData.DataGridviewSave(dataGridView1, false, openExcelAddress);
                          
                            datagridview1curr = 2;
                            aaa(1, false, 0);
                       
                    }
                }
                else
                {
                    
                    Console.WriteLine("~~~~~~~~~~");
                    Runtime = true;
          /*          if (countconnect < 2)
                    {
                        if (CheckESLOnly)
                        {
                            Console.WriteLine("CheckESLOnly");
                            APESLState.Text = "重新嘗試" + countconnect + "次";
                            //1/2新年第一天上工摟~~~
                            ConnectTimer.Interval = 1000;
                            ConnectTimer.Start();
                            countconnect++;

                        }
                        else
                        {
                            Console.WriteLine("ERROR" + countconnect);
                            ConnectTimer.Interval = 1000;
                            ConnectTimer.Start();
                            countconnect++;
                        }

                    }
                    else
                    {*/

                        countconnect = 0;
                        button4.Visible = true;
                    for (int i = 0; i < PageList.Count; i++)
                    {
                        Console.WriteLine("aaaaaaa" + PageList[i].APLink + deviceIP);
                        if (PageList[i].APLink + ":8899" == deviceIP && PageList[i].UpdateState == null)
                        {
                            str_data = "AP 更新 "+ PageList[i].usingAddress+ " 失敗";
                            if (reset)
                            {
                                PageList[i].UpdateState = "更新失敗";
                                PageList[i].UpdateTime = DateTime.Now.ToString();

                                foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                {
                                    Console.WriteLine("22222222222" + dr4.Cells[1].Value.ToString());
                                    if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                    {
                                        dataGridView4.Rows[dr4.Index].Cells[0].Selected = false;
                                        dr4.Cells[2].Style.BackColor = Color.Red;
                                        dr4.Cells[2].Value = DateTime.Now.ToString();
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                if (PageList[i].actionName == "down" || PageList[i].actionName == "sale" || PageList[i].actionName == "saletime")
                                {
                                    macaddress = PageList[i].usingAddress;
                                    Console.WriteLine("saletime macaddress"+ macaddress);
                                    if(down)
                                    {
                                        Console.WriteLine("down FAIL");
                                        foreach (DataGridViewRow dr in dataGridView1.Rows)
                                        {
                                            if (dr.Cells[1].Value!=null&&PageList[i].usingAddress== dr.Cells[1].Value.ToString()) {

                                                dr.Cells[0].Value = true;
                                                dr.Cells[0].ReadOnly = false;
                                            }
                                           
                                        }
                                    }
                                }
                                else {
                                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                    {
                                        if (dr.Cells[1].Value != null&& dr.Cells[1].Value.ToString()!="")
                                        {
                                        if (dr.Cells[1].Value.ToString().Contains(',' + PageList[i].usingAddress))
                                        {
                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + PageList[i].usingAddress);
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                break;
                                        }
                                        if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress + ','))
                                        {

                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress + ',');
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                break;
                                        }

                                        if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                        {

                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress);
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                break;
                                        }
                                        }

                                    }

                                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                    {
                                        if (dr.Cells[12].Value != null && dr.Cells[12].Value.ToString() != "")
                                        {
                                            if (dr.Cells[12].Value.ToString().Contains(PageList[i].usingAddress))
                                            {
                                                if (dr.Cells[1].Value.ToString().Length > 0)
                                                    dr.Cells[1].Value = dr.Cells[1].Value.ToString()+PageList[i].usingAddress;
                                                else
                                                    dr.Cells[1].Value = PageList[i].usingAddress;
                                            }
                                        }

                                    }

                                }
                                    foreach (DataGridViewRow dr in dataGridView1.Rows)
                                {

                                    if (dr.Cells[1].Value != null)
                                    {
                                        if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                        {
                                            ESLFailData.Clear();
                                            dr.Cells[4].Style.BackColor = Color.Red;
                                            dr.Cells[4].Value = DateTime.Now.ToString();
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 4, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            PageList[i].UpdateState = "更新失敗";
                                            PageList[i].UpdateTime = DateTime.Now.ToString();
                                            dataGridView1.Rows[dr.Index].Cells[0].Selected = false;
                                            int failcount = ESLUpdaateFail.Count;
                                            //Console.WriteLine("dr.Cells[1].Value.ToString()" + dr.Cells[1].Value.ToString() + failcount);
                                            ESLFailData.Add(PageList[listcount].BleAddress);
                                            ESLFailData.Add(DateTime.Now.ToString());
                                            ESLFailData.Add("連線失敗");
                                            ESLUpdaateFail.Add(ESLFailData);
                                            // Page1 mPage1 = PageList[listcount];
                                            foreach (DataGridViewRow dr3 in dataGridView3.Rows)
                                            {
                                                Console.WriteLine("1111111111" + dr3.Cells[0].Value.ToString());
                                                if (dr3.Cells[0].Value != null && dr3.Cells[0].Value.ToString().Contains(macaddress))
                                                {

                                                    dr3.Cells[4].Value = "連線失敗";
                                                    dr3.Cells[6].Value = DateTime.Now.ToString();

                                                    break;
                                                }
                                            }

                                            foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                            {
                                                Console.WriteLine("22222222222" + dr4.Cells[1].Value.ToString());
                                                if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString().Contains(macaddress))
                                                {

                                                    dr4.Cells[2].Style.BackColor = Color.Red;
                                                    dr4.Cells[2].Value = DateTime.Now.ToString();
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    break;
                                                }
                                            }


                                        }
                                    }
                                }
                            }
                            deviceIPData = PageList[i].APLink;
                            // PageList.RemoveAt(i);
                            Console.WriteLine("aaaaaaa");
                            break;
                        }    
                    }

                        if (listcount + 1 < PageList.Count)
                        {
                          //  UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                          //  Console.WriteLine("3333333333");
                            listcount++;
                        //   Console.WriteLine("error-----------"+ listcount);
                        //   Page1 mPage1 = PageList[listcount];
                        Console.WriteLine("device" + deviceIP);
                        button4.Visible = true;
                        /* Console.WriteLine("mPage1-----------" + mPage1.product_name);
                         Bitmap bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                   mPage1.specification, mPage1.price, mPage1.Special_offer,
                                    mPage1.barcode, mPage1.Web, mPage1.HeadertextALL, ESLFormat);

                         pictureBoxPage1.Image = bmp;
                         mSmcEsl.TransformImageToData(bmp);
                         */
                        /*  int numVal = Convert.ToInt32(mPage1.no) - 1;
                          if (reset)
                          {
                              dataGridView4.ClearSelection();
                              dataGridView4.Rows[numVal].Cells[0].Selected = true;
                          }
                          else {
                              dataGridView1.ClearSelection();
                              dataGridView1.Rows[numVal].Cells[0].Selected = true;
                          }
                          */
                        // aaa(datagridview1curr, true, numVal);

                        bufferConnect(deviceIP);
                          //  ConnectTimer.Interval = 3000;
                          //  ConnectTimer.Start();

                    }
                        else
                        {
                        if (sale)
                        {
                            bool TT = false;
                            foreach (DataGridViewRow dr in dataGridView1.Rows)
                            {
                                if (dr.Cells[5].Style.ForeColor != Color.Black || dr.Cells[6].Style.ForeColor != Color.Black || dr.Cells[7].Style.ForeColor != Color.Black || dr.Cells[8].Style.ForeColor != Color.Black || dr.Cells[9].Style.ForeColor != Color.Black || dr.Cells[10].Style.ForeColor != Color.Black || dr.Cells[11].Style.ForeColor != Color.Black)
                                {
                                    TT = true;
                                }
                            }
                            if (!TT)
                            {
                                button2.Enabled = false;
                                button2.BackColor = Color.Gray;

                            }


                        }

                        if (EslStyleChangeUpdate)
                        {
                            button19.BackColor = Color.FromArgb(255, 255, 192);
                            button19.Enabled = false;
                        }

                       /* if (!down && !sale && !reset && !saletime)
                        {
                            foreach (DataGridViewRow dr in dataGridView1.Rows)
                            {
                                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                                {
                                    dr.Cells[0].Value = true;
                                    dr.Cells[0].ReadOnly = false;

                                }
                                else
                                {
                                    dr.Cells[0].Value = false;
                                    dr.Cells[0].ReadOnly = true;
                                }
                                dr.Cells[1].Value = dr.Cells[12].Value;
                            }

                        }*/

                        testest = false;
                        onlockedbutton(testest);
                        checkClick = false;
                            down = false;
                            OldRunAPList.Clear();
                            backESLList.Clear();
                            listcount = 0;
                            progressBar1.Visible = false;
                            down = false;
                            sale = false;
                            reset = false;
                            saletime = false;
                            immediateUpdate = false;
                            EslStyleChangeUpdate = false;
                        pictureBoxPage1.Image = null;
                        richTextBox1.Text = "";

                        stopwatch.Stop();//碼錶停止
                            TimeSpan ts = stopwatch.Elapsed;
                            string elapsedTime = String.Format("{0:00} 分 {1:00} 秒 {2:000} ms",
                           ts.Minutes, ts.Seconds,
                           ts.Milliseconds);
                            dataGridView1.Enabled = true;
                            removeESLingstate = false;
                            onsaleESLingstate = false;
                            updateESLingstate = false;
                            checkClick = false;
                            ESLStyleDataChange = false;
                            ESLSaleStyleDataChange = false;
                            
                            datagridview1curr = 1;
                            //mSmcEsl.DisConnectBleDevice();
                            ConnectTimer.Stop();
                            int ESLBindCount = 0;
                            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                            {

                                if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() != "未綁定")
                                {
                                    ESLBindCount = ESLBindCount + 1;
                                    BindESL.Text = ESLBindCount.ToString();

                                }

                            }
                        CheckBeaconTimer.Start();
                         //   pictureBox4.Visible = false;
                            testest = false;
                        
                        
                            MessageBox.Show("全部更新完成  \r\n" + elapsedTime);
                      //   mExcelData.UpdateDataList(false, "esldemoV2.xlsx", PageList);
                     //   mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                      //  mExcelData.DataGridviewSave(dataGridView1, false, openExcelAddress);
                        PageList.Clear();
                        datagridview1curr = 2;
                            aaa(1, false, 0);
                        }
                //    }
             //       richTextBox1.Text = "連線失敗" + "\r\n" + richTextBox1.Text;
                }
            }
            // ---------  AP  設定時間 -------
            else if (msgId == EslUdpTest.SmcEsl.msg_SetRTCTime)
            {
                if (status)
                {
                    str_data = "AP 設定時間 完成";
                }
                else
                {
                    str_data = "AP 設定時間 失敗";
                }
            }
            // ---------  AP  取得時間 -------
            else if (msgId == EslUdpTest.SmcEsl.msg_GetRTCTime)
            {

                string yy = data.Substring(0, 2);
                string MM = data.Substring(2, 2);
                string dd = data.Substring(4, 2);
                string ww = data.Substring(6, 2);
                string HH = data.Substring(8, 2);
                string mm = data.Substring(10, 2);
                string ss = data.Substring(12, 2);
                str_data = yy + "/" + MM + "/" + dd + "  星期:" + ww + "  " + HH + ":" + mm + ":" + ss;

            }
            // ---------  AP  Beacon Time  -------
            else if (msgId == EslUdpTest.SmcEsl.msg_SetBeaconTime)
            {
                if (status)
                {
                    str_data = "Beacon 設定時間 完成";
                }
                else
                {
                    str_data = "Beacon 設定時間 失敗";
                }
            }



            /* if (str_data.Equals("資料寫入成功") || msgId == EslUdpTest.SmcEsl.msg_ScanDevice)
             {

             }
             else
             {
                 richTextBox1.Text= richTextBox1.Text+ "\r\n" + (Environment.NewLine + deviceIP + "  Time = " + DateTime.Now.ToString("HH:mm:ss") + " =>" + str_data); 
                 richTextBox1.SelectionStart = richTextBox1.Text.Length;
                 richTextBox1.ScrollToCaret();
             }*/
            if (str_data != "") {
                richTextBox1.Text = richTextBox1.Text + "\r\n" + (Environment.NewLine + deviceIP + "  Time = " + DateTime.Now.ToString("HH:mm:ss") + " =>" + str_data);
                richTextBox1.SelectionStart = richTextBox1.Text.Length;
                richTextBox1.ScrollToCaret();
            }



        }

     /*   private void UpdateProgressBar(string data, string deviceIP)
        {
            Console.WriteLine("UpdateProgressBar");
            progressBar1.Value += progressBar1.Step;//讓進度條增加一次
        }*/
        
         private void CheckESLLoad(object sender, EventArgs e)
        {

            //mSmcEsl.DisConnectBleDevice();
            foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
            {
                 
            }

        }


        private void DisConnectBle(object sender, EventArgs e)
        {

                //mSmcEsl.DisConnectBleDevice();
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {

                    if (kvp.Key.Contains(checkESLV[listcount].APID))
                    {
                        kvp.Value.mSmcEsl.DisConnectBleDevice();
                    }
                }
            richTextBox1.Text = richTextBox1.Text+checkESLV[listcount].ESLID + "連線逾時" + "\r\n" ;
            listcount++;
            if (checkV)
            {
                Console.WriteLine("ISRUN");
                CheckVTimer.Interval = 1000;
                CheckVTimer.Start();
            }
        }

        private void ConnectBle(object sender, EventArgs e)
        {
            Console.WriteLine("ConnectBle");

            //mSmcEsl.DisConnectBleDevice();
            /* if(Runtime == false)
             {
                 stopwatch.Reset();
                 stopwatch.Start();
             }*/


            ConnectTimer.Stop();
            //ConnectBleTimeOut.Start();
            Page1 mPage1 = new Page1();
            for (int i = 0; i < PageList.Count; i++)
            {
                if (PageList[i].APLink == deviceIPData&& PageList[i].UpdateState ==null)
                {
                    Console.WriteLine("aadddss"+PageList[i].APLink);
                    mPage1 = PageList[i];
                    break;
                }
            }
            
            Bitmap bmp;
            if (mPage1.APLink != null) {
            foreach (DataGridViewRow dr in this.dataGridView5.Rows)
            {
                Console.WriteLine(" PageList[listcount]" + PageList[listcount].product_name);
                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == PageList[listcount].APLink)
                {
                    if (dr.Cells[4].Value.ToString() == "")
                    {
                        if (down || sale)
                        {
                            MessageBox.Show(PageList[listcount].usingAddress + "該ESL綁定AP未啟用");
                        }
                        else
                        {
                            MessageBox.Show(PageList[listcount].BleAddress + "該ESL綁定AP未啟用");
                        }
                        listcount++;
                        if (listcount + 1 > PageList.Count)
                        {
                            down = false;
                            sale = false;
                            reset = false;
                            saletime = false;
                            stopwatch.Stop();//碼錶停止
                            TimeSpan ts = stopwatch.Elapsed;
                            string elapsedTime = String.Format("{0:00} 分 {1:00} 秒 {2:000} ms",
                           ts.Minutes, ts.Seconds,
                           ts.Milliseconds);
                            dataGridView1.Enabled = true;
                            removeESLingstate = false;
                            onsaleESLingstate = false;
                            updateESLingstate = false;
                            datagridview1curr = 1;
                            //mSmcEsl.DisConnectBleDevice();
                            ConnectTimer.Stop();
                            int ESLBindCount = 0;
                            foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                            {

                                if (dr4.Cells[6].Value != null && dr4.Cells[6].Value.ToString() != "未綁定")
                                {
                                    ESLBindCount = ESLBindCount + 1;
                                    BindESL.Text = ESLBindCount.ToString();

                                }

                            }
                           // mExcelData.UpdateDataList(false, "esldemoV2.xlsx", PageList);
                            testest = false;
                           // mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                           // mExcelData.DataGridviewSave(dataGridView1, false, openExcelAddress);
                            MessageBox.Show("全部更新完成  \r\n" + elapsedTime);
                            
                            datagridview1curr = 2;
                            aaa(1, false, 0);
                            return;
                        }
                    }
                }

            }




            if (down || sale || reset)
            {

                //dr.Cells[17].Value = DateTime.Now.ToString();
                macaddress = mPage1.usingAddress;
                Console.WriteLine("WWWWWWTTTFFFFFFFFF");
                    if (reset)
                    {
                        foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                        {

                            if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                            {
                             //   Console.WriteLine("AAAAAAAAAFFFFFFFFFFFFSSSSSSSSSSSSSS");
                                dataGridView4.Rows[dr.Index].Cells[0].Selected = true;
                            }
                        }
                    }
                    else {
                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                        {
                            if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                            {
                                dataGridView1.Rows[dr.Index].Cells[0].Selected = true;
                            }
                        }

                    }
                    

                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    Console.WriteLine("mPage1.APLink" + mPage1.APLink);
                    if (kvp.Key.Contains(mPage1.APLink))
                    {
                        Console.WriteLine("ININ");
                        if (down)
                        {
                            bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        }
                        if (reset)
                        {
                            bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        }

                        if (sale)
                        {
                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                         mPage1.specification, mPage1.price, mPage1.Special_offer,
                                            mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            //    bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                        }

                        //kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress,0);
                            //.Threading.Thread.Sleep(1000);
                            EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                        mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3,0);
                    }
                }
                //  mSmcEsl.ConnectBleDevice(mPage1.usingAddress);
                foreach (DataGridViewRow dr in this.dataGridView3.Rows)
                {
                    if (macaddress == dr.Cells[0].Value.ToString())
                    {
                        dr.Cells[4].Value = "連線中";
                    }
                }

                // int CurrentRow = dataGridView1.CurrentRow.Index;
                // dataGridView1.Rows[CurrentRow].Cells[17].Value = DateTime.Now.ToString();
                richTextBox1.Text = "正連接:" + mPage1.usingAddress + "\r\n" + richTextBox1.Text;
                // mSmcEsl.WriteESLData(PageList[listcount].usingAddress);
            }
            else if (saletime)
            {
                macaddress = mPage1.usingAddress;
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {
                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                        {
                            dataGridView1.Rows[dr.Index].Cells[0].Selected = true;
                        }
                    }
                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {

                    if (kvp.Key.Contains(mPage1.APLink))
                    {

                        string format = "yyyy/MM/dd HH:mm:ss";
                        string start = Convert.ToDateTime(mPage1.onSaleTimeS).ToString("yyyy/MM/dd HH:mm:ss");
                        string end = Convert.ToDateTime(mPage1.onSaleTimeE).ToString("yyyy/MM/dd HH:mm:ss");
                        DateTime strDate = DateTime.ParseExact(start, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                        DateTime endDate = DateTime.ParseExact(end, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                        if (DateTime.Compare(strDate, DateTime.Now) < 0 && DateTime.Compare(endDate, DateTime.Now) > 0)
                        {
                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                        }
                        else
                        {
                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                        }
                        kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.BleAddress,0);
                            EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                        mSmcEsl.UpdataESLDataFromBuffer(mPage1.BleAddress, 0, 8,0);
                    }
                }

            }

            else
            {

                macaddress = mPage1.BleAddress;
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {
                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.BleAddress)
                        {
                            dataGridView1.Rows[dr.Index].Cells[0].Selected = true;
                        }
                    }
                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {

                    if (kvp.Key.Contains(mPage1.APLink))
                    {

                        bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                        kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.BleAddress,0);
                            EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                        mSmcEsl.UpdataESLDataFromBuffer(mPage1.BleAddress, 0, 8,0);
                    }
                }
                //mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                //int CurrentRow = dataGridView1.CurrentRow.Index;
                //  dataGridView1.Rows[CurrentRow].Cells[17].Value = DateTime.Now.ToString();
               // richTextBox1.Text = richTextBox1.Text+"正連接:" + mPage1.BleAddress + "\r\n";
                //mSmcEsl.WriteESLData(PageList[listcount].BleAddress);
            }
        }

        }


        private void bleConnect(string deviceIP)
        {
            Console.WriteLine("bleConnect");

            //mSmcEsl.DisConnectBleDevice();
            /* if(Runtime == false)
             {
                 stopwatch.Reset();
                 stopwatch.Start();
             }*/


            ConnectTimer.Stop();
            //ConnectBleTimeOut.Start();
            Page1 mPage1 = new Page1();
            for (int i = 0; i < PageList.Count; i++)
            {
                Console.WriteLine("bleConnect" + PageList[i].BleAddress);
                Console.WriteLine("bleConnect:" + PageList[i].APLink+" "+deviceIP+" "+ PageList[i].UpdateState);
                if ((PageList[i].APLink+":8899") == deviceIP && PageList[i].UpdateState == null)
                {
                    
                    mPage1 = PageList[i];
                    break;
                }
                if (i == PageList.Count - 1)
                {
                    List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).Where(p => p.UpdateState == null).ToList();
                    foreach (Page1 p in list)
                    {
                        OldRunAPList.Add(p.APLink);

                    }
                }

            }

            Bitmap bmp;
            if (mPage1.APLink != null)
            {
                foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                {
                   
                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == PageList[listcount].APLink)
                    {
                        if (dr.Cells[4].Value.ToString() == "")                                                                                                                                                                                                                                                                                 
                        {
                            //  listcount++;
                            if (listcount + 1 > PageList.Count)
                            {
                                down = false;
                                sale = false;
                                reset = false;
                                saletime = false;
                                stopwatch.Stop();//碼錶停止
                                TimeSpan ts = stopwatch.Elapsed;
                                string elapsedTime = String.Format("{0:00} 分 {1:00} 秒 {2:000} ms",
                               ts.Minutes, ts.Seconds,
                               ts.Milliseconds);
                                dataGridView1.Enabled = true;
                                removeESLingstate = false;
                                onsaleESLingstate = false;
                                updateESLingstate = false;
                                datagridview1curr = 1;
                                //mSmcEsl.DisConnectBleDevice();
                                ConnectTimer.Stop();
                                int ESLBindCount = 0;
                                foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                                {

                                    if (dr4.Cells[6].Value != null && dr4.Cells[6].Value.ToString() != "未綁定")
                                    {
                                        ESLBindCount = ESLBindCount + 1;
                                        BindESL.Text = ESLBindCount.ToString();

                                    }

                                }
                                mExcelData.UpdateDataList(false, "esldemoV2.xlsx", PageList);
                                testest = false;
                                mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                                // mExcelData.DataGridviewSave(dataGridView1, false, openExcelAddress);
                                MessageBox.Show("全部更新完成  \r\n" + elapsedTime);
                                datagridview1curr = 2;
                                aaa(1, false, 0);
                                return;
                            }
                        }
                    }

                }
                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {

                        if (kvp.Key.Contains(mPage1.APLink))
                        {
                        //  ConnectBleTimeOut.Start();
                        Console.WriteLine("BleConnectTimer" + mPage1.BleAddress);
                        
                    /*    mPage1.TimerConnect = new System.Windows.Forms.Timer();
                        mPage1.TimerConnect.Interval = (30 * 1000);
                        mPage1.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                        mPage1.TimerSeconds = new Stopwatch();*/
                        mPage1.TimerSeconds.Start();
                        mPage1.TimerConnect.Start();

                        kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                        richTextBox1.Text = mPage1.BleAddress + "  嘗試連線中請稍候... \r\n";
                        }
                    }
            }
        }


        private void bufferConnect(string deviceIP) {
            Console.WriteLine("bufferConnect");

            //mSmcEsl.DisConnectBleDevice();
            /* if(Runtime == false)
             {
                 stopwatch.Reset();
                 stopwatch.Start();
             }*/


            ConnectTimer.Stop();
            //ConnectBleTimeOut.Start();
            Page1 mPage1 = new Page1();
            for (int i = 0; i < PageList.Count; i++)
            {
                if (PageList[i].APLink == deviceIPData && PageList[i].UpdateState == null)
                {
                    Console.WriteLine("aadddss" + PageList[i].APLink);
                    mPage1 = PageList[i];
                    break;
                }
                if(i== PageList.Count - 1) {
                    List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).Where(p => p.UpdateState ==null).ToList();
                    foreach (Page1 p in list)
                    {
                       OldRunAPList.Add(p.APLink);

                    }
                }
                  
            }

            Bitmap bmp;
            if (mPage1.APLink != null)
            {
                foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                {
                    Console.WriteLine("mPage1" + mPage1.BleAddress);
                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == PageList[listcount].APLink)
                    {
                        if (dr.Cells[4].Value.ToString() == "")
                        {
                            listcount++;
                            if (listcount + 1 > PageList.Count)
                            {
                                down = false;
                                sale = false;
                                reset = false;
                                saletime = false;
                                stopwatch.Stop();//碼錶停止
                                TimeSpan ts = stopwatch.Elapsed;
                                string elapsedTime = String.Format("{0:00} 分 {1:00} 秒 {2:000} ms",
                               ts.Minutes, ts.Seconds,
                               ts.Milliseconds);
                                dataGridView1.Enabled = true;
                                removeESLingstate = false;
                                onsaleESLingstate = false;
                                updateESLingstate = false;
                                datagridview1curr = 1;
                                //mSmcEsl.DisConnectBleDevice();
                                ConnectTimer.Stop();
                                int ESLBindCount = 0;
                                foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                                {

                                    if (dr4.Cells[6].Value != null && dr4.Cells[6].Value.ToString() != "未綁定")
                                    {
                                        ESLBindCount = ESLBindCount + 1;
                                        BindESL.Text = ESLBindCount.ToString();

                                    }

                                }
                                mExcelData.UpdateDataList(false, "esldemoV2.xlsx", PageList);
                                testest = false;
                                mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                               // mExcelData.DataGridviewSave(dataGridView1, false, openExcelAddress);
                                MessageBox.Show("全部更新完成  \r\n" + elapsedTime);
                                datagridview1curr = 2;
                                aaa(1, false, 0);
                                return;
                            }
                        }
                    }

                }




                if (mPage1.actionName=="down" || mPage1.actionName == "sale" || mPage1.actionName == "reset" || mPage1.actionName == "EslStyleChangeUpdate")
                {

                    //dr.Cells[17].Value = DateTime.Now.ToString();
                    macaddress = mPage1.usingAddress;
                    Console.WriteLine("WWWWWWTTTFFFFFFFFFBBBCCC");
                    if (mPage1.actionName == "reset")
                    {
                        foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                        {
                            if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                            {
                                dataGridView4.Rows[dr.Index].Cells[0].Selected = true;
                            }
                        }
                    }
                    else
                    {
                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                        {
                            if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                            {
                                dataGridView1.Rows[dr.Index].Cells[0].Selected = true;
                            }
                        }
                    }
                   

                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {
                        Console.WriteLine("mPage1.APLink" + mPage1.APLink);
                        if (kvp.Key.Contains(mPage1.APLink))
                        {
                            Console.WriteLine("ININ");
                            if (mPage1.actionName=="down")
                            {
                                Console.WriteLine("d");
                                bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                                kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                pictureBoxPage1.Image = bmp;
                            }
                       /*     if (immediateUpdate)
                            {
                                Console.WriteLine("d");
                                if (mPage1.onsale == "V")
                                {
                                    bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                             mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                }
                                else
                                {
                                    bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                             mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                }
                                
                                kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                pictureBoxPage1.Image = bmp;
                            }*/
                            
                            if (mPage1.actionName=="reset")
                            {
                                Console.WriteLine("r");
                                bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                                kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                pictureBoxPage1.Image = bmp;
                            }

                            if (mPage1.actionName == "sale")
                            {
                                Console.WriteLine("s");
                                if (mPage1.onsale == "V")
                                {
                                    bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                             mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                }
                                else
                                {
                                    bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                             mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                }
                                    
                                kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                pictureBoxPage1.Image = bmp;
                                //    bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                            }
                            if (mPage1.actionName == "EslStyleChangeUpdate")
                            {
                                Console.WriteLine("s");
                                if (mPage1.onsale == "V")
                                {
                                    bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                            mPage1.specification, mPage1.price, mPage1.Special_offer,
                                               mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                    kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                }
                                else
                                {
                                    bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                           mPage1.specification, mPage1.price, mPage1.Special_offer,
                                              mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                    kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                }
                              
                                pictureBoxPage1.Image = bmp;
                                //    bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                            }


                            dataGridView1.Rows[Convert.ToInt32(mPage1.no)-1].Cells[17].Value = DateTime.Now.ToString();
                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, Convert.ToInt32(mPage1.no) - 1, false, openExcelAddress, excel, excelwb, mySheet);
                            //kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress,0);

                            //.Threading.Thread.Sleep(1000);
                            EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                          //  mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3);
                         //   System.Threading.Thread.Sleep(200);
                         //   Console.WriteLine("WWWWWWWWWWWWWWWWWWWWWWWTTTTTTTTT1");
                        }
                    }
                    //  mSmcEsl.ConnectBleDevice(mPage1.usingAddress);
                    foreach (DataGridViewRow dr in this.dataGridView3.Rows)
                    {
                        if (macaddress == dr.Cells[0].Value.ToString())
                        {
                            dr.Cells[4].Value = "連線中";
                        }
                    }

                    // int CurrentRow = dataGridView1.CurrentRow.Index;
                    // dataGridView1.Rows[CurrentRow].Cells[17].Value = DateTime.Now.ToString();
             //       richTextBox1.Text = "正連接:" + mPage1.usingAddress + "\r\n" + richTextBox1.Text;
                    // mSmcEsl.WriteESLData(PageList[listcount].usingAddress);
                }
                else if (mPage1.actionName == "saletime")
                {
                    macaddress = mPage1.usingAddress;
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {
                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                        {
                            Console.WriteLine("乾 最好進不來");
                            dataGridView1.Rows[dr.Index].Cells[0].Selected = true;
                        
                        }
                    }
                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {

                        if (kvp.Key.Contains(mPage1.APLink))
                        {

                            string format = "yyyy/MM/dd HH:mm:ss";
                            string start = Convert.ToDateTime(mPage1.onSaleTimeS).ToString("yyyy/MM/dd HH:mm:ss");
                            string end = Convert.ToDateTime(mPage1.onSaleTimeE).ToString("yyyy/MM/dd HH:mm:ss");
                            DateTime strDate = DateTime.ParseExact(start, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                            DateTime endDate = DateTime.ParseExact(end, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                            if (DateTime.Compare(strDate, DateTime.Now) < 0 && DateTime.Compare(endDate, DateTime.Now) > 0)
                            { 
                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                         mPage1.specification, mPage1.price, mPage1.Special_offer,
                                            mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                pictureBoxPage1.Image = bmp;

                                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                {
                                    if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() == mPage1.product_name)
                                    {
                                        dr.Cells[15].Value = "V";
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    }
                                }
                            }
                            else
                            {
                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                         mPage1.specification, mPage1.price, mPage1.Special_offer,
                                            mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                pictureBoxPage1.Image = bmp;
                                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                {
                                    if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() == mPage1.product_name)
                                    {
                                        dr.Cells[15].Value = "X";
                                        dr.Cells[19].Value = DBNull.Value;
                                        dr.Cells[20].Value = DBNull.Value;
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    }
                                }
                             
                            }
                            dataGridView1.Rows[Convert.ToInt32(mPage1.no)-1].Cells[17].Value = DateTime.Now.ToString();
                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, Convert.ToInt32(mPage1.no) - 1, false, openExcelAddress, excel, excelwb, mySheet);
                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress,0);
                            EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                          //  mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3);
                          //  Console.WriteLine("WWWWWWWWWWWWWWWWWWWWWWWTTTTTTTTT2");
                           // System.Threading.Thread.Sleep(200);

                        }
                    }

                }

                else
                {

                    macaddress = mPage1.BleAddress;
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {
                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.BleAddress)
                        {
                            dataGridView1.Rows[dr.Index].Cells[0].Selected = true;
                        }
                    }
                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {

                        if (kvp.Key.Contains(mPage1.APLink))
                        {
                            /*     if (mPage1.onsale == "V")
                                 {
                                     bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                             mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                 }
                                 else
                                 {
                                     bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                              mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                 mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                 }

                                 dataGridView1.Rows[Convert.ToInt32(mPage1.no)-1].Cells[17].Value = DateTime.Now.ToString();
                                 mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, Convert.ToInt32(mPage1.no) - 1, false, openExcelAddress, excel, excelwb, mySheet);
                                 kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                 kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.BleAddress,0);*/
                          //  ConnectBleTimeOut.Start();
                            kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                            //   SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                            // mSmcEsl.UpdataESLDataFromBuffer(mPage1.BleAddress, 0, 3);
                            //  System.Threading.Thread.Sleep(200);
                        }
                    }
                    //mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                    //int CurrentRow = dataGridView1.CurrentRow.Index;
                    //  dataGridView1.Rows[CurrentRow].Cells[17].Value = DateTime.Now.ToString();
               //     richTextBox1.Text = "正連接:" + mPage1.BleAddress + "\r\n" + richTextBox1.Text;
                    //mSmcEsl.WriteESLData(PageList[listcount].BleAddress);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="data"></param>
        /// <param name="deviceIP"></param>
        private void UpdateUI_Scan(string data, string deviceIP, double battery)
        {
            bool check = false;
            bool ESLnotnull = false;
            //Console.WriteLine("deviceIP" + deviceIP+"data" + data);
            string RssiS = data.Substring(data.Length - 2, 2);

            data = data.Substring(0, data.Length - 2);


            /*   for (int i = 0; i < dataGridView3.RowCount; i++)
                {
                   Console.WriteLine("+++++++++++++++++++");
                    if (dataGridView3.Rows[i].Cells[0].Value.ToString() != null)
                    {
                        if (dataGridView3.Rows[i].Cells[0].Equals(data))
                        {
                            dataGridView3.Rows[i + 1].Cells[1].Value = RssiS;
                            break;
                        }
                    }

                }*/
            for (int b = 0; b < leftmosueESL.Count; b++) {
                //    Console.WriteLine("leftmosueESL[b]" + leftmosueESL[b] + "data" + data);
                if (leftmosueESL[b] == data) {
                    //   Console.WriteLine("lefsssssssssssssssss[b");
                    dataGridView3.Rows[b].Cells[1].Value = (int)Convert.ToByte(RssiS, 16);
                    dataGridView3.Rows[b].Cells[8].Value = battery;
                    if (dataGridView3.Rows[b].Cells[8].Value != null && Convert.ToDouble(dataGridView3.Rows[b].Cells[8].Value) < 2.85) {
                        dataGridView3.Rows[b].Cells[8].Style.ForeColor = Color.Red;
                    }
                }
            }

            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
            {
                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == data)
                {
                    if (dr.Cells[4].Value != null && dr.Cells[4].Value.ToString() == "") {
                      
                        if (dr.Cells[8].Value != null && dr.Cells[8].Value.ToString() == ESLFromIP)
                        {
                              if (checkESLRSSIClick)
                        {
                            Console.WriteLine("FFFFFFFFFFFFFFFF");
                            CheckESLStateTimer.Stop();
                            CheckESLStateTimer.Start();
                        }
                       
                            realESL.Text = (Convert.ToInt32(realESL.Text) + 1).ToString();
                            dr.Cells[4].Value = (int)Convert.ToByte(RssiS, 16);
                            dr.Cells[5].Value = battery;
                            if (dr.Cells[5].Value != null && Convert.ToDouble(dr.Cells[5].Value) < 2.85)
                            {
                                dr.Cells[5].Style.ForeColor = Color.Red;
                            }
                        }
                    }
                    ESLnotnull = true;
                }
            }
        /*    if (dataGridView6.Rows.Count > 0) { 
            foreach (DataGridViewRow dr in this.dataGridView6.Rows)
            {
                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == data && dr.Cells[3].Value.ToString() == ESLFromIP)
                {
                    dr.Cells[2].Value = (int)Convert.ToByte(RssiS, 16);
                }
            }
            }*/
            if (!ESLnotnull)
            {
                DataTable dt = dataGridView4.DataSource as DataTable;
                dt.Rows.Add(new object[] { data, null, null, (int)Convert.ToByte(RssiS, 16), battery,"未綁定" });
               
                ESLnotnull = false;
                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 1, dataGridView4.Rows.Count-2, false, openExcelAddress, excel, excelwb, mySheet);
                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dataGridView4.Rows.Count-2, false, openExcelAddress, excel, excelwb, mySheet);
                CountESLAll.Text = (Convert.ToInt32(CountESLAll.Text) + 1).ToString();
                realESL.Text = (Convert.ToInt32(realESL.Text) + 1).ToString();
            }
            
           if (autoNullESLMate)
            {
  
                if (dataGridView4.RowCount > 0)
                {
    
                    for (int i = 0; i < autoNullESLData.Count; i++)
                    {
        
                        if (autoNullESLData[i] != null && autoNullESLData[i] == data)
                        {
                            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                            {
                       
                            if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString()== autoNullESLData[i]) {
                                   int nowAPESLCount = 0;
                                    foreach (DataGridViewRow dr5 in this.dataGridView5.Rows) {
                                        if (dr5.Cells[2].Value!=null&&ESLFromIP == dr5.Cells[2].Value.ToString()) {
                                            nowAPESLCount = Convert.ToInt32(dr5.Cells[5].Value.ToString());
                                        }
                                    }
                                if (dr.Cells[8].Value!=null&&dr.Cells[8].Value.ToString() == "")
                                {
                         
                                        if (nowAPESLCount < 255) {


                                            dr.Cells[8].Value = ESLFromIP;
                                            foreach (DataGridViewRow dr3 in this.dataGridView3.Rows)
                                            {
                                                dr3.Cells[3].Value = ESLFromIP;
                                            }
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 8, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.Cells[4].Value = (int)Convert.ToByte(RssiS, 16);
                                            dr.Cells[5].Value = battery;

                                            if (dr.Cells[5].Value != null && Convert.ToDouble(dr.Cells[5].Value) < 2.85)
                                            {
                                                dr.Cells[5].Style.ForeColor = Color.Red;
                                            }
                                            //CountESLAll.Text = (Convert.ToInt32(CountESLAll.Text) + 1).ToString();
                                            foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                            {
                                                if (dr.Cells[8].Value != null && dr5.Cells[2].Value != null &&  dr5.Cells[2].Value.ToString()=="")
                                                {
                                                   // Console.WriteLine(dr5.Cells[5].Value.ToString()+"YYYYYYYYYYYYYYYY111");
                                                    dr5.Cells[5].Value = Convert.ToInt32(dr5.Cells[5].Value.ToString()) - 1;
                                                }
                                                    if (dr5.Cells[2].Value!=null&&ESLFromIP == dr5.Cells[2].Value.ToString())
                                                {
                                                    
                                                    dr5.Cells[5].Value = Convert.ToInt32(dr5.Cells[5].Value.ToString()) + 1;
                                                }
                                            }


                                        }
                                    }
                                else
                                {
                                        int oldAPESLCount=0;

                                
                                        foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                        {
                                            if (dr.Cells[8].Value!=null && dr5.Cells[2].Value!=null&&dr.Cells[8].Value.ToString() == dr5.Cells[2].Value.ToString())
                                            {
                                                oldAPESLCount = Convert.ToInt32(dr5.Cells[5].Value.ToString());
                                            }
                                        }
                                        if (nowAPESLCount < oldAPESLCount)
                                    {
                                            if (nowAPESLCount<255&& (int)Convert.ToByte(RssiS, 16)<76) {
                                          
                                                dr.Cells[8].Value = ESLFromIP;
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 8, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                                {
                                                    if (dr.Cells[8].Value != null && dr5.Cells[2].Value != null && dr.Cells[8].Value.ToString() == dr5.Cells[2].Value.ToString())
                                                    {
                                                        dr5.Cells[5].Value = Convert.ToInt32(dr5.Cells[5].Value.ToString())-1;
                                                    }
                                                    if (dr5.Cells[2].Value!=null&&ESLFromIP == dr5.Cells[2].Value.ToString())
                                                    {
                                                        dr5.Cells[5].Value = Convert.ToInt32(dr5.Cells[5].Value.ToString()) + 1;
                                                    }
                                                }
                                            }
                                        
                                    }

                                }
                                    break;
                                }
                               
                            }
                        }
                    }
                }
            }

                if (autoMateESL) {
                if (dataGridView4.RowCount > 0) {
                foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                {
                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == data)
                    {
                          //  Console.WriteLine("YYYYYYYYYYYYYYYY111");
                            int nowAPESLCount = 0;
                            foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                            {
                                if (dr5.Cells[2].Value!=null&&ESLFromIP == dr5.Cells[2].Value.ToString())
                                {
                                //    Console.WriteLine("YYYYYYYYYYYYYYYY222");
                                    if (dr5.Cells[5].Value != null && dr5.Cells[5].Value.ToString() == "")
                                    {
                                        //dr5.Cells[5].Value = 0;
                                        nowAPESLCount = Convert.ToInt32(dr5.Cells[5].Value.ToString());
                                    }
                                }
                            }
                          //  Console.WriteLine("YYYYYYY3333333333333");
                            if (dr.Cells[8].Value!=null&&dr.Cells[8].Value.ToString() == "")
                            {

                                //   Console.WriteLine("YYYYYYYYYYYYYYYY333");
                                foreach (DataGridViewRow dr3 in this.dataGridView3.Rows)
                                {
                                    dr3.Cells[3].Value = ESLFromIP;
                                }
                                dr.Cells[8].Value = ESLFromIP;
                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 8, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                dr.Cells[4].Value = (int)Convert.ToByte(RssiS, 16);
                                dr.Cells[5].Value = battery;
                                if (dr.Cells[5].Value != null && Convert.ToDouble(dr.Cells[5].Value) < 2.85)
                                {
                                    dr.Cells[5].Style.ForeColor = Color.Red;
                                }
                                //CountESLAll.Text = (Convert.ToInt32(CountESLAll.Text) + 1).ToString();
                                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                {
                                    if (dr.Cells[8].Value != null && dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString()=="")
                                    {
                                        dr5.Cells[5].Value = Convert.ToInt32(dr5.Cells[5].Value.ToString()) - 1;
                                    }
                                    if (dr5.Cells[2].Value!=null&&ESLFromIP == dr5.Cells[2].Value.ToString())
                                    {
                                      //  Console.WriteLine("YYYYYYYYYYYYYYYY4444");
                                        dr5.Cells[5].Value = Convert.ToInt32(dr5.Cells[5].Value.ToString()) + 1;
                                    }
                                }

                             
                            }
                            else {
                                int oldAPESLCount = 0;
                                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                {
                                    if (dr.Cells[8].Value != null && dr5.Cells[2].Value != null && dr.Cells[8].Value.ToString() == dr5.Cells[2].Value.ToString())
                                    {
                                     //   Console.WriteLine("YYYYYYYYYYYYYYYY5555");
                                        oldAPESLCount = Convert.ToInt32(dr5.Cells[5].Value.ToString());
                                     //   Console.WriteLine("YYYYYYYYYYYYYYYY5hhh"+ oldAPESLCount);
                                    }
                                }
                           //     Console.WriteLine((int)Convert.ToByte(RssiS, 16)+"sf"+ Convert.ToInt32(dr.Cells[4].Value));
                                if (nowAPESLCount < oldAPESLCount) {

                                    if (nowAPESLCount < 255&& (int)Convert.ToByte(RssiS, 16)<76)
                                    {
                                       // Console.WriteLine("YYYYYYYYYYYYYYYY666");
                                        dr.Cells[8].Value = ESLFromIP;
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView4, 8, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                        {
                                        //    Console.WriteLine("YYYYYYYYYYYYYYYY777");
                                            if (dr.Cells[8].Value != null && dr5.Cells[2].Value != null && dr.Cells[8].Value.ToString() == dr5.Cells[2].Value.ToString())
                                            {
                                                dr5.Cells[5].Value = Convert.ToInt32(dr5.Cells[5].Value.ToString()) - 1;
                                            }
                                            if (dr5.Cells[2].Value != null && ESLFromIP == dr5.Cells[2].Value.ToString())
                                            {
                                                dr5.Cells[5].Value = Convert.ToInt32(dr5.Cells[5].Value.ToString()) + 1;
                                            }
                                        }
                                    }
                                }
                                   
                            }
                            break;
                        }
                }
                }
              /*  if (dataGridView6.RowCount > 0) {
                    foreach (DataGridViewRow dr in this.dataGridView6.Rows)
                    {
                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == data)
                        {
                            if (dr.Cells[3].Value.ToString() == null)
                                dr.Cells[3].Value = ESLFromIP;

                            if ((int)Convert.ToByte(RssiS, 16) <= Convert.ToInt32(dr.Cells[2].Value))
                                dr.Cells[3].Value = ESLFromIP;
                        }
                    }
                }*/
            }
           
            /*     foreach (DataGridViewRow dr in this.dataGridView3.Rows)
                 {
                 Console.WriteLine(" dr.Cells[0].Value.ToString()" + dr.Cells[0].Value.ToString() + "data"+data);
                 if (data == dr.Cells[0].Value.ToString())
                     {
                     Console.WriteLine("AAAAAAAAAAAAAAAAAaaa"+ RssiS);
                         dr.Cells[1].Value = RssiS;
                     break;
                     }
                 }*/


            /*    if (dataGridView3.RowCount != 1)
                  {
                      for (int i = 0; i < firstbuildlistID.Count; i++)
                      {
                          if (firstbuildlistID[i] == data)
                          {
                              check = true;
                              break;
                          }
                      }
                      if (!check)
                      {
                          if (firstbuildcount + 1 <= dataGridView1.RowCount - 1)
                          {
                              Console.WriteLine("EEEEeeee");
                              if (!check)
                              {
                                  Console.WriteLine("QQQ");
                                  dataGridView3.Rows[firstbuildcount].Cells[1].Value = RssiS;
                                  firstbuildlistID.Add(data);
                              }
                          }
                          else
                          {
                              int count = firstbuildcount + 2 - dataGridView1.RowCount;
                              richTextBox1.Text = "多於標籤數:" + count + "\r\n" + richTextBox1.Text;
                          }
                          firstbuildcount++;
                      }
                  }*/
            //throw new NotImplementedException();
        }

        private void CheckConnectBle(object sender, EventArgs e)
        {
            Console.WriteLine("CheckConnectBle");
            /* if(Runtime == false)
             {
                 stopwatch.Reset();
                 stopwatch.Start();
             }*/
            CheckConnectTimer.Stop();

            Console.WriteLine("checkaddress.Count" + checkaddress.Count + "ssssss:" + checkconnectcount);
         //   ConnectBleTimeOut.Start();
            mSmcEsl.ConnectBleDevice(checkaddress[checkconnectcount]);
            richTextBox1.Text = "正連接:" + checkaddress[checkconnectcount] + "\r\n" + richTextBox1.Text;

        }

        private void CheckVConnectBle(object sender, EventArgs e)
        {
            CheckVTimer.Stop();
            if (listcount < checkESLV.Count)
            {
                macaddress = checkESLV[listcount].ESLID;
                Console.WriteLine(listcount + "WORK" + checkESLV.Count);
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    if (kvp.Key.Contains(checkESLV[listcount].APID))
                    {
                      //  ConnectBleTimeOut.Start();
                        kvp.Value.mSmcEsl.ConnectBleDevice(checkESLV[listcount].ESLID);
                        dataGridView4.Rows[Convert.ToInt32(checkESLV[listcount].No)].Cells[0].Selected = true;
                        richTextBox1.Text = richTextBox1.Text + "\r\n" + "正連接:" + checkESLV[listcount].ESLID;
                    }
                }
                DisConnectTimer.Start();

            }
            else
            {
                countconnect = 0;
                   checkV = false;
                Console.WriteLine("OUT");
                MessageBox.Show("電壓更新完畢");
               // mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                datagridview1curr = 2;
                aaa(1, false, 0);
            }

        }
        

       private void CheckESLState(object sender, EventArgs e)
        {

            Console.WriteLine("FFFFFFFFFFFFFFFF");
            CheckESLStateTimer.Stop();
            foreach (DataGridViewRow dr4 in dataGridView4.Rows)
            {
                if (dr4.Cells[4].Value != null && dr4.Cells[5].Value != null && dr4.Cells[8].Value != null)
                {
                    if (dr4.Cells[4].Value.ToString() == ""&& dr4.Cells[8].Value.ToString() != "")
                    {
                        dr4.Cells[4].Style.BackColor = Color.Red;
                        dr4.Cells[1].Style.BackColor = Color.Red;
                        dr4.Cells[7].Style.BackColor = Color.Red;
                    }

                    if (dr4.Cells[5].Value.ToString()!=""&&Convert.ToDouble(dr4.Cells[5].Value) < 2.85 && dr4.Cells[8].Value.ToString() != "")
                    {
                        dr4.Cells[5].Style.ForeColor = Color.Red;
                        dr4.Cells[1].Style.BackColor = Color.Red;
                        dr4.Cells[7].Style.BackColor = Color.Red;
                    }
                }
            }
            checkESLRSSIClick = false;
            MessageBox.Show("檢測完畢");

        }



        //尺寸確認計時
        private void ReadTypeTimer_TimeOut(object sender, EventArgs e)
        {
            Console.WriteLine("ReadTypeTimer_TimeOut");
            //ReadTypeTimer.Stop();
            Page1 mPage1 = new Page1();
           
            for (int i = 0; i < PageList.Count; i++)
            {
                if (PageList[i].APLink == deviceIPData && PageList[i].UpdateState == null)
                {
                    Console.WriteLine("aadddss" + PageList[i].APLink);
                    mPage1 = PageList[i];
                    break;
                }
                if (i == PageList.Count - 1)
                {
                    List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).Where(p => p.UpdateState == null).ToList();
                    foreach (Page1 p in list)
                    {
                        OldRunAPList.Add(p.APLink);

                    }
                }

            }
            Bitmap bmp;
            if (mPage1.actionName == "down" || mPage1.actionName == "sale" || mPage1.actionName == "reset" || mPage1.actionName == "EslStyleChangeUpdate")
            {

                //dr.Cells[17].Value = DateTime.Now.ToString();
                macaddress = mPage1.usingAddress;
                Console.WriteLine("WWWWWWTTTFFFFFFFFFBBBCCC");
                if (mPage1.actionName == "reset")
                {
                    foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                    {
                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                        {
                            dataGridView4.Rows[dr.Index].Cells[0].Selected = true;
                        }
                    }
                }
                else
                {
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {
                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                        {
                            dataGridView1.Rows[dr.Index].Cells[0].Selected = true;
                        }
                    }
                }


                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    Console.WriteLine("mPage1.APLink" + mPage1.APLink);
                    if (kvp.Key.Contains(mPage1.APLink))
                    {
                        Console.WriteLine("ININ");
                        if (mPage1.actionName == "down")
                        {
                            Console.WriteLine("d");

                            bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            pictureBoxPage1.Image = bmp;
                        }
                        /*     if (immediateUpdate)
                             {
                                 Console.WriteLine("d");
                                 if (mPage1.onsale == "V")
                                 {
                                     bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                              mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                 mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                 }
                                 else
                                 {
                                     bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                              mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                 mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                 }

                                 kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                 pictureBoxPage1.Image = bmp;
                             }*/

                        if (mPage1.actionName == "reset")
                        {
                            Console.WriteLine("r");
                            bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            pictureBoxPage1.Image = bmp;
                        }

                        if (mPage1.actionName == "sale")
                        {

                            Console.WriteLine("s");
                            if (mPage1.onsale == "V")
                            {
                                foreach (DataGridViewRow dr in this.dataGridView7.Rows)
                                {
                                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && dr.Cells[4].Value.ToString()== "0")
                                    {
                                        PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                    }
                                }
                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                         mPage1.specification, mPage1.price, mPage1.Special_offer,
                                            mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                            }
                            else
                            {
                                foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                                {
                                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && dr.Cells[4].Value.ToString() == "0")
                                    {
                                        PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                    }
                                }
                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                         mPage1.specification, mPage1.price, mPage1.Special_offer,
                                            mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                            }

                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            pictureBoxPage1.Image = bmp;
                            //    bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                        }
                        if (mPage1.actionName == "EslStyleChangeUpdate")
                        {
                            Console.WriteLine("s");
                            if (mPage1.onsale == "V")
                            {

                                foreach (DataGridViewRow dr in this.dataGridView7.Rows)
                                {
                                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && dr.Cells[4].Value.ToString() == "0")
                                    {
                                        PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                    }
                                }

                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                        mPage1.specification, mPage1.price, mPage1.Special_offer,
                                           mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            }
                            else
                            {
                                foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                                {
                                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && dr.Cells[4].Value.ToString() == "0")
                                    {
                                        PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                    }
                                }

                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                       mPage1.specification, mPage1.price, mPage1.Special_offer,
                                          mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            }

                            pictureBoxPage1.Image = bmp;
                            //    bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                        }


                        dataGridView1.Rows[Convert.ToInt32(mPage1.no) - 1].Cells[17].Value = DateTime.Now.ToString();
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, Convert.ToInt32(mPage1.no) - 1, false, openExcelAddress, excel, excelwb, mySheet);
                        ProgressBarVisible(PageList.Count);
                        //kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        //  kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress, 0);
                        kvp.Value.mSmcEsl.WriteESLDataWithBle();

                        //.Threading.Thread.Sleep(1000);
                        EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                        //  mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3);
                        //   System.Threading.Thread.Sleep(200);
                        //   Console.WriteLine("WWWWWWWWWWWWWWWWWWWWWWWTTTTTTTTT1");
                    }
                }
                //  mSmcEsl.ConnectBleDevice(mPage1.usingAddress);
                foreach (DataGridViewRow dr in this.dataGridView3.Rows)
                {
                    if (macaddress == dr.Cells[0].Value.ToString())
                    {
                        dr.Cells[4].Value = "連線中";
                    }
                }

                // int CurrentRow = dataGridView1.CurrentRow.Index;
                // dataGridView1.Rows[CurrentRow].Cells[17].Value = DateTime.Now.ToString();
                //       richTextBox1.Text = "正連接:" + mPage1.usingAddress + "\r\n" + richTextBox1.Text;
                // mSmcEsl.WriteESLData(PageList[listcount].usingAddress);
            }
            else if (mPage1.actionName == "saletime")
            {
                macaddress = mPage1.usingAddress;
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                    {
                        Console.WriteLine("乾 最好進不來");
                        dataGridView1.Rows[dr.Index].Cells[0].Selected = true;

                    }
                }
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {

                    if (kvp.Key.Contains(mPage1.APLink))
                    {

                        string format = "yyyy/MM/dd HH:mm:ss";
                        string start = Convert.ToDateTime(mPage1.onSaleTimeS).ToString("yyyy/MM/dd HH:mm:ss");
                        string end = Convert.ToDateTime(mPage1.onSaleTimeE).ToString("yyyy/MM/dd HH:mm:ss");
                        DateTime strDate = DateTime.ParseExact(start, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                        DateTime endDate = DateTime.ParseExact(end, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                        if (DateTime.Compare(strDate, DateTime.Now) < 0 && DateTime.Compare(endDate, DateTime.Now) > 0)
                        {
                            foreach (DataGridViewRow dr in this.dataGridView7.Rows)
                            {
                                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && dr.Cells[4].Value.ToString() == "0")
                                {
                                    PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                }
                            }

                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                            pictureBoxPage1.Image = bmp;

                            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                            {
                                if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() == mPage1.product_name)
                                {
                                    dr.Cells[15].Value = "V";
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                }
                            }
                        }
                        else
                        {

                            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                            {
                                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && dr.Cells[4].Value.ToString() == "0")
                                {
                                    PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                }
                            }
                        
                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                            pictureBoxPage1.Image = bmp;
                            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                            {
                                if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() == mPage1.product_name)
                                {
                                    dr.Cells[15].Value = "X";
                                    dr.Cells[19].Value = DBNull.Value;
                                    dr.Cells[20].Value = DBNull.Value;
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                }
                            }

                        }
                        dataGridView1.Rows[Convert.ToInt32(mPage1.no) - 1].Cells[17].Value = DateTime.Now.ToString();
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, Convert.ToInt32(mPage1.no) - 1, false, openExcelAddress, excel, excelwb, mySheet);
                        ProgressBarVisible(PageList.Count);
                        kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        // kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress, 0);
                        kvp.Value.mSmcEsl.WriteESLDataWithBle();
                        EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                        //  mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3);
                        //  Console.WriteLine("WWWWWWWWWWWWWWWWWWWWWWWTTTTTTTTT2");
                        // System.Threading.Thread.Sleep(200);

                    }
                }

            }

            else
            {

                macaddress = mPage1.BleAddress;
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.BleAddress)
                    {
                        dataGridView1.Rows[dr.Index].Cells[0].Selected = true;
                    }
                }

                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {

                    if (kvp.Key.Contains(mPage1.APLink))
                    {
                        if (mPage1.onsale == "V")
                        {
                            foreach (DataGridViewRow dr in this.dataGridView7.Rows)
                            {
                                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && dr.Cells[4].Value.ToString() == "0")
                                {
                                    PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                }
                            }

                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                    mPage1.specification, mPage1.price, mPage1.Special_offer,
                                       mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                        }
                        else
                        {
                            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                            {
                                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && dr.Cells[4].Value.ToString() == "0")
                                {
                                    PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                }
                            }

                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                        }

                        dataGridView1.Rows[Convert.ToInt32(mPage1.no) - 1].Cells[17].Value = DateTime.Now.ToString();
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, Convert.ToInt32(mPage1.no) - 1, false, openExcelAddress, excel, excelwb, mySheet);
                        ProgressBarVisible(PageList.Count);
                        kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        kvp.Value.mSmcEsl.WriteESLDataWithBle();
                        // kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.BleAddress,0);
                        //  kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.BleAddress);

                    }
                }
            }  
}

//資料傳輸超時，斷線
private void WriteESL_TimeOut(object sender, EventArgs e)
        {

            BleWriteTimer.Stop();
          // this.progressBar1.Visible = false;
            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {
                if (kvp.Key.Contains(PageList[listcount].APLink))
                {
                    kvp.Value.mSmcEsl.DisConnectBleDevice();
                }
            }
        }

        //藍牙連線超過時間
        private void ConnectBle_TimeOut(object sender, EventArgs e)
        {
            
//ConnectBleTimeOut.Stop();
         //   this.progressBar1.Visible = false; //隱藏進度條
            richTextBox1.Text = "連線逾時...";
            richTextBox1.ForeColor = Color.Red;
            Console.WriteLine("~~~~~~~~~~");
            Runtime = true;
            /*          if (countconnect < 2)
                      {
                          if (CheckESLOnly)
                          {
                              Console.WriteLine("CheckESLOnly");
                              APESLState.Text = "重新嘗試" + countconnect + "次";
                              //1/2新年第一天上工摟~~~
                              ConnectTimer.Interval = 1000;
                              ConnectTimer.Start();
                              countconnect++;

                          }
                          else
                          {
                              Console.WriteLine("ERROR" + countconnect);
                              ConnectTimer.Interval = 1000;
                              ConnectTimer.Start();
                              countconnect++;
                          }

                      }
                      else
                      {*/

            countconnect = 0;
            button4.Visible = true;
            string str_data = "";
            Page1 mPage1 = new Page1();
            Console.WriteLine("PageList[i] ConnectTimerOut"+ PageList.Count);
            for (int i = 0; i < PageList.Count; i++)
            {
                Console.WriteLine("PageList[i].TimerSeconds.Elapsed.Seconds"+ PageList[i].TimerSeconds.Elapsed.Seconds);
                Console.WriteLine("PageList[i].CounnectTimerOut" + PageList[i].usingAddress+" " + PageList[i].UpdateState+" " + PageList[i].APLink);
                if (PageList[i].TimerSeconds.Elapsed.Seconds >=29  && PageList[i].UpdateState == null)
                {
                    str_data = "AP 更新 " + PageList[i].usingAddress + " 失敗";
                    PageList[i].TimerSeconds.Stop();
                    PageList[i].TimerConnect.Stop();
                    mPage1 = PageList[i];
                    if (reset)
                    {

                        PageList[i].UpdateState = "更新失敗";

                        PageList[i].UpdateTime = DateTime.Now.ToString();

                        foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                        {
                            Console.WriteLine("22222222222" + dr4.Cells[1].Value.ToString());
                            if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                            {
                                dataGridView4.Rows[dr4.Index].Cells[0].Selected = false;
                                dr4.Cells[2].Style.BackColor = Color.Red;
                                dr4.Cells[2].Value = DateTime.Now.ToString();
                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                break;
                            }
                        }
                    }
                    else
                    {
                        if (PageList[i].actionName == "down" || PageList[i].actionName == "sale" || PageList[i].actionName == "saletime")
                        {
                            macaddress = PageList[i].usingAddress;
                            Console.WriteLine("saletime macaddress" + macaddress);
                            if (down)
                            {
                                Console.WriteLine("down FAIL"+ PageList[i].usingAddress);
                                foreach (DataGridViewRow dr in dataGridView1.Rows)
                                {
                                    if (dr.Cells[1].Value != null && PageList[i].usingAddress == dr.Cells[1].Value.ToString())
                                    {

                                        dr.Cells[0].Value = true;
                                        dr.Cells[0].ReadOnly = false;
                                    }

                                }
                            }
                            foreach (DataGridViewRow dr in dataGridView1.Rows)
                            {

                                if (dr.Cells[1].Value != null)
                                {
                                    Console.WriteLine("ConnectBleTimeOu=====PageList[i].usingAddress" + PageList[i].usingAddress);
                                    Console.WriteLine("ConnectBleTimeOu=====PageList[i].usingAddress" + dr.Cells[1].Value.ToString());
                                    Console.WriteLine("ConnectBleTimeOu=====PageList[i].usingAddress" + dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress));
                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                    {
                                        ESLFailData.Clear();
                                        dr.Cells[4].Style.BackColor = Color.Red;
                                        dr.Cells[4].Value = DateTime.Now.ToString();
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 4, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        PageList[i].UpdateState = "更新失敗";
                                        Console.WriteLine("ConnectBleTimeOu=====PageList[i].UpdateState" + PageList[i].UpdateState);
                                        PageList[i].UpdateTime = DateTime.Now.ToString();
                                        dataGridView1.Rows[dr.Index].Cells[0].Selected = false;
                                        int failcount = ESLUpdaateFail.Count;
                                        //Console.WriteLine("dr.Cells[1].Value.ToString()" + dr.Cells[1].Value.ToString() + failcount);
                                        ESLFailData.Add(PageList[listcount].BleAddress);
                                        ESLFailData.Add(DateTime.Now.ToString());
                                        ESLFailData.Add("連線失敗");
                                        ESLUpdaateFail.Add(ESLFailData);
                                        // Page1 mPage1 = PageList[listcount];
                                        foreach (DataGridViewRow dr3 in dataGridView3.Rows)
                                        {
                                            Console.WriteLine("1111111111" + dr3.Cells[0].Value.ToString());
                                            if (dr3.Cells[0].Value != null && dr3.Cells[0].Value.ToString().Contains(macaddress))
                                            {

                                                dr3.Cells[4].Value = "連線失敗";
                                                dr3.Cells[6].Value = DateTime.Now.ToString();

                                                break;
                                            }
                                        }

                                        foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                        {
                                            Console.WriteLine("22222222222" + dr4.Cells[1].Value.ToString());
                                            if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString().Contains(macaddress))
                                            {

                                                dr4.Cells[2].Style.BackColor = Color.Red;
                                                dr4.Cells[2].Value = DateTime.Now.ToString();
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                break;
                                            }
                                        }


                                    }
                                }
                            }
                        }
                        else
                        {

                            foreach (DataGridViewRow dr in dataGridView1.Rows)
                            {

                                if (dr.Cells[1].Value != null)
                                {
                                    Console.WriteLine("ConnectBleTimeOu=====PageList[i].usingAddress" + PageList[i].usingAddress);
                                    Console.WriteLine("ConnectBleTimeOu=====PageList[i].usingAddress" + dr.Cells[1].Value.ToString());
                                    Console.WriteLine("ConnectBleTimeOu=====PageList[i].usingAddress" + dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress));
                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                    {
                                        ESLFailData.Clear();
                                        dr.Cells[4].Style.BackColor = Color.Red;
                                        dr.Cells[4].Value = DateTime.Now.ToString();
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 4, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        PageList[i].UpdateState = "更新失敗";
                                        Console.WriteLine("ConnectBleTimeOu=====PageList[i].UpdateState" + PageList[i].UpdateState);
                                        PageList[i].UpdateTime = DateTime.Now.ToString();
                                        dataGridView1.Rows[dr.Index].Cells[0].Selected = false;
                                        int failcount = ESLUpdaateFail.Count;
                                        //Console.WriteLine("dr.Cells[1].Value.ToString()" + dr.Cells[1].Value.ToString() + failcount);
                                        ESLFailData.Add(PageList[listcount].BleAddress);
                                        ESLFailData.Add(DateTime.Now.ToString());
                                        ESLFailData.Add("連線失敗");
                                        ESLUpdaateFail.Add(ESLFailData);
                                        // Page1 mPage1 = PageList[listcount];
                                        foreach (DataGridViewRow dr3 in dataGridView3.Rows)
                                        {
                                            Console.WriteLine("1111111111" + dr3.Cells[0].Value.ToString());
                                            if (dr3.Cells[0].Value != null && dr3.Cells[0].Value.ToString().Contains(macaddress))
                                            {

                                                dr3.Cells[4].Value = "連線失敗";
                                                dr3.Cells[6].Value = DateTime.Now.ToString();

                                                break;
                                            }
                                        }

                                        foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                        {
                                            Console.WriteLine("22222222222" + dr4.Cells[1].Value.ToString());
                                            if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString().Contains(macaddress))
                                            {

                                                dr4.Cells[2].Style.BackColor = Color.Red;
                                                dr4.Cells[2].Value = DateTime.Now.ToString();
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                break;
                                            }
                                        }


                                    }
                                }
                            }
                            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                            {
                                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                                {
                                    if (dr.Cells[1].Value.ToString().Contains(',' + PageList[i].usingAddress))
                                    {
                                        int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + PageList[i].usingAddress);
                                        dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                    }
                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress + ','))
                                    {

                                        int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress + ',');
                                        dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                    }

                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                    {

                                        int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress);
                                        dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                    }
                                }

                            }

                            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                            {
                                if (dr.Cells[12].Value != null && dr.Cells[12].Value.ToString() != "")
                                {
                                    if (dr.Cells[12].Value.ToString().Contains(PageList[i].usingAddress))
                                    {
                                        if (dr.Cells[1].Value.ToString().Length > 0)
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString() + PageList[i].usingAddress;
                                        else
                                            dr.Cells[1].Value = PageList[i].usingAddress;
                                    }
                                }

                            }

                        }

                    }
                    deviceIPData = PageList[i].APLink;
                    // PageList.RemoveAt(i);
                    Console.WriteLine("aaaaaaa");
                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {
                        if (kvp.Key.Contains(mPage1.APLink))
                        {
                            kvp.Value.mSmcEsl.DisConnectBleDevice();
                        }
                    }


                    break;
                }
            }



          


        }
        //---------------


        //==================================================================================
        #region Even
        //    掃秒
        /*     private void ScanDeviceEven(object sender, EventArgs e)
             {
                 string receivedText = (e as SmcEsl.ScanDeviceEventArgs).address;
                 string deviceIP = (e as SmcEsl.ScanDeviceEventArgs).deviceIP;
                 ESLFromIP = deviceIP;
                 UIInvoker stc = new UIInvoker(UpdateUI_Scan);
                 this.BeginInvoke(stc, receivedText, deviceIP);
                 Console.Write("ssssssssssssss" + deviceIP + Environment.NewLine);
             }

             //連線
             private void ConnectEslDeviceEven(object sender, EventArgs e)
             {
                 Console.WriteLine("ConnectEslDeviceEven");
                 bool isConnect = (e as SmcEsl.ConnectDeviceEventArgs).isConnect;
                 string deviceIP = (e as SmcEsl.ConnectDeviceEventArgs).deviceIP;
                 UIInvoker stc = new UIInvoker(UpdateUI);
                 if (isConnect)
                 {
                     this.BeginInvoke(stc, "連線成功", deviceIP);
                 }
                 else
                 {
                     this.BeginInvoke(stc, "連線失敗", deviceIP);
                 }
             }
             //斷線
             private void DisconnectEslDeviceEven(object sender, EventArgs e)
             {
                 Console.WriteLine("DisconnectEslDeviceEven");
                 bool isDisconnect = (e as SmcEsl.DisconnectDeviceEventArgs).isDisconnect;
                 string deviceIP = (e as SmcEsl.DisconnectDeviceEventArgs).deviceIP;

                 UIInvoker stc = new UIInvoker(UpdateUI);
                 if (isDisconnect)
                 {
                     this.BeginInvoke(stc, "斷線成功", deviceIP);

                 }
                 else
                 {
                     this.BeginInvoke(stc, "斷線失敗", deviceIP);
                 }
             }

             // 連線超過時間
             private void ConnectBleTimeOut(object sender, EventArgs e)
             {
                 UIInvoker stc = new UIInvoker(UpdateUI);

                 this.BeginInvoke(stc, "ConnectBleTimeOut", "");
                 Console.Write("ConnectBleTimeOut" + Environment.NewLine);
             }

             // 寫入超過時間
             private void WriteEslDataTimeOut(object sender, EventArgs e)
             {
                 UIInvoker stc = new UIInvoker(UpdateUI);
                 this.BeginInvoke(stc, "ESL_TimeOut", "");
                 Console.Write("ESL_TimeOut" + Environment.NewLine);

             }
             // 取得設備名稱
             private void ReadDeviceNameEven(object sender, EventArgs e)
             {
                 Console.WriteLine("ReadDeviceNameEven");
                 string receivedText = (e as SmcEsl.ReadDeviceNamenEventArgs).name;
                 string deviceIP = (e as SmcEsl.ReadDeviceNamenEventArgs).deviceIP;
                 receivedText = "Device Name : " + receivedText;
                 UIInvoker stc = new UIInvoker(UpdateUI);
                 this.BeginInvoke(stc, receivedText, deviceIP);
             }
             // 寫入設備名稱
             private void WriteDeviceNameEven(object sender, EventArgs e)
             {
                 Console.WriteLine("WriteDeviceNameEven");
                 bool isStatus = (e as SmcEsl.WriteDeviveNameEventArgs).isStatus;
                 string deviceIP = (e as SmcEsl.WriteDeviveNameEventArgs).deviceIP;
                 UIInvoker stc = new UIInvoker(UpdateUI);
                 if (isStatus)
                 {
                     this.BeginInvoke(stc, "DeviceName更新成功", deviceIP);
                 }
                 else
                 {
                     this.BeginInvoke(stc, "DeviceName更新失敗", deviceIP);
                 }
             }
             // 設定翻頁時間
             private void SetEslTurnPageTimeEven(object sender, EventArgs e)
             {
                 Console.WriteLine("SetEslTurnPageTimeEven");
                 bool isStatus = (e as SmcEsl.SetTimeEventArgs).isStatus;
                 string deviceIP = (e as SmcEsl.SetTimeEventArgs).deviceIP;
                 UIInvoker stc = new UIInvoker(UpdateUI);
                 if (isStatus)
                 {
                     this.BeginInvoke(stc, "換頁時間設定成功", deviceIP);
                 }
                 else
                 {
                     this.BeginInvoke(stc, "換頁時間設定失敗", deviceIP);
                 }
             }
             //寫入資料
             private void WriteEslDataEven(object sender, EventArgs e)
             {
                 Console.WriteLine("WriteEslDataEven");
                 bool isStatus = (e as SmcEsl.WriteDataEventArgs).isStatus;
                 string deviceIP = (e as SmcEsl.WriteDataEventArgs).deviceIP;
                 //String qwe = (e as SmcEsl.ScanDeviceEventArgs).address;
                 UIInvoker stc = new UIInvoker(UpdateProgressBar);
                 if (isStatus)
                 {
                     this.BeginInvoke(stc, "資料寫入成功", deviceIP);

                 }
                 else
                 {
                     this.BeginInvoke(stc, "資料寫入失敗", deviceIP);
                 }
             }
             private void WriteEslDataFinishEven(object sender, EventArgs e)
             {
                 Console.WriteLine("WriteEslDataFinishEven");
                 string deviceIP = (e as SmcEsl.WriteEslDataFinishEven).deviceIP;
                 UIInvoker stc = new UIInvoker(UpdateUI);
                 this.BeginInvoke(stc, "資料寫入完成完成完成完成完成完成完成完成", deviceIP);
             }
             // 寫入設備名稱
             private void WriteBeaconEven(object sender, EventArgs e)
             {
                 Console.WriteLine("WriteBeaconEven");
                 bool isStatus = (e as SmcEsl.WriteBeaconEventArgs).isStatus;
                 string deviceIP = (e as SmcEsl.WriteBeaconEventArgs).deviceIP;
                 UIInvoker stc = new UIInvoker(UpdateUI);
                 if (isStatus)
                 {
                     this.BeginInvoke(stc, "Beacon更新成功", deviceIP);
                 }
                 else
                 {
                     this.BeginInvoke(stc, "Beacon更新失敗", deviceIP);
                 }
             }*/
        #endregion

        private void saveFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void tbMessageBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        // This is the method to run when the timer is raised.
        private void TimerEventProcessor(Object myObject, EventArgs myEventArgs)
        {

        }

        private void checkconnect_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先載入資料表");
                return;
            }

            int a = 1;
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                if (a % 2 == 0)
                {
                    dr.Cells[4].Style.BackColor = Color.Beige;
                    dr.Cells[4].Value = DBNull.Value;
                }
                else
                {
                    dr.Cells[4].Style.BackColor = Color.Bisque;
                    dr.Cells[4].Value = DBNull.Value;
                }
                a++;
            }
            dataGridView1.RowsDefaultCellStyle.BackColor = Color.Bisque;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
            int kk = 0;
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                //dr.Cells[4].Style.BackColor = Color.Bisque;
                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != String.Empty)
                {
                    kk++;
                  //  Console.WriteLine("kk" + kk + "ss" + dr.Cells[12].Value.ToString());
                    if (dr.Cells[1].Value.ToString().Length > 14)
                    {
                        string[] drrow = dr.Cells[1].Value.ToString().Split(',');
                        for (int po = 0; po < drrow.Length; po++)
                        {
                            checkaddress.Add(drrow[po]);
                        }
                    }
                    else
                    {
                        checkaddress.Add(dr.Cells[1].Value.ToString());
                    }

                }
            }
            /* if (checkaddress !=null)
             {*/

            CheckConnectTimer.Interval = 3000;
            CheckConnectTimer.Start();
            // }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先載入資料表");
                dataGridView1.Enabled = true;
                return;
            }
            resetbeacon = true;

            Page mPage = new Page();
            mPage.BeaconProduct = "0000000000000";
            BeaconList.Add(mPage);
            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {

                //setBeaconTime(BeaconList[beacon_index].APID);

                kvp.Value.mSmcEsl.WriteBeaconData("ESL143AP01", BeaconList[beacon_index].BeaconProduct, true);
            }
            

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            /* if (!APStart)
             {
                 MessageBox.Show("請先連接AP");
                 dataGridView1.Enabled = true;
                 return;
             }
             */

            Console.WriteLine("END");

            if (e.RowIndex == dataGridView1.RowCount - 2)
            {
                dataGridView1.Rows[e.RowIndex].Cells[15].Value = "X";
                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, e.RowIndex, false, openExcelAddress, excel, excelwb, mySheet);
                dataGridView1.Rows[e.RowIndex].Cells[16].Value = "V";
                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 16, e.RowIndex, false, openExcelAddress, excel, excelwb, mySheet);
            }
        //    if (e.ColumnIndex !=0&&e.ColumnIndex != 3)
           
            //11/30----------------------------------------------------
            /*      if (dataGridView1.Rows[a].Cells[1].Value.ToString() == "")
              {
                  dataGridView1.Rows[a].Cells[1].Value = dataGridView1.RowCount - 1;
              }*/



            Console.WriteLine("e.RowIndex" + e.RowIndex+ "dataGridView1.RowCount" + dataGridView1.RowCount+ "e.ColumnIndex" + e.ColumnIndex);

            if (e.ColumnIndex == 1)
            {

                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Length < 12)
                {
                    //dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                    //dataGridView1.CurrentCell.ErrorText = "需要輸入12位ESL條碼";
                    DialogResult result = MessageBox.Show("需要輸入12位ESL條碼!", "錯誤格式", MessageBoxButtons.OK);
                    if (result == DialogResult.OK)
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                        return;
                    }
                }
                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Length > 12 && dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Length % 13 != 12&& !dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Contains(","))
                {
                    //dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                   // dataGridView1.CurrentCell.ErrorText = "請按格式輸入12位ESL條碼，複數請用,隔開";
                    DialogResult result = MessageBox.Show("請按格式輸入12位ESL條碼，複數請用 , 隔開!", "錯誤格式", MessageBoxButtons.OK);
                    if (result == DialogResult.OK)
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                        return;
                    }
                }

                

                if (editdatagirdcell != dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString())
                {

              

                BackPage bpage = new BackPage();
                bool eslIsNull = false;
                Console.WriteLine("INININ1111");
                dataGridView1.Rows[e.RowIndex].Cells[1].Value= dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().ToUpper();
                if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() != "") {
                if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Length > 13)
                {
                    
                    string[] esllistSplit = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Split(',');
                    for (int i = 0; i < esllistSplit.Length; i++)
                    {
                            bool isEslIDNull = false;
                            foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                            {
                                if (dr4.Cells[1].Value != null && esllistSplit[i] == dr4.Cells[1].Value.ToString())
                                {
                                    isEslIDNull = false;
                                    
                                    break;
                                }
                                isEslIDNull = true;
                            }
                            if (isEslIDNull)
                            {
                                if (dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor == Color.Black)
                                {
                                    dataGridView1.Rows[e.RowIndex].Cells[1].Value = dataGridView1.Rows[e.RowIndex].Cells[12].Value;
                                    dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Black;
                                }
                                else
                                {
                                    dataGridView1.Rows[e.RowIndex].Cells[1].Value = DBNull.Value;
                                }
                                MessageBox.Show("輸入ESL未在選單內，請檢查是否輸入錯誤或掃描未掃到", "輸入錯誤");
                                return;
                            }

                            if (OldEslList.Count != 0)
                            {
                                for (int a = 0; a < OldEslList.Count; a++)
                                {
                                    Console.WriteLine("a" + a + "OldEslList[a].ESLID" + OldEslList[a].ESLID + "esllistSplit[i]" + esllistSplit[i]);
                                    if (OldEslList[a].ESLID == esllistSplit[i]&& OldEslList[a].dataGridRowIndex!=e.RowIndex)
                                    {
                                        eslIsNull = false;
                                        if (dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Contains(',' + esllistSplit[i]))
                                        {
                                            int changeaddr = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().IndexOf(',' + esllistSplit[i]);
                                            dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Remove(changeaddr, 13);
                                            bpage.OldMateProduct = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[5].Value.ToString();
                                            dataGridView1[1, e.RowIndex].Style.ForeColor = Color.Orange;
                                            bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                                            bpage.isBack = false;
                                            backESLList.Add(bpage);
                                            //mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, OldEslList[a].dataGridRowIndex, false, openExcelAddress, excel, excelwb, mySheet);
                                            break;
                                        }
                                        if (dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Contains(esllistSplit[i] + ','))
                                        {

                                            int changeaddr = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().IndexOf(esllistSplit[i] + ',');
                                            dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Remove(changeaddr, 13);
                                            bpage.OldMateProduct = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[5].Value.ToString();
                                            dataGridView1[1, e.RowIndex].Style.ForeColor = Color.Orange;
                                            bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                                            bpage.isBack = false;
                                            backESLList.Add(bpage);
                                            //    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            break;
                                        }

                                        if (dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Contains(esllistSplit[i]))
                                        {

                                            int changeaddr = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().IndexOf(esllistSplit[i]);
                                            dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Remove(changeaddr, 12);
                                            bpage.OldMateProduct = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[5].Value.ToString();
                                            dataGridView1[1, e.RowIndex].Style.ForeColor = Color.Orange;
                                            bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                                            bpage.isBack = false;
                                            backESLList.Add(bpage);
                                            //   mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            break;
                                        }
                                    }
                                    eslIsNull = true;

                                }

                                if (eslIsNull)
                                {
                                    Console.WriteLine("e.RowIndex" + e.RowIndex + "dataGridView1.RowCount" + dataGridView1.RowCount + "e.ColumnIndex" + e.ColumnIndex);
                                    if (dataGridView1.Rows[e.RowIndex].Cells[1].Value != null)
                                    {

                                        if (dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor != Color.Orange)
                                            dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Green;

                                        bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                                        bpage.isBack = false;
                                        backESLList.Add(bpage);

                                    }
                                }

                            }
                            else
                            {
                                if (dataGridView1.Rows[e.RowIndex].Cells[1].Value != null)
                                {

                                    if (dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor != Color.Orange)
                                        dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Green;

                                    bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                                    bpage.isBack = false;
                                    backESLList.Add(bpage);

                                }
                            }



                        }


                }
                else
                {
                    Console.WriteLine("INININ<13");
                    Console.WriteLine(" OldEslList.Count" + OldEslList.Count);

                        bool isEslIDNull = false;
                        foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                        {
                            if (dr4.Cells[1].Value != null && dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() == dr4.Cells[1].Value.ToString())
                            {
                                isEslIDNull = false;

                                break;
                            }
                            isEslIDNull = true;
                        }
                        if (isEslIDNull)
                        {
                            if (dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor == Color.Black)
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[1].Value = dataGridView1.Rows[e.RowIndex].Cells[12].Value;
                                dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Black;
                            }
                            else
                            { 
                                dataGridView1.Rows[e.RowIndex].Cells[1].Value = DBNull.Value;
                            }
                            MessageBox.Show("輸入ESL未在選單內，請檢查是否輸入錯誤或掃描未掃到", "輸入錯誤");
                            return;
                        }

                        if (OldEslList.Count != 0)
                        {
                            for (int a = 0; a < OldEslList.Count; a++)
                            {
                                Console.WriteLine("OldEslList[a].ESLID" + OldEslList[a].ESLID + " dataGridView1[1, e.RowIndex].Value.ToString()" + dataGridView1[1, e.RowIndex].Value.ToString() + e.RowIndex);
                                if (OldEslList[a].ESLID == dataGridView1[1, e.RowIndex].Value.ToString() && OldEslList[a].dataGridRowIndex != e.RowIndex)
                                {

                                    eslIsNull = false;
                                    if (dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Contains(',' + dataGridView1[1, e.RowIndex].Value.ToString()))
                                    {
                                        int changeaddr = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().IndexOf(',' + dataGridView1[1, e.RowIndex].Value.ToString());
                                        dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Remove(changeaddr, 13);
                                        bpage.OldMateProduct = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[5].Value.ToString();
                                        dataGridView1[1, e.RowIndex].Style.ForeColor = Color.Orange;
                                        bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                                        bpage.isBack = false;
                                        backESLList.Add(bpage);
                                        //mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, OldEslList[a].dataGridRowIndex, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                    }
                                    if (dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Contains(dataGridView1[1, e.RowIndex].Value.ToString() + ','))
                                    {

                                        int changeaddr = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().IndexOf(dataGridView1[1, e.RowIndex].Value.ToString() + ',');
                                        dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Remove(changeaddr, 13);
                                        bpage.OldMateProduct = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[5].Value.ToString();
                                        dataGridView1[1, e.RowIndex].Style.ForeColor = Color.Orange;
                                        bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                                        bpage.isBack = false;
                                        backESLList.Add(bpage);
                                        //    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                    }

                                    if (dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Contains(dataGridView1[1, e.RowIndex].Value.ToString()))
                                    {

                                        int changeaddr = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().IndexOf(dataGridView1[1, e.RowIndex].Value.ToString());
                                        dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[1].Value.ToString().Remove(changeaddr, 12);
                                        bpage.OldMateProduct = dataGridView1.Rows[OldEslList[a].dataGridRowIndex].Cells[5].Value.ToString();
                                        dataGridView1[1, e.RowIndex].Style.ForeColor = Color.Orange;
                                        bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                                        bpage.isBack = false;
                                        backESLList.Add(bpage);
                                        //   mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                    }
                                }
                                eslIsNull = true;

                            }

                            if (eslIsNull)
                            {
                                Console.WriteLine("e.RowIndex" + e.RowIndex + "dataGridView1.RowCount" + dataGridView1.RowCount + "e.ColumnIndex" + e.ColumnIndex);
                                if (dataGridView1.Rows[e.RowIndex].Cells[1].Value != null)
                                {

                                    if (dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor != Color.Orange)
                                        dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Green;

                                    bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                                    bpage.isBack = false;
                                    backESLList.Add(bpage);

                                }
                            }

                        }
                        else
                        {
                            if (dataGridView1.Rows[e.RowIndex].Cells[1].Value != null)
                            {

                                if (dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor != Color.Orange)
                                    dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Green;

                                bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                                bpage.isBack = false;
                                backESLList.Add(bpage);

                            }
                        }
                    }

                }

                }




                /*      foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                      {
                          Console.WriteLine("dr.Cells[1].Value.ToString()x" + dr.Cells[1].Value.ToString() + "dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString()" + dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
                          if (dr.Cells[1].Value.ToString() == dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString())
                          {
                              if (dr.Cells[1].Value.ToString().Length > 1)
                              {
                                  dataGridView1[1, rowIndex].Style.ForeColor = Color.Orange;
                                  //      mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                  //dataGridView1[2, rowIndex].Value = "已綁定";
                              }
                              else
                              {
                                  dataGridView1[1, rowIndex].Style.ForeColor = Color.Orange;
                                  //      mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                  //dataGridView1[2, rowIndex].Value = "已綁定";
                              }
                              eslIsNull = false;
                              bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                              bpage.isBack = false;
                              backESLList.Add(bpage);
                              break;
                          }

                          eslIsNull = true;

                      }
                      Console.WriteLine("eslIsNull" + eslIsNull);
                      if (eslIsNull) {
                          Console.WriteLine("e.RowIndex" + e.RowIndex + "dataGridView1.RowCount" + dataGridView1.RowCount + "e.ColumnIndex" + e.ColumnIndex);
                          if (dataGridView1.Rows[e.RowIndex].Cells[1].Value != null)
                      {

                              if (dataGridView1.Rows[e.RowIndex].Cells[1].Value != null && dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString().Length >= 12)
                          {

                              if (dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor != Color.Orange)
                                  dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Green;
                          }
                          else
                          {
                              if (dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor != Color.Orange)
                                  dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor = Color.Green;

                          }
                              bpage.NewMateESL = dataGridView1[1, e.RowIndex].Value.ToString();
                              bpage.isBack = false;
                              backESLList.Add(bpage);

                          }
                      }*/

            }


            if (e.ColumnIndex == 5)
            {
                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value!=null&&dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()!="" )
                {
                    if(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Length != 13) {
                    //dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                    dataGridView1.CurrentCell.ErrorText = "需要輸入13位數字";
                    DialogResult result = MessageBox.Show("barcode需要輸入13位數字!", "Confirmation", MessageBoxButtons.YesNoCancel);
                    if (result == DialogResult.Yes)
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                            return;    
                    }
                    }
                }
            }
            else if (e.ColumnIndex == 11)
            {

                if (!Regex.IsMatch(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), @"^http(s)?://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)?$"))
                {
                    // dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                    dataGridView1.CurrentCell.ErrorText = "請輸入網址";
                    DialogResult result = MessageBox.Show("請輸入網址!", "Confirmation", MessageBoxButtons.YesNoCancel);
                    if (result == DialogResult.Yes)
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = "";
                        return;
                    }

                }
            }

            if (e.ColumnIndex == 5 || e.ColumnIndex == 6 || e.ColumnIndex == 7 || e.ColumnIndex == 8 || e.ColumnIndex == 9 || e.ColumnIndex == 10 || e.ColumnIndex == 11)
            {
                if (e.ColumnIndex == 5)
                {
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {
                        
                    }
                }
                if(dataGridView1.Rows[e.RowIndex].Cells[1].Value != null && editdatagirdcell!= dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString())
                {
                    if (dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() != ""&& dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor!=Color.Orange&& dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor != Color.Green)
                    {
                        if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null && dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() != textBeforeEdit)
                            {
                                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.ForeColor = Color.Red;
                                button2.Enabled = true;
                                button2.BackColor = Color.SpringGreen;


                            }
                        Console.WriteLine(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() + "ININ1" + dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor);
                    }
                    else
                    {
                        Console.WriteLine(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() + "ININ2"+ dataGridView1.Rows[e.RowIndex].Cells[1].Style.ForeColor);
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, e.ColumnIndex, e.RowIndex, false, openExcelAddress, excel, excelwb, mySheet);
                    }
                }
            }
                if (productAll.Text != dataGridView1.Rows.Count.ToString())
                productAll.Text = (dataGridView1.Rows.Count - 1).ToString();
           // mExcelData.DataGridviewSave(dataGridView1, false, openExcelAddress);

        }


        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs anError)
        {

            switch (anError.ColumnIndex)
            {
                case 5:

                    MessageBox.Show("barcode需填數字");
                    break;
                case 9:
                    MessageBox.Show("售價需填數字");
                    break;
                case 10:
                    MessageBox.Show("促銷價需填數字");
                    break;


            }
            //dataGridView1.Rows[anError.RowIndex].Cells[anError.ColumnIndex].Value = "";

            // MessageBox.Show("Error happened " + anError.Context.ToString());
            /*  MessageBox.Show("Error happened " + anError.Context.ToString() + dataGridView1.Rows[anError.RowIndex].Cells[anError.ColumnIndex].Value.ToString()+ anError.ColumnIndex);
              //11/30----------------------------------------------------
              if (dataGridView1.Rows[a].Cells[1].Value.ToString() == "")
              {
                  dataGridView1.Rows[a].Cells[1].Value = dataGridView1.RowCount - 1;
              }

              if (anError.Context == DataGridViewDataErrorContexts.Commit)
                 {
                     MessageBox.Show("Commit error");
                 }
                 if (anError.Context == DataGridViewDataErrorContexts.CurrentCellChange)
                 {
                     MessageBox.Show("Cell change");
                 }
                 if (anError.Context == DataGridViewDataErrorContexts.Parsing)
                 {
                     MessageBox.Show("parsing error");
                 }
                 if (anError.Context == DataGridViewDataErrorContexts.LeaveControl)
                 {
                     MessageBox.Show("leave control error");
                 }

                 if ((anError.Exception) is ConstraintException)
                 {
                     DataGridView view = (DataGridView)sender;
                     view.Rows[anError.RowIndex].ErrorText = "an error";
                     view.Rows[anError.RowIndex].Cells[anError.ColumnIndex].ErrorText = "an error";

                     anError.ThrowException = false;
                 }*/
            return;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {

                kvp.Value.mSmcEsl.stopScanBleDevice();
            }
            if (!testest)
            {
                nullMsg = null;
                removeESLingstate = true;
                string eslAPNoSetMsg = null;
                stopwatch.Reset();
            stopwatch.Start();
            PageList.Clear();
            UpdateESLDen.Text = "0";
            updateESLper.Text = "0";
          //  UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
            listcount = 0;
            dataGridView1.Enabled = false;
            dataGridView1.ClearSelection();
            // mSmcEsl.DisConnectBleDevice();
          //  Console.WriteLine("PageList" + PageList);
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先載入資料表");
                dataGridView1.Enabled = true;
                return;
            }

            if (!APStart)
            {
                MessageBox.Show("請先連接AP");
                dataGridView1.Enabled = true;
                return;
            }

      /*      if (styleName == null)
            {
                if (dataGridView2.Rows.Count > 1)
                {
                    foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                    {
                      //  Console.WriteLine(dr.Cells[1].RowIndex + dr.Cells[1].Value.ToString());

                        if (dr.Cells[1].RowIndex == 0)
                        {

                            for (int i = 0; i < dr.Cells.Count; i++)
                            {

                                if (dr.Cells[i].Value != null && dr.Cells[i].Value.ToString() != "")
                                {
                                 //   Console.WriteLine("HGEEGE");
                                    if (i == 1)
                                    {
                                        styleName = dr.Cells[1].Value.ToString();
                                    }

                                    if (i != 0 && i != 1)
                                    {

                                        ESLFormat.Add(dr.Cells[i].Value.ToString());

                                      //  Console.WriteLine(dr.Cells[i].Value.ToString());
                                    }
                                }
                            }

                        }
                    }
                }
            }*/

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    if (dr.Cells[1].Value.ToString() != "")
                    {
                        if (dr.Cells[1].Value.ToString().Length > 13)
                        {

                            string[] usingAddressSplit = dr.Cells[1].Value.ToString().Split(',');
                            //   string Special_offer;
                            /*    if (dr.Cells[15].Value.ToString() == "V")
                                {
                                    dr.Cells[15].Value = "X";
                                    Special_offer = dr.Cells[10].Value.ToString();
                                }
                                else
                                {
                                    dr.Cells[15].Value = "V";
                                    Special_offer = dr.Cells[10].Value.ToString();

                                }*/
                            for (int i = 0; i < usingAddressSplit.Length; i++)
                            {
                                Page1 mPageA = new Page1();
                                UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                mPageA.no = (dr.Index + 1).ToString();
                                mPageA.BleAddress = dr.Cells[1].Value.ToString();
                                mPageA.barcode = dr.Cells[5].Value.ToString();
                                mPageA.product_name = dr.Cells[6].Value.ToString();
                                mPageA.Brand = dr.Cells[7].Value.ToString();
                                mPageA.specification = dr.Cells[8].Value.ToString();
                                mPageA.price = dr.Cells[9].Value.ToString();

                                mPageA.Web = dr.Cells[11].Value.ToString();
                                mPageA.usingAddress = usingAddressSplit[i];
                                mPageA.HeadertextALL = headertextall;
                                mPageA.Special_offer = dr.Cells[10].Value.ToString();
                                mPageA.onsale = dr.Cells[15].Value.ToString();
                                foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                {
                                    if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == usingAddressSplit[i])
                                    {
                                        mPageA.APLink = drAP.Cells[8].Value.ToString();
                                        break;
                                    }
                                }

                                    foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                    {
                                        if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPageA.APLink)
                                        {
                                            if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                            {
                                                PageList.Add(mPageA);
                                            }
                                            else
                                            {
                                                if (eslAPNoSetMsg == null)
                                                    eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                                if (eslAPNoSetMsg != null)
                                                    eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                            }
                                        }
                                    }
                                }
                        }
                        else
                        {
                            //Console.WriteLine("dr.Cells[6].Value.ToString()" + dr.Cells[6].Value.ToString());
                            Page1 mPageC = new Page1();
                            mPageC.no = (dr.Index + 1).ToString();
                            mPageC.BleAddress = dr.Cells[1].Value.ToString();
                            mPageC.barcode = dr.Cells[5].Value.ToString();
                            mPageC.product_name = dr.Cells[6].Value.ToString();
                            mPageC.Brand = dr.Cells[7].Value.ToString();
                            mPageC.specification = dr.Cells[8].Value.ToString();
                            mPageC.price = dr.Cells[9].Value.ToString();

                            mPageC.Web = dr.Cells[11].Value.ToString();
                            mPageC.usingAddress = dr.Cells[1].Value.ToString();
                            mPageC.HeadertextALL = headertextall;
                            //mPageC.Special_offer = dr.Cells[10].Value.ToString();
                            mPageC.Special_offer = dr.Cells[10].Value.ToString();
                            mPageC.onsale = dr.Cells[15].Value.ToString();

                            foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                            {
                                if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == dr.Cells[1].Value.ToString())
                                {
                                    mPageC.APLink = drAP.Cells[8].Value.ToString();
                                    break;
                                }
                            }
                                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                {
                                    if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPageC.APLink)
                                    {
                                        if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                        {
                                            PageList.Add(mPageC);
                                        }
                                        else
                                        {
                                            if (eslAPNoSetMsg == null)
                                                eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                            if (eslAPNoSetMsg != null)
                                                eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                        }
                                    }
                                }
                            }

                    }
                    else
                    {
                        if (nullMsg == null)
                        {
                            nullMsg = dr.Cells[6].Value.ToString();
                        }
                        else
                        {
                            nullMsg = nullMsg + "," + dr.Cells[6].Value.ToString();
                        }
                    }

                }
            }

            

            if (nullMsg != null)
            {
               // MessageBox.Show("勾選" + nullMsg + "未綁定電子標籤");
                PageList.Clear();
                dataGridView1.Enabled = true;
                nullMsg = null;
                datagridview1curr = 2;
                aaa(1, false, 0);
                    DialogResult dialogResult = MessageBox.Show("勾選" + nullMsg + "未綁定電子標籤" + "\r\n" + "是否繼續綁定", "未綁定", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }

                if (eslAPNoSetMsg != null)
                {

                 //   MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用");

                    dataGridView1.Enabled = true;
                    DialogResult dialogResult = MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用" + "\r\n" + "是否繼續綁定", "AP未啟用", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                    //return;
                }

                // ------------------明天 初始修改
                if (PageList.Count!=0) {
                down = true;
                testest = true;
                List<string> RunAPList = new List<string>();
                List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).ToList();
                foreach (Page1 p in list)
                {
                    RunAPList.Add(p.APLink);

                }

                

                   
                    for (int a = 0; a < RunAPList.Count; a++)
                    {
                        for (int i = 0; i < PageList.Count; i++)
                        {
                            if (PageList[i].APLink == RunAPList[a])
                            {
                                Page1 mPage1 = PageList[i];
                            if (mPage1.usingAddress != "")
                            {
                                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                                {

                                    if (kvp.Key.Contains(mPage1.APLink))
                                    {

                                        Bitmap bmp = mElectronicPriceData.writeIDimage(PageList[listcount].usingAddress);
                                        int numVal = Convert.ToInt32(mPage1.no) - 1;
                                        // Console.WriteLine("mPage1.no" + mPage1.no);
                                        dataGridView1.Rows[numVal].Selected = true;
                                        dataGridView1.Rows[numVal].Cells[17].Value = DateTime.Now.ToString();

                                        pictureBoxPage1.Image = bmp;
                                        // Console.WriteLine("ININ");
                                        kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                        kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress,0);
                                        //  System.Threading.Thread.Sleep(100);
                                        pictureBoxPage1.Image = bmp;
                                        EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                        mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3,0);
                                        richTextBox1.Text = mPage1.usingAddress + "  嘗試連線中請稍候... \r\n";
                                    }
                                }
                                break;
                            }
                            else
                            {
                                    
                                    MessageBox.Show("該商品" + mPage1.product_name + "未裝置電子標籤");
                                dataGridView1.Enabled = true;
                                down = false;
                            }
                        }
                    }
                    // mSmcEsl.TransformImageToData(bmp);
                    

                    //   mSmcEsl.ConnectBleDevice(mPage1.usingAddress);
                    //  mSmcEsl.WriteESLData(mPage1.usingAddress);
                 //   Console.WriteLine("listcount" + listcount);
                    macaddress = PageList[listcount].usingAddress;
                    


                }
                
            }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            //UpdateESLDen.Text=(Convert.ToInt32(UpdateESLDen.Text)+1).ToString();
            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {

                    kvp.Value.mSmcEsl.stopScanBleDevice();
            }


            if (!testest) {
                nullMsg = null;
                UpdateESLDen.Text = "0";
            updateESLper.Text = "0";
            onsaleESLingstate = true;
                string eslAPNoSetMsg = null;
                string eslVState = null;
                string eslNotMateAP = null;
                stopwatch.Reset();
            stopwatch.Start();
            PageList.Clear();
            listcount = 0;
            dataGridView1.ClearSelection();
            // mSmcEsl.DisConnectBleDevice();
           // Console.WriteLine("PageList" + PageList);
            
            if (dataGridView1.RowCount == 0)    
            {
                MessageBox.Show("請先載入資料表");
                dataGridView1.Enabled = true;
                return;
            }

            if (!APStart)
            {
                MessageBox.Show("請先連接AP");
                dataGridView1.Enabled = true;
                return;
            }

            System.Threading.Thread.Sleep(100);


          /*  if (styleName == null)
            {
                if (dataGridView2.Rows.Count > 1)
                {
                    foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                    {
                      //  Console.WriteLine(dr.Cells[1].RowIndex + dr.Cells[1].Value.ToString());

                        if (dr.Cells[1].RowIndex == 0)
                        {

                            for (int i = 0; i < dr.Cells.Count; i++)
                            {

                                if (dr.Cells[i].Value != null && dr.Cells[i].Value.ToString() != "")
                                {
                               //     Console.WriteLine("HGEEGE");
                                    if (i == 1)
                                    {
                                        styleName = dr.Cells[1].Value.ToString();
                                    }

                                    if (i != 0 && i != 1)
                                    {

                                        ESLFormat.Add(dr.Cells[i].Value.ToString());

                                     //   Console.WriteLine(dr.Cells[i].Value.ToString());
                                    }
                                }
                            }

                        }
                    }
                }
            }*/

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                Console.WriteLine("dr.Cells[0].Value" + dr.Cells[0].Value);
                if (dr.Cells[5].Style.ForeColor==Color.Red  || dr.Cells[5].Style.ForeColor == Color.Red || dr.Cells[6].Style.ForeColor == Color.Red || dr.Cells[7].Style.ForeColor == Color.Red || dr.Cells[8].Style.ForeColor == Color.Red || dr.Cells[9].Style.ForeColor == Color.Red || dr.Cells[10].Style.ForeColor == Color.Red || dr.Cells[11].Style.ForeColor == Color.Red)
                {
                    if (dr.Cells[1].Value.ToString() != "")
                    {
                        if (dr.Cells[1].Value.ToString().Length > 13)
                        {

                            string[] usingAddressSplit = dr.Cells[1].Value.ToString().Split(',');
                         //   string Special_offer;
                        /*    if (dr.Cells[15].Value.ToString() == "V")
                            {
                                dr.Cells[15].Value = "X";
                                Special_offer = dr.Cells[10].Value.ToString();
                            }
                            else
                            {
                                dr.Cells[15].Value = "V";
                                Special_offer = dr.Cells[10].Value.ToString();

                            }*/
                            for (int i = 0; i < usingAddressSplit.Length; i++)
                            {
                                Page1 mPageA = new Page1();
                                UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                mPageA.no = (dr.Index + 1).ToString();
                                mPageA.BleAddress = dr.Cells[1].Value.ToString();
                                mPageA.barcode = dr.Cells[5].Value.ToString();
                                mPageA.product_name = dr.Cells[6].Value.ToString();
                                mPageA.Brand = dr.Cells[7].Value.ToString();
                                mPageA.specification = dr.Cells[8].Value.ToString();
                                mPageA.price = dr.Cells[9].Value.ToString();

                                mPageA.Web = dr.Cells[11].Value.ToString();
                                mPageA.usingAddress = usingAddressSplit[i];
                                mPageA.HeadertextALL = headertextall;
                                mPageA.Special_offer = dr.Cells[10].Value.ToString();
                                mPageA.onsale = dr.Cells[15].Value.ToString();
                                mPageA.ProductStyle = dr.Cells[13].Value.ToString();
                                    mPageA.TimerConnect = new System.Windows.Forms.Timer();
                                    mPageA.TimerConnect.Interval = (30 * 1000);
                                    mPageA.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                    mPageA.TimerSeconds = new Stopwatch();

                                    mPageA.actionName = "sale";
                               foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                {
                                    if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == usingAddressSplit[i])
                                    {

                                            if (drAP.Cells[8].Value.ToString() == "")
                                            {
                                                if (eslNotMateAP == null)
                                                    eslNotMateAP = usingAddressSplit[i];
                                                else
                                                    eslNotMateAP = eslNotMateAP + "," + usingAddressSplit[i];
                                                // MessageBox.Show("請先配對ESL IP");
                                                //break;
                                            }
                                            if (drAP.Cells[5].Value.ToString() != "" && Convert.ToDouble(drAP.Cells[5].Value) < 2.85)
                                            {
                                                if (eslVState == null)
                                                    eslVState = usingAddressSplit[i];
                                                else
                                                    eslVState = eslVState + "," + usingAddressSplit[i];
                                                // MessageBox.Show("請先配對ESL IP");
                                                //break;
                                            }
                                            mPageA.APLink = drAP.Cells[8].Value.ToString();
                                        break;
                                    }
                                }

                                    foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                    {
                                        if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPageA.APLink)
                                        {
                                            if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                            {
                                                PageList.Add(mPageA);
                                            }
                                            else
                                            {
                                                if (eslAPNoSetMsg == null)
                                                    eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                                if (eslAPNoSetMsg != null)
                                                    eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                            }
                                        }
                                    }
                                }
                        }
                        else
                        {
                            //Console.WriteLine("dr.Cells[6].Value.ToString()" + dr.Cells[6].Value.ToString());
                            Page1 mPageC = new Page1();
                            mPageC.no = (dr.Index + 1).ToString();
                            mPageC.BleAddress = dr.Cells[1].Value.ToString();
                            mPageC.barcode = dr.Cells[5].Value.ToString();
                            mPageC.product_name = dr.Cells[6].Value.ToString();
                            mPageC.Brand = dr.Cells[7].Value.ToString();
                            mPageC.specification = dr.Cells[8].Value.ToString();
                            mPageC.price = dr.Cells[9].Value.ToString();

                            mPageC.Web = dr.Cells[11].Value.ToString();
                            mPageC.usingAddress = dr.Cells[1].Value.ToString();
                            mPageC.HeadertextALL = headertextall;
                            //mPageC.Special_offer = dr.Cells[10].Value.ToString();
                            mPageC.Special_offer = dr.Cells[10].Value.ToString();
                            mPageC.onsale = dr.Cells[15].Value.ToString();
                            mPageC.ProductStyle = dr.Cells[13].Value.ToString();
                            mPageC.TimerConnect = new System.Windows.Forms.Timer();
                            mPageC.TimerConnect.Interval = (30 * 1000);
                            mPageC.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                            mPageC.TimerSeconds = new Stopwatch();
                            mPageC.actionName = "sale";

                            foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                            {
                                if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == dr.Cells[1].Value.ToString())
                                {

                                        if (drAP.Cells[8].Value.ToString() == "")
                                        {
                                            if (eslNotMateAP == null)
                                                eslNotMateAP = dr.Cells[1].Value.ToString();
                                            else
                                                eslNotMateAP = eslNotMateAP + "," + dr.Cells[1].Value.ToString();
                                            // MessageBox.Show("請先配對ESL IP");
                                            //break;
                                        }
                                        if (drAP.Cells[5].Value.ToString() != "" && Convert.ToDouble(drAP.Cells[5].Value) < 2.85)
                                        {
                                            if (eslVState == null)
                                                eslVState = dr.Cells[1].Value.ToString();
                                            else
                                                eslVState = eslVState + "," + dr.Cells[1].Value.ToString();
                                            // MessageBox.Show("請先配對ESL IP");
                                            //break;
                                        }

                                        mPageC.APLink = drAP.Cells[8].Value.ToString();
                                    break;
                                }
                            }
                                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                {
                                    if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPageC.APLink)
                                    {
                                        if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                        {
                                            PageList.Add(mPageC);
                                        }
                                        else
                                        {
                                            if (eslAPNoSetMsg == null)
                                                eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                            if (eslAPNoSetMsg != null)
                                                eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                        }
                                    }
                                }
                            }

                    }
                    else {
                        if (nullMsg == null)
                        {
                            nullMsg = dr.Cells[6].Value.ToString();
                        }
                        else
                        {
                            nullMsg = nullMsg + "," + dr.Cells[6].Value.ToString();
                        }
                    }

                }
            }

            

            if (nullMsg != null)
            {
              //  MessageBox.Show("勾選" + nullMsg + "未綁定電子標籤");
                PageList.Clear();
                dataGridView1.Enabled = true;
                nullMsg = null;
                datagridview1curr = 2;
                aaa(1,false,0);
                DialogResult dialogResult = MessageBox.Show( "勾選" + nullMsg + "未綁定電子標籤" + "\r\n" + "是否繼續執行", "未綁定", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }

                    
            }


                if (eslNotMateAP != null)
                {
                    //MessageBox.Show("勾選"+ nullMsg + "未綁定電子標籤");
                    PageList.Clear();
                    dataGridView1.Enabled = true;
                    nullMsg = null;
                    datagridview1curr = 2;
                    aaa(1, false, 0);
                    DialogResult dialogResult = MessageBox.Show(eslNotMateAP + "未配對AP請自動配對" + "\r\n" + "是否繼續執行", "未配對ESL", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }

                if (eslVState != null)
                {
                    //MessageBox.Show("勾選"+ nullMsg + "未綁定電子標籤");
                    PageList.Clear();
                    dataGridView1.Enabled = true;
                    nullMsg = null;
                    datagridview1curr = 2;
                    aaa(1, false, 0); 
                    DialogResult dialogResult = MessageBox.Show(eslVState + "電壓未達2.85V" + "\r\n" + "是否繼續執行", "電壓", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }

                if (eslAPNoSetMsg != null)
                {
                   
                   // MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用");

                    dataGridView1.Enabled = true;
                    DialogResult dialogResult = MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用" + "\r\n" + "是否繼續執行", "AP未啟用", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                    else if (PageList.Count == 1)
                    {
                        return;
                    }
                    //return;
                }
                // ------------------明天 初始修改

                if (PageList.Count != 0) {
                sale = true;
                testest = true;
                onlockedbutton(testest);
                    //pictureBox4.Visible = true;
                    dataGridView1.Enabled = false;
                UpdateESLDen.Text = PageList.Count.ToString();
                ProgressBarVisible(PageList.Count);
                List<string> RunAPList = new List<string>();
                List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).ToList();
                foreach (Page1 p in list)
                {
                    RunAPList.Add(p.APLink);

                }



                ESLFormatUpdate.Clear();
                ESLSaleFormatUpdate.Clear();
                for (int a = 0; a < RunAPList.Count; a++)
                    {
                        for (int i = 0; i < PageList.Count; i++)
                        {
                            if (PageList[i].APLink == RunAPList[a])
                            {
                            Page1 mPage1 = PageList[i];
                            if (mPage1.usingAddress != "")
                            {
                                
                                int Blcount = mPage1.BleAddress.Length;
                                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                                {

                                    if (kvp.Key.Contains(mPage1.APLink))
                                    {

                                        foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
                                        {
                                            if (dr2.Cells[1].Value != null)
                                            {
                                                if (dr2.Cells[2].Value.ToString() == "V")
                                                {
                                                    for (int c = 0; c < dr2.Cells.Count; c++)
                                                    {
                                                       // Console.WriteLine(i + "dr2.Cells[i].Value" + dr2.Cells[i].Value);
                                                        if (c != 0)
                                                        {
                                                            if (dr2.Cells[c].Value != null && dr2.Cells[c].Value.ToString() != "")
                                                            {
                                                                Console.WriteLine("HGEEGE");
                                                                if (c == 1)
                                                                {
                                                                    styleName = dr2.Cells[1].Value.ToString();
                                                                }
                                                                if (c != 0 && c != 1 && c != 2)
                                                                {
                                                                    ESLFormatUpdate.Add(dr2.Cells[c].Value.ToString());
                                                                    Console.WriteLine(dr2.Cells[c].Value.ToString());
                                                                }
                                                            }
                                                            else
                                                            {
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }

                                            }
                                        }


                                        foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
                                        {
                                            if (dr7.Cells[1].Value != null)
                                            {
                                                if (dr7.Cells[2].Value.ToString() == "V")
                                                {
                                                    for (int c = 0; c < dr7.Cells.Count; c++)
                                                    {
                                                        // Console.WriteLine(i + "dr2.Cells[i].Value" + dr2.Cells[i].Value);
                                                        if (c != 0)
                                                        {
                                                            if (dr7.Cells[c].Value != null && dr7.Cells[c].Value.ToString() != "")
                                                            {
                                                                Console.WriteLine("HGEEGE");
                                                                if (c == 1)
                                                                {
                                                                    styleSaleName = dr7.Cells[1].Value.ToString();
                                                                }
                                                                if (c != 0 && c != 1 && c != 2)
                                                                {
                                                                    ESLSaleFormatUpdate.Add(dr7.Cells[c].Value.ToString());
                                                                    Console.WriteLine(dr7.Cells[c].Value.ToString());
                                                                }
                                                            }
                                                            else
                                                            {
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                       /*     Bitmap bmp;
                                            if (mPage1.onsale == "V")
                                            {
                                                 bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                        mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormatUpdate);
                                            }
                                            else
                                            {
                                                 bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                        mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormatUpdate);
                                            }*/
                                       
                                        //mSmcEsl.TransformImageToData(bmp);
                                        int numVal = Convert.ToInt32(mPage1.no) - 1;
                                        // Console.WriteLine("mPage1.no" + mPage1.no);
                                        // Console.WriteLine("mPage1.APLink" + mPage1.APLink);
                                        //dataGridView1.ClearSelection();
                                        dataGridView1.Rows[numVal].Cells[0].Selected = true;
                                        aaa(datagridview1curr, true, numVal);
                                        dataGridView1.Rows[numVal].Cells[17].Value = DateTime.Now.ToString();
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, numVal, false, openExcelAddress, excel, excelwb, mySheet);
                                      //  pictureBoxPage1.Image = bmp;
                                            //    Console.WriteLine("ININ");
                                            //  kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                            //    kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress,0);
                                            // System.Threading.Thread.Sleep(100);
                                            //  EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                            //pictureBoxPage1.Image = bmp;
                                            //   mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 8,0);
                                        /*    mPage1.TimerConnect = new System.Windows.Forms.Timer();
                                            mPage1.TimerConnect.Interval = (30 * 1000);
                                            mPage1.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                            mPage1.TimerSeconds = new Stopwatch();*/
                                            mPage1.TimerSeconds.Start();
                                            mPage1.TimerConnect.Start();
                                            kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                                            richTextBox1.Text = mPage1.usingAddress + "  嘗試連線中請稍候... \r\n";
                                    }
                                }
                                break;
                            }
                            else
                            {
                                MessageBox.Show("該商品" + mPage1.product_name + "未裝置電子標籤");
                                dataGridView1.Enabled = true;
                                sale = false;
                            }
                        }
                        }
                    }

                    // dataGridView3.Rows[datagridview2no].Cells[4].Value = "連線中";
                  
                    //  mSmcEsl.TransformImageToData(bmp);
                   
                    // mSmcEsl.ConnectBleDevice(mPage1.usingAddress);

                    // mSmcEsl.WriteESLData(mPage1.usingAddress);
                    macaddress = PageList[listcount].usingAddress;
                    string sub = Environment.CurrentDirectory;
                  //  Console.WriteLine("sub" + sub);
                    
                   // dataGridView3.Rows[0].Cells[4].Value = "連線中";
            



                
            }
            }
            else
            {
                MessageBox.Show("ESL更新中請稍後", "更新中");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2();//產生Form2的物件，才可以使用它所提供的Method
            f.ShowDialog(this);//設定Form2為Form1的上層，並開啟Form2視窗。由於在Form1的程式碼內使用this，所以this為Form1的物件本身
        }

        private void delete_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先載入資料表");
                return;
            }

            if (!testest) { 
            int relrowno = 0;
            List<DataGridViewRow> toDelete = new List<DataGridViewRow>();
            List<int> deldataview1no = new List<int>();
            DialogResult result = MessageBox.Show("該欄位是否刪除?", "Confirmation", MessageBoxButtons.YesNoCancel);
            if (result == DialogResult.Yes)
            { 
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[0].Value != null && dr.Cells[0].Value.ToString() == "True")
                {
                        deldataview1no.Add(dr.Cells[0].RowIndex + 2 - relrowno);
                            relrowno++;
                            toDelete.Add(dr);

                    }
                  //  Console.WriteLine("111111");
            }

            foreach (DataGridViewRow row in toDelete)
            {

               
                dataGridView1.Rows.Remove(row);
                    productAll.Text = (Convert.ToInt32(productAll.Text) - 1).ToString();
                  //  Console.WriteLine("2222222");
                }
             //   Console.WriteLine("3333333");
                mExcelData.dataviewdel(dataGridView1, deldataview1no, "工作表1", openExcelAddress, excel, excelwb, mySheet);
              //  mExcelData.DataGridviewSave(dataGridView1, false, openExcelAddress);
            }
            }
        }

        private void FirstBuild_Click(object sender, EventArgs e)
        {

        }

        static int VALIDATION_DELAY = 80; // Delay Timer (ms)

        System.Threading.Timer timer = null;

        private void scancode_TextChanged(object sender, EventArgs e)
        {
            Console.WriteLine("TextChanged");
            //      TextBox origin = sender as TextBox;

            //    if (!origin.ContainsFocus)

            //       return;


            //   DisposeTimer();

            //  timer = new System.Threading.Timer(TimerElapsed, null, VALIDATION_DELAY, VALIDATION_DELAY);


        }


        private void TimerElapsed(Object obj)
        {

            CheckSyntaxAndReport();

            DisposeTimer();
        }



        private void DisposeTimer()
        {

            if (timer != null)
            {

                timer.Dispose();

                timer = null;

            }

        }


        private void CheckSyntaxAndReport()
        {

      //      this.Invoke(new Action(() =>
      //      {
            string decoded = scancode.Text.ToUpper(); //Do everything on the UI thread itself
                                                      // label1.Text = s;


     /*       bool BoolValue = false;

            for (int i = 0; i < decoded.Length; i++)
            {
                Regex rx = new Regex("^[\u4e00-\u9fa5]$");
　　        if (rx.IsMatch(decoded[i].ToString()))
            {
　　          BoolValue = false;
              Console.WriteLine("EEEEEEEEEEEEEEE" + decoded);
                        break;
            }
　　        else
　　        {
　　          BoolValue = true;
　　        }
　　}*/

             //   Console.WriteLine("decoded" + decoded);
              //  Console.WriteLine("scancodeas" + scancodeas);
              
            if (decoded != scancodeas)
                {
                   
                    if (decoded.Length >= 12)
                    {
                        packet.Clear();
                        Console.WriteLine(decoded);

                        decimal number3 = 0;
                        Boolean canConvert = decimal.TryParse(decoded, out number3);
                    Console.WriteLine("canConvert "+ canConvert);
                    if (onetwo == false)
                        {
                           

                            rowIndex = 0;
                            if (canConvert == true)
                            {
                                Console.WriteLine("@@@@@@@@@  ");
                                doubletype = false;
                                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                {
                                    if (dr.Cells[5].Value!=null&&dr.Cells[5].Value.ToString() == decoded)
                                    {
                                        onetwo = true;
                                        MacAddressList.Add(decoded);
                                        Console.WriteLine("@@@@   " + rowIndex);
                                        break;
                                    }
                                    rowIndex++;
                                }
                            }
                            else
                            {
                                onetwo = true;
                                doubletype = true;
                                dataTemp = decoded;
                                Console.WriteLine("XXXXX  ");
                                bool rowadd = false;
                                //dataGridView1[1, rowIndex].Value = decoded;
                             /*   foreach (DataGridViewRow dr in this.dataGridView4.Rows) {
                                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == dataTemp)
                                    {
                                        rowadd = false;
                                    }
                                    else
                                    {
                                        rowadd = true;
                                        break;
                                    }
                                }
                                if (rowadd) {
                                    dataGridView4.Rows[dataGridView1.Rows.Count - 1].Cells[1].Value = dataTemp;
                                   // dataGridView4.Rows.Add("", dataTemp,"","","","","未綁定");
                                }*/
                            }
                        }
                        else
                        {
                          //  Console.WriteLine("dfffffffffffffffff");

                            if (canConvert == true && doubletype == true)
                            {
                               Console.WriteLine("zzzzzzzzzzzzzz");

                                Boolean maccheck = false;
                                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                {
                                    if (dr.Cells[1].Value!=null&&dr.Cells[1].Value.ToString().Contains(dataTemp))
                                    {
                                    Console.WriteLine("BBBBBBBBBB");
                                    maccheck = true;
                                    break;
                                    }
                                }
                                rowIndex = 0;
                                if (maccheck == false)
                                {
                                BackPage bpage = new BackPage();
                                Console.WriteLine("QQQQQQQQq");
                                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                    {
                                        if (dr.Cells[5].Value!=null&&dr.Cells[5].Value.ToString() == decoded)
                                        {
                                        //dataGridView1[1, rowIndex].Value = dataTemp;

                                        Console.WriteLine("fffffff");
                                            //----------------------------------
                                            if (dr.Cells[1].Value != null&&dr.Cells[1].Value.ToString().Length >= 12)
                                            {
                                            //    dataGridView1[2, rowIndex].Style.BackColor=Color.Green;
                                            //    dataGridView1[2, rowIndex].Value = "未綁定";
                                                dataGridView1[1, rowIndex].Value = dataGridView1[1, rowIndex].Value + "," + dataTemp;
                                            if (dr.Cells[1].Style.ForeColor != Color.Orange)
                                                dr.Cells[1].Style.ForeColor = Color.Green;


                                            
                                           // mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);

                                            /*    string [] BleESLdata=dataGridView1[1, rowIndex].Value.ToString().Split(',');
                                                for (int i=0;i< BleESLdata.Length;i++) {
                                                    foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                                                    {
                                                    if (dr4.Cells[1].Value!=null&&BleESLdata[i] == dr4.Cells[1].Value.ToString())
                                                        {
                                                            if (dr4.Cells[8].Value != null && dr4.Cells[8].Value.ToString() != "") {
                                                                if(dr.Cells[3].Style.BackColor!= Color.Red)
                                                                dr.Cells[3].Style.BackColor = Color.Green;
                                                            }
                                                        }
                                                    }
                                                }*/

                                        }
                                            else
                                            {

                                           //     dataGridView1[2, rowIndex].Style.BackColor = Color.Green;
                                         //       dataGridView1[2, rowIndex].Value = "未綁定";
                                                dataGridView1[1, rowIndex].Value = dataTemp;
                                               if (dr.Cells[1].Style.ForeColor != Color.Orange)
                                                       dr.Cells[1].Style.ForeColor = Color.Green;


                                       //     mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);

                                        }

                                        Console.WriteLine("radioButton1" + radioButton1.Checked);
                                        Console.WriteLine("radioButton2" + radioButton2.Checked);
                                        if (radioButton1.Checked)
                                        {

                                            Console.WriteLine("111111111"+ dataTemp);
                                            immediateESLUpdate(dataTemp);
                                        }

                                        bpage.NewMateESL = dataTemp;
                                        bpage.isBack = false;
                                        backESLList.Add(bpage);
                                        MacAddressList.Add(dataTemp);
                                            break;
                                        }

                                        rowIndex++;

                                    }

                                }
                                else
                                {
                                BackPage bpage = new BackPage();
                                //Console.WriteLine("xxxxxxxxxxxx");
                                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                    {

                                        if (dr.Cells[1].Value.ToString().Contains(',' + dataTemp))
                                        {
                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + dataTemp);
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                            bpage.OldMateProduct = dr.Cells[5].Value.ToString();
                                      //  mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                        }
                                        if (dr.Cells[1].Value.ToString().Contains(dataTemp + ','))
                                        {

                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(dataTemp + ',');
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                            bpage.OldMateProduct = dr.Cells[5].Value.ToString();
                                    //    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                        }

                                        if (dr.Cells[1].Value.ToString().Contains(dataTemp))
                                        {

                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(dataTemp);
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                                            bpage.OldMateProduct = dr.Cells[5].Value.ToString();
                                     //   mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                        }

                                    }
                                 //   string aaa = MacAddressList[MacAddressList.Count - 1];
                                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                    {
                                        if (dr.Cells[5].Value.ToString() == decoded)
                                        {
                                            if (dr.Cells[1].Value.ToString().Length > 1)
                                            {
                                                dataGridView1[1, rowIndex].Value = dataGridView1[1, rowIndex].Value + "," + dataTemp;
                                                dataGridView1[1, rowIndex].Style.ForeColor = Color.Orange;
                                      //      mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            //dataGridView1[2, rowIndex].Value = "已綁定";
                                        }
                                            else
                                            {
                                                dataGridView1[1, rowIndex].Value = dataTemp;
                                                dataGridView1[1, rowIndex].Style.ForeColor = Color.Orange;
                                      //      mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            //dataGridView1[2, rowIndex].Value = "已綁定";
                                        }
                                            break;
                                        }
                                        rowIndex++;

                                    }

                                Console.WriteLine("radioButton1" + radioButton1.Checked);
                                Console.WriteLine("radioButton2" + radioButton2.Checked);
                                if (radioButton1.Checked)
                                {
                                    Console.WriteLine("222222222" + dataTemp);
                                    immediateESLUpdate(dataTemp);
                                }
                                bpage.NewMateESL = dataTemp;
                                bpage.isBack = false;
                                backESLList.Add(bpage);

                                MacAddressList.Add(decoded);
                                }

                                scanstate.Text = "配對成功";
                             //   ConnectStatus.ForeColor = Color.Green;
                            }
                            if (canConvert == false && doubletype == false)
                            {
                           //     Console.WriteLine("ssssssssssssss");
                                Boolean maccheck = false;
                            /*  foreach (string mac in MacAddressList)
                              {
                                  if (mac.Equals(decoded))
                                  {
                                      maccheck = true;
                                  }
                              }*/
                            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                            {
                                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString().Contains(decoded))
                                {
                                  //  Console.WriteLine("BBBBBBBBBB");
                                    maccheck = true;
                                    break;
                                }
                            }
                            if (maccheck == false)
                                {
                                BackPage bpage = new BackPage();
                                string aaa = MacAddressList[MacAddressList.Count - 1];

                                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                    {
                                        if (dr.Cells[5].Value.ToString() == aaa)
                                        {
                                            if (dr.Cells[1].Value.ToString().Length > 1)
                                            {
                                                dataGridView1[1, rowIndex].Value = dataGridView1[1, rowIndex].Value + "," + decoded;
                                               // dataGridView1[2, rowIndex].Value = "未綁定";
                                                dataGridView1[1, rowIndex].Style.ForeColor = Color.Green;
                                       //     mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        }
                                            else
                                            {
                                                dataGridView1[1, rowIndex].Value = decoded;
                                              //  dataGridView1[2, rowIndex].Value = "未綁定";
                                                dataGridView1[1, rowIndex].Style.ForeColor = Color.Green;
                                       //     mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        }
                                            break;
                                        }

                                    }


                                Console.WriteLine("radioButton1" + radioButton1.Checked);
                                Console.WriteLine("radioButton2" + radioButton2.Checked);
                                if (radioButton1.Checked)
                                {
                                    Console.WriteLine("333333333" + dataTemp);
                                    immediateESLUpdate(decoded);
                                }
                                bpage.NewMateESL = decoded;
                                bpage.isBack = false;
                                backESLList.Add(bpage);
                                MacAddressList.Add(decoded);

                                }
                                else
                                {
                                 BackPage bpage = new BackPage();
                                   // Console.WriteLine("dddddddddddddddd");
                                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                    {

                                        if (dr.Cells[1].Value.ToString().Contains(',' + decoded))
                                        {
                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + decoded);
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                            bpage.OldMateProduct = dr.Cells[5].Value.ToString();
                                  //      mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                        }
                                        if (dr.Cells[1].Value.ToString().Contains(decoded + ','))
                                        {

                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(decoded + ',');
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                            bpage.OldMateProduct = dr.Cells[5].Value.ToString();
                                       // mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                        }

                                        if (dr.Cells[1].Value.ToString().Contains(decoded))
                                        {

                                            int changeaddr = dr.Cells[1].Value.ToString().IndexOf(decoded);
                                            dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                                            bpage.OldMateProduct = dr.Cells[5].Value.ToString();
                                   //     mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        break;
                                        }
                                 /*       if (dr.Cells[1].Value.ToString().Length > 13) {
                                            string [] Blesadrss = dr.Cells[1].Value.ToString().Split(',');
                                            for (int i = 0; i < Blesadrss.Length; i++) {


                                            }
                                        }*/

                                    }
                                        string aaa = MacAddressList[MacAddressList.Count - 1];
                                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                    {
                                        if (dr.Cells[5].Value.ToString() == aaa)
                                        {
                                            if (dr.Cells[1].Value.ToString().Length > 1)
                                            {
                                                dataGridView1[1, rowIndex].Value = dataGridView1[1, rowIndex].Value + "," + decoded;
                                              //  dataGridView1[2, rowIndex].Value = "已綁定";
                                                dataGridView1[1, rowIndex].Style.ForeColor = Color.Orange;
                                  //          mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        }
                                            else
                                            {
                                                dataGridView1[1, rowIndex].Value = decoded;
                                               // dataGridView1[2, rowIndex].Value = "已綁定";
                                                dataGridView1[1, rowIndex].Style.ForeColor = Color.Orange;
                                      //      mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        }
                                            break;
                                        }

                                    }


                                Console.WriteLine("radioButton1" + radioButton1.Checked);
                                Console.WriteLine("radioButton2" + radioButton2.Checked);
                                if (radioButton1.Checked)
                                {
                                    Console.WriteLine("44444444" + dataTemp);
                                    immediateESLUpdate(decoded);
                                }
                                bpage.NewMateESL = decoded;
                                    bpage.isBack = false;
                                    backESLList.Add(bpage);
                                    MacAddressList.Add(decoded);

                                }
                                scanstate.Text = "配對成功";
                              //  ConnectStatus.ForeColor = Color.Green;
                            }

                            if (canConvert == false && doubletype == true && onetwo == true)
                            {
                                scanstate.Text = "配對失敗，請重新配對";
                              //  ConnectStatus.ForeColor = Color.Red;
                                // Boolean canConvert = decimal.TryParse(decoded, out number3);
                            }
                            else if (canConvert == true && doubletype == false && onetwo == true)
                            {
                                scanstate.Text = "配對失敗，請重新配對";
                             //   ConnectStatus.ForeColor = Color.Red;
                            }


                          //  Console.WriteLine("onetwo false");
                            onetwo = false;
                        }
                        scancode.Text = "";

                    }
                    else
                    {
                        scanstate.Text = "請掃描正確格式";
                    }
                }
                else
                {
                    scancode.Text = "";
                    return;
                }
                scancodeas = decoded;
           // }
              //  ));
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            Form2 f = new Form2();//產生Form2的物件，才可以使用它所提供的Method
            f.ShowDialog(this);//設定Form2為Form1的上層，並開啟Form2視窗。由於在Form1的程式碼內使用this，所以this為Form1的物件本身
        }
        string TimeS;
        string TimeE;
        string thisESLstate;
        string Container;
        string ContainVss;
        string ESLstyle;
        string APip;
        private void aaa(int KK, bool QQ, int AA) {

            if (KK > 1)
            {
              //  Console.WriteLine("++----------------------" + datagridview1curr+KK);
                if (!testest)
                {

                    Container = null;
                    ContainVss = null;
                    ESLstyle = null;
                    APip = null;
                  //  Console.WriteLine("please");
                    dataGridView3.Columns.Clear();
                    leftmosueESL.Clear();
                    dataGridView3.ColumnCount = 10;
                    DataTable bd = new DataTable();
                    // dataGridView1.MouseDown += new MouseEventHandler(dataGridView1_MouseDown);
                    thisESLstate = "待機中";
                    dataGridView3.Columns[0].Name = "ESLID";
                    dataGridView3.Columns[1].Name = "RSSI";
                    dataGridView3.Columns[2].Name = "尺寸";
                    dataGridView3.Columns[3].Name = "AP";
                    dataGridView3.Columns[4].Name = "動作";
                    dataGridView3.Columns[5].Name = "變動時間S";
                    dataGridView3.Columns[6].Name = "變動時間E";
                    dataGridView3.Columns[7].Name = "貨架";
                    dataGridView3.Columns[8].Name = "電壓";
                    dataGridView3.Columns[9].Name = "套用格式";


                    this.dataGridView3.RowsDefaultCellStyle.BackColor = Color.Bisque;
                    this.dataGridView3.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                    this.dataGridView3.AllowUserToAddRows = false;
                    this.dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    //  img1 = ImageDecoder.DecodeFromFile(openFileDialog1.FileName);
                    //MessageBox.Show(openFileDialog1.FileName );

                    dataGridView3.CellEndEdit += new DataGridViewCellEventHandler(dataGridView3_CellEndEdit);


                    string BindESL;

                    // DataGridView dgv = sender as DataGridView;

                    if (!QQ&& this.dataGridView1.CurrentCellAddress.Y != -1)
                    {

                        
                        int currentRow = dataGridView1.CurrentCell.RowIndex;
                        BindESL = dataGridView1.Rows[currentRow].Cells[1].Value.ToString();
                        if (BindESL.Length > 13)
                        {
                            Console.WriteLine("BindESL"+ BindESL);
                            string [] ALLBindESL = BindESL.Split(',');
                         //   Console.WriteLine("ALLBindESL" + ALLBindESL[0]+ ALLBindESL[1]);
                            for (int i = 0; i < ALLBindESL.Length; i++)
                                {
                                Console.WriteLine("BB" + dataGridView4.RowCount);
                                foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                                {
                                   // Console.WriteLine("ALLBindESL" + i+ ALLBindESL[i]+"ss"+ dr.Cells[1].Value.ToString());
                                    if (dr.Cells[1].Value != null && ALLBindESL[i] == dr.Cells[1].Value.ToString())
                                    {
                                        Console.WriteLine("ALLBindESL[i] "+ ALLBindESL[i]+ " dr.Cells[7].Value.ToString()" + dr.Cells[7].Value.ToString()+ "dr.Cells[5].Value.ToString()"+ dr.Cells[5].Value.ToString());
                                        if (Container==null)
                                        {
                                            if (dr.Cells[7].Value.ToString() == "")
                                                Container = "未設置";
                                            if (dr.Cells[7].Value.ToString() != "")
                                                Container = dr.Cells[7].Value.ToString();
                                        }
                                        else
                                        {

                                            if (dr.Cells[7].Value.ToString() == "")
                                                Container = Container + ",未設置";
                                            if (dr.Cells[7].Value.ToString() != "")
                                                Container = Container + "," + dr.Cells[7].Value.ToString();
                                          
                                        }
                                        if (ContainVss == null)
                                        {
                                            if (dr.Cells[5].Value.ToString() == "")
                                                ContainVss = "未偵測";
                                            if (dr.Cells[5].Value.ToString() != "")
                                                ContainVss = dr.Cells[5].Value.ToString();
                                        }
                                        else
                                        {
                                            if (dr.Cells[5].Value.ToString() == "")
                                                ContainVss = ContainVss + ",未偵測";
                                            if (dr.Cells[5].Value.ToString() != "")
                                                ContainVss = ContainVss + "," + dr.Cells[5].Value.ToString();
                                        }
                                        if (ESLstyle == null)
                                        {
                                            if (dr.Cells[3].Value.ToString() == "")
                                                ESLstyle = "無";
                                            if (dr.Cells[3].Value.ToString() != "")
                                                ESLstyle = dr.Cells[3].Value.ToString();
                                        }
                                        else
                                        {

                                            if (dr.Cells[3].Value.ToString() == "")
                                                ESLstyle = ESLstyle + ",無";
                                            if (dr.Cells[3].Value.ToString() != "")
                                                ESLstyle = ESLstyle + "," + dr.Cells[3].Value.ToString();
                                        }

                                        if (APip == null)
                                        {
                                            if (dr.Cells[8].Value.ToString() == "")
                                                APip = "無";
                                            if (dr.Cells[8].Value.ToString() != "")
                                                APip = dr.Cells[8].Value.ToString();
                                        }
                                        else
                                        {

                                            if (dr.Cells[8].Value.ToString() == "")
                                                APip = APip + ",無";
                                            if (dr.Cells[8].Value.ToString() != "")
                                                APip = APip + "," + dr.Cells[8].Value.ToString();
                                        }
                                        

                                    }

                                }
                                

                            }
                        }
                        else
                        {
                            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                            {
                                if (dr.Cells[1].Value != null && BindESL == dr.Cells[1].Value.ToString())
                                {
                                    if (dr.Cells[7].Value.ToString() == "")
                                        Container = "未設置";
                                    if (dr.Cells[5].Value.ToString() == "")
                                        ContainVss = "未偵測";
                                    if (dr.Cells[7].Value.ToString() != "")
                                        Container = dr.Cells[7].Value.ToString();
                                    if (dr.Cells[5].Value.ToString() != "")
                                        ContainVss = dr.Cells[5].Value.ToString();
                                    if (dr.Cells[3].Value.ToString() == "")
                                        ESLstyle = "無";
                                    if (dr.Cells[3].Value.ToString() != "")
                                        ESLstyle = dr.Cells[3].Value.ToString();
                                    if (dr.Cells[8].Value.ToString() == "")
                                        APip = "無";
                                    if (dr.Cells[8].Value.ToString() != "")
                                        APip = dr.Cells[8].Value.ToString();

                                }
                                    
                            }
                        }
                    //    Console.WriteLine("currentRow" + currentRow);
                        TimeS = dataGridView1.Rows[currentRow].Cells[17].Value.ToString();
                        TimeE = dataGridView1.Rows[currentRow].Cells[18].Value.ToString();
                    }
                    else
                    {
                     //   Console.WriteLine("INSAAIOUT");
                        BindESL = dataGridView1.Rows[AA].Cells[1].Value.ToString();
                        string[] ALLBindESL = BindESL.Split(',');
                        if (BindESL.Length > 13)
                        {
                            for (int i = 0; i < ALLBindESL.Length; i++)
                            {
                                //Console.WriteLine("BB" + dr.Cells[1].Value.ToString());
                                foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                                    {
                                    if (dr.Cells[1].Value != null && ALLBindESL[i] == dr.Cells[1].Value.ToString())
                                    {
                                        if (Container == null)
                                        {
                                            if (dr.Cells[7].Value.ToString() == "")
                                                Container = "未設置";
                                            if (dr.Cells[7].Value.ToString() != "")
                                                Container = dr.Cells[7].Value.ToString();      
                                        }
                                        else
                                        {

                                            if (dr.Cells[7].Value.ToString() == "")
                                                Container = Container + ",未設置";
                                            if (dr.Cells[7].Value.ToString() != "")
                                                Container = Container + "," + dr.Cells[7].Value.ToString();
                                            
                                        }

                                        if (ContainVss == null)
                                        {
                                            if (dr.Cells[5].Value.ToString() == "")
                                                ContainVss = "未偵測";
                                            if (dr.Cells[5].Value.ToString() != "")
                                                ContainVss = dr.Cells[5].Value.ToString();
                                        }
                                        else
                                        {
                                            if (dr.Cells[5].Value.ToString() == "")
                                                ContainVss = ContainVss + ",未偵測";
                                            if (dr.Cells[5].Value.ToString() != "")
                                                ContainVss = ContainVss + "," + dr.Cells[5].Value.ToString();
                                        }
                                        
                                        if (ESLstyle == null)
                                        {
                                            if (dr.Cells[3].Value.ToString() == "")
                                                ESLstyle = "無";
                                            if (dr.Cells[3].Value.ToString() != "")
                                                ESLstyle = dr.Cells[3].Value.ToString();
                                        }
                                        else
                                        {

                                            if (dr.Cells[3].Value.ToString() == "")
                                                ESLstyle = ESLstyle + ",無";
                                            if (dr.Cells[3].Value.ToString() != "")
                                                ESLstyle = ESLstyle + "," + dr.Cells[3].Value.ToString();
                                        }

                                        if (APip == null)
                                        {
                                            if (dr.Cells[8].Value.ToString() == "")
                                                APip = "無";
                                            if (dr.Cells[8].Value.ToString() != "")
                                                APip = dr.Cells[8].Value.ToString();
                                        }
                                        else
                                        {

                                            if (dr.Cells[8].Value.ToString() == "")
                                                APip = APip + ",無";
                                            if (dr.Cells[8].Value.ToString() != "")
                                                APip = APip + "," + dr.Cells[8].Value.ToString();
                                        }

                                    }

                                }
                            }
                        }
                        else
                        {
                            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                            {
                                if (dr.Cells[1].Value != null && BindESL == dr.Cells[1].Value.ToString())
                                {
                                    if (dr.Cells[7].Value.ToString() == "")
                                        Container = "未設置";
                                    if (dr.Cells[5].Value.ToString() == "")
                                        ContainVss = "未偵測";
                                    if (dr.Cells[7].Value.ToString() != "")
                                        Container = dr.Cells[7].Value.ToString();
                                    if (dr.Cells[5].Value.ToString() != "")
                                        ContainVss = dr.Cells[5].Value.ToString();
                                    if (dr.Cells[3].Value.ToString() == "")
                                        ESLstyle = "無";
                                    if (dr.Cells[3].Value.ToString() != "")
                                        ESLstyle = dr.Cells[3].Value.ToString();
                                    if (dr.Cells[8].Value.ToString() == "")
                                        APip = "無";
                                    if (dr.Cells[8].Value.ToString() != "")
                                        APip = dr.Cells[8].Value.ToString();
                                }
                            }
                        }

                     //   Console.WriteLine("currentRow" + AA);
                        TimeS = dataGridView1.Rows[AA].Cells[17].Value.ToString();
                        TimeE = dataGridView1.Rows[AA].Cells[18].Value.ToString();
                    }
                    if (BindESL != "") {
                        Console.WriteLine("BindESL"+ BindESL);
                        if (BindESL.Length > 13)
                    {
                        string[] str = BindESL.Split(',');
                        string[] Containerstr = Container.Split(',');
                        string[] ContainVssstr = ContainVss.Split(',');
                        string[] ESLstylestr = ESLstyle.Split(',');
                        string[] APipstr = APip.Split(',');
                        totalRows = str.Length;
                        for (int i = 0; i < str.Length; i++)
                        {

                            //leftmosue[i] = 1;
                            dataGridView3.Rows.Add(str[i], "", "", APipstr[i], thisESLstate, TimeS, TimeE, Containerstr[i], ContainVssstr[i], ESLstylestr[i]);
                            leftmosueESL.Add(str[i]);
                        }

                    }
                    else
                    {
                       // Console.WriteLine("BindESL" + BindESL);
                        if(BindESL!="")
                        dataGridView3.Rows.Add(BindESL, "", "", APip, thisESLstate, TimeS, TimeE, Container,ContainVss, ESLstyle);
                        //leftmosue[0] = 1;
                        leftmosueESL.Add(BindESL);
                    }

                    }

                    dataviewcurrentbefore = AA;
                  //  Console.WriteLine("+++++++++++++--");
                }
            }
            else {
                if (KK == 1) {
                   /* foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {
                            kvp.Value.mSmcEsl.startScanBleDevice(255);
                       // Console.WriteLine("BBA");
                    }*/
                  //  mSmcEsl.startScanBleDevice(255);
                  //  Console.WriteLine("KK");
                }

            //    Console.WriteLine("++----------------------" + datagridview1curr);
                datagridview1curr++;
                

                //  mSmcEsl.startScanBleDevice(255);
                //mSmcEsl.stopScanBleDevice();
            }

        }

        private void dataGridView1_CurrentCellChanged(object sender, EventArgs e)
        {


            aaa(datagridview1curr, false, 0);

        }

        private void dataGridView3_CellEndEdit(object sender, EventArgs e)
        {
            int currentRow = dataGridView1.CurrentCell.RowIndex;
            dataGridView1.Rows[currentRow].Cells[14].Value = "";
            foreach (DataGridViewRow dr in this.dataGridView3.Rows)
            {

                if (dataGridView1.Rows[currentRow].Cells[14].Value.ToString() != "")
                {
                    dataGridView1.Rows[currentRow].Cells[14].Value = dataGridView1.Rows[currentRow].Cells[14].Value.ToString() + "," + dr.Cells[7].Value;
                }
                else
                {
                    dataGridView1.Rows[currentRow].Cells[14].Value = dr.Cells[7].Value;
                }

                foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                {
                    if (dr4.Cells[1].Value!=null&&dr4.Cells[1].Value.ToString() == dr.Cells[0].Value.ToString())
                    {
                        dr4.Cells[7].Value = dr.Cells[7].Value;
                    }
                }
            }



        }

        public void ESLUpdateFormat_Click(object sender, EventArgs e)
        {

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }
            bool data2check = false;
            bool data7check = false;
            string oldEslStyle="";
            string oldEslSaleStyle="";
            foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
            {
                if (dr2.Cells[2].Value!=null&& dr2.Cells[2].Value.ToString()=="V") {
                    oldEslStyle = dr2.Cells[1].Value.ToString();
                }
               
                if (dr2.Cells[0].Value != null && (bool)dr2.Cells[0].Value)
                {
                    data2check = true;
                    //  mExcelData.dataGridViewRowCellUpdate(dataGridView2,2, dr2.Index, false, openExcelAddress, excel, excelwb, mySheet);
                }
            }

            foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
            {

                if (dr7.Cells[2].Value != null && dr7.Cells[2].Value.ToString() == "V")
                {
                    oldEslSaleStyle = dr7.Cells[1].Value.ToString();
                }

                if (dr7.Cells[0].Value != null && (bool)dr7.Cells[0].Value)
                {
                    data7check = true;
                    //  mExcelData.dataGridViewRowCellUpdate(dataGridView2,2, dr2.Index, false, openExcelAddress, excel, excelwb, mySheet);
                }
            }


            if (data2check) {
           
            String selectFormatType = "";
                foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
                {
                    if (dr2.Cells[0].Value != null && (bool)dr2.Cells[0].Value)
                    {
                        selectFormatType = dr2.Cells[4].Value.ToString();
                        if(selectFormatType=="1")
                            ESL29Format.Clear();
                        else if(selectFormatType=="2")
                            ESL42Format.Clear();
                        else
                            ESLFormat.Clear();
                    }
                }

             foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
            {
                if (dr2.Cells[0].Value != null && (bool)dr2.Cells[0].Value)
                {
                       
                        dr2.Cells[2].Value = "V";

                         mExcelData.EslStyleCgange(dataGridView2, "V", dr2.Cells[1].Value.ToString(), false, openExcelAddress, excel, excelwb, mySheet);

                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                        {
                            if (dr.Cells[15].Value != null && dr.Cells[15].Value.ToString() == "X")
                            {
                                if(dr.Cells[13].Value!=null&& dr2.Cells[1].Value != null && dr.Cells[13].Value.ToString()!= dr2.Cells[1].Value.ToString())
                                {
                                    ESLStyleDataChange = true;
                                    button19.Enabled = true;
                                    button19.BackColor = Color.FromArgb(255, 255, 192);
                                    break;
                                }
                            }
                        }


                      /*      if (dr2.Cells[1].Value.ToString()!= oldEslStyle)
                        {
                            ESLStyleDataChange = true;
                            button19.Enabled = true;
                            button19.BackColor = Color.FromArgb(255, 255, 192);
                        }*/
                }
                else {

                            if (dr2.Cells[4].Value.ToString() == selectFormatType) { 
                    dr2.Cells[2].Value = DBNull.Value;
                    mExcelData.EslStyleCgange(dataGridView2, "", dr2.Cells[1].Value.ToString(), false, openExcelAddress, excel, excelwb, mySheet);
                        }
                    }
                   
            }
            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString()== "V")
                {

                    for (int i = 0; i < dr.Cells.Count; i++) {
                      
                       //     Console.WriteLine("HGEEGE");
                            if (i== 1) {
                                styleName = dr.Cells[1].Value.ToString();
                            }
                            if (i != 0 && i != 1&&i!=2)
                            {
                                /* if (dr.Cells[i].Value.ToString() != "")
                                 {

                                     ESLFormat.Add(dr.Cells[i].Value.ToString());
                                 }
                                 else
                                 {
                                         if (i < dataGridView2.ColumnCount)
                                             if (dr.Cells[i - 1].Value.ToString() != "")
                                                 ESLFormat.Add(dr.Cells[i].Value.ToString());
                                 }*/

                                if (dr.Cells[i].Value != null && dr.Cells[i].Value.ToString() != "")
                                {
                                    if (dr.Cells[4].Value.ToString() == "1")
                                        ESL29Format.Add(dr.Cells[i].Value.ToString());
                                    else if (dr.Cells[4].Value.ToString() == "2")
                                        ESL42Format.Add(dr.Cells[i].Value.ToString());
                                    else
                                        ESLFormat.Add(dr.Cells[i].Value.ToString());
                                    //   Console.WriteLine("ESLSaleFormat" + dr7.Cells[i].Value.ToString());
                                }
                                else
                                {
                                    if (i < dataGridView2.ColumnCount)
                                        if (dr.Cells[i - 1].Value.ToString() != "" && dr.Cells[4].Value.ToString() == "0")
                                            ESLFormat.Add(dr.Cells[i].Value.ToString());
                                        else if (dr.Cells[i - 1].Value.ToString() != "" && dr.Cells[4].Value.ToString() == "1")
                                            ESL29Format.Add(dr.Cells[i].Value.ToString());
                                        else if (dr.Cells[i - 1].Value.ToString() != "" && dr.Cells[4].Value.ToString() == "2")
                                            ESL42Format.Add(dr.Cells[i].Value.ToString());
                                }
                                //       Console.WriteLine(dr.Cells[i].Value.ToString());
                            }
                    }

                }
            }

            }

            if (data7check) { 
            ESLSaleFormat.Clear();
                String selectFormatType = "";
                foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
                {
                    if (dr7.Cells[0].Value != null && (bool)dr7.Cells[0].Value)
                    {
                        selectFormatType = dr7.Cells[4].Value.ToString();
                        if (selectFormatType == "1")
                            ESLSale29Format.Clear();
                        else if (selectFormatType == "2")
                            ESLSale42Format.Clear();
                        else
                            ESLSaleFormat.Clear();
                    }
                }
            foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
            {
                if (dr7.Cells[0].Value != null && (bool)dr7.Cells[0].Value)
                {


                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                        {
                            if (dr.Cells[15].Value != null && dr.Cells[15].Value.ToString() == "V")
                            {
                                if (dr.Cells[13].Value != null && dr7.Cells[1].Value != null && dr.Cells[13].Value.ToString() != dr7.Cells[1].Value.ToString())
                                {
                                    ESLSaleStyleDataChange = true;
                                    button19.Enabled = true;
                                    button19.BackColor = Color.FromArgb(255, 255, 192);
                                    break;
                                }
                            }
                        }

                     /*   if (dr7.Cells[1].Value.ToString() != oldEslSaleStyle)
                        {
                            ESLSaleStyleDataChange = true;
                            button19.Enabled = true;
                            button19.BackColor = Color.FromArgb(255, 255, 192);
                        }*/
                        dr7.Cells[2].Value = "V";
                        mExcelData.EslStyleCgange(dataGridView7, "V", dr7.Cells[1].Value.ToString(), false, openExcelAddress, excel, excelwb, mySheet);
                }
                else
                {
                    if (dr7.Cells[4].Value.ToString() ==selectFormatType)
                        {
                            dr7.Cells[2].Value = DBNull.Value;
                            mExcelData.EslStyleCgange(dataGridView7, "", dr7.Cells[1].Value.ToString(), false, openExcelAddress, excel, excelwb, mySheet);
                        }
                        
                }
            }
            foreach (DataGridViewRow dr77 in this.dataGridView7.Rows)
            {
                if (dr77.Cells[2].Value != null && dr77.Cells[2].Value.ToString() == "V")
                {

                    for (int i = 0; i < dr77.Cells.Count; i++)
                    {
                       
                            //           Console.WriteLine("HGEEGE");
                            if (i == 1)
                            {
                                styleSaleName = dr77.Cells[1].Value.ToString();
                            }
                            if (i != 0 && i != 1 && i != 2)
                            {
                                /*  if (dr77.Cells[i].Value.ToString() != "")
                                  {
                                      ESLSaleFormat.Add(dr77.Cells[i].Value.ToString());

                                      //            Console.WriteLine(dr.Cells[i].Value.ToString());
                                  }
                                  else
                                  {
                                          if (i < dataGridView7.ColumnCount)
                                              if (dr77.Cells[i - 1].Value.ToString() != "")
                                              ESLSaleFormat.Add(dr77.Cells[i].Value.ToString());
                                  }*/

                                if (dr77.Cells[i].Value != null && dr77.Cells[i].Value.ToString() != "")
                                {
                                    if (dr77.Cells[4].Value.ToString() == "1")
                                        ESLSale29Format.Add(dr77.Cells[i].Value.ToString());
                                    else if (dr77.Cells[4].Value.ToString() == "2")
                                        ESLSale42Format.Add(dr77.Cells[i].Value.ToString());
                                    else
                                        ESLSaleFormat.Add(dr77.Cells[i].Value.ToString());
                                    //   Console.WriteLine("ESLSaleFormat" + dr777.Cells[i].Value.ToString());
                                }
                                else
                                {
                                    if (i < dataGridView7.ColumnCount)
                                        if (dr77.Cells[i - 1].Value.ToString() != "" && dr77.Cells[4].Value.ToString() == "0")
                                            ESLSaleFormat.Add(dr77.Cells[i].Value.ToString());
                                        else if (dr77.Cells[i - 1].Value.ToString() != "" && dr77.Cells[4].Value.ToString() == "1")
                                            ESLSale29Format.Add(dr77.Cells[i].Value.ToString());
                                        else if (dr77.Cells[i - 1].Value.ToString() != "" && dr77.Cells[4].Value.ToString() == "2")
                                            ESLSale42Format.Add(dr77.Cells[i].Value.ToString());
                                }
                            }

                    }

                }
            }
            }

        }
        //12/13--------------------------------------------
        /* private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
         {
             dataGridView2.DataSource = d;
             DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
             dgvc.Width = 60;
             dgvc.Name = "選取";
             dgvc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
             this.dataGridView2.Columns.Insert(0, dgvc);
            // dataGridView3.MouseDown += new MouseEventHandler(dataGridView3_MouseDown);


             openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
             //  mSmcEsl.DisConnectBleDevice();
             if (openFileDialog1.ShowDialog() == DialogResult.OK)
             {
                 //  img1 = ImageDecoder.DecodeFromFile(openFileDialog1.FileName);
                 //MessageBox.Show(openFileDialog1.FileName );
                 string tableName = "[工作表1$]";//在頁簽名稱後加$，再用中括號[]包起來
                 string sql = "select * from " + tableName;//SQL查詢
                 DataTable kk = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql);
                 totalRows = kk.Rows.Count;
                 dataGridView2.DataSource = kk;


                 this.dataGridView2.RowsDefaultCellStyle.BackColor = Color.Bisque;
                 this.dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                 this.dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
             }
             rownoinsertcount = dataGridView2.RowCount;
             Console.WriteLine("this.dataGridView2.ColumnCount" + this.dataGridView2.ColumnCount);

         }*/

        //自製部分


        private void LabelDemo_Click(object sender, EventArgs e)
        {///
            Label aaa = sender as Label;
            MyPropertiesGridLabel property = new MyPropertiesGridLabel();
            /*  property.Text = aaa.Text;
              property.Location = aaa.Location;
              property.Font = aaa.Font;
              property.ForeColor = aaa.ForeColor;
              property.AutoSize = aaa.AutoSize;
              property.BackColor = aaa.BackColor;
              property.Width = aaa.Width;
              property.Height = aaa.Height;*/
            property.SelectedObject(aaa);
            propertyGrid1.SelectedObject = property;
           // propertyGrid1.SelectedObject=sender;
        }

       private void TextBoxDemo_Click(object sender, EventArgs e)
        {///
            TextBox aaa = sender as TextBox;
            MyPropertiesGridTextBox property = new MyPropertiesGridTextBox();
            property.SelectedObject(aaa);
            propertyGrid1.SelectedObject = property;
            /*   property.Text = aaa.Text;
               property.Location = aaa.Location;
               property.Font = aaa.Font;
               property.ForeColor = aaa.ForeColor;
               property.AutoSize = aaa.AutoSize;
               property.BackColor = aaa.BackColor;
               property.Width = aaa.Width;
               property.Height = aaa.Height;
               propertyGrid1.SelectedObject = property;*/
            //propertyGrid1.SelectedObject.TextBoxProperty(sender) ;
        }

        private void PictureBoxDemo_Click(object sender, EventArgs e)
        {///
            PictureBox aaa = sender as PictureBox;
            MyPropertiesGridPicBox property = new MyPropertiesGridPicBox();
            property.SelectedObject(aaa);
            propertyGrid1.SelectedObject = property;
            //propertyGrid1.SelectedObject.TextBoxProperty(sender) ;
        }

        private void pictureBox2Q_SizeChanged(object sender, EventArgs e)
        {///
            pictureboxBarcode pictureBoxQ =  sender as pictureboxBarcode;
            Console.WriteLine("pictureBox2Q_SizeChanged");
            foreach (Control x in pictureBoxQ.Controls)
            {
                Console.WriteLine("pictureBox2Q");
                Console.WriteLine(x.Name);
            }
            Bitmap bqr = new Bitmap(pictureBoxQ.Height, pictureBoxQ.Width);
            BarcodeWriter qr = new BarcodeWriter();       // 建立條碼物件
            qr.Format = BarcodeFormat.QR_CODE;
            qr.Options.Width = Convert.ToInt32(pictureBoxQ.Height);
            qr.Options.Height = Convert.ToInt32(pictureBoxQ.Width);
            qr.Options.Margin = 0;
            bqr = qr.Write(pictureBoxQ.barcodedata.ToString());
            pictureBoxQ.Image = bqr;
        }

        private void pictureBox2B_SizeChanged(object sender, EventArgs e)
        {///
            pictureboxBarcode pictureBoxB = sender as pictureboxBarcode;
            Console.WriteLine("pictureBox2B_SizeChanged");


            foreach (Control x in pictureBoxB.Controls)
            {
                Console.WriteLine("pictureBox2B");
                Console.WriteLine(x.Name);
            }
            Bitmap bqr = new Bitmap(pictureBoxB.Height, pictureBoxB.Width);
            BarcodeWriter barcode_w = new BarcodeWriter();       // 建立條碼物件
            barcode_w.Format = BarcodeFormat.CODE_93;
            barcode_w.Options.Width = Convert.ToInt32(pictureBoxB.Height);
            barcode_w.Options.Height = Convert.ToInt32(pictureBoxB.Width);
            barcode_w.Options.Margin = 0;
            barcode_w.Options.PureBarcode = true;
            bqr = barcode_w.Write(pictureBoxB.barcodedata.ToString());
            pictureBoxB.Image = bqr;
        }


        private void tvHostingStorage_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag == null)
            {
                this.propertyGrid1.SelectedObject = null;
                return;
            }
            //Console.WriteLine("e.Node.Tag"+ e.Node.Tag);
            this.propertyGrid1.SelectedObject = e.Node.Tag;
        }


        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            string filename = @"C:\Users\abby\Desktop\test.xlsx";
            Excel.Application excel = new Excel.Application();
            Excel.Workbook excelwb = excel.Workbooks.Open(filename);
            //    excel.Application.Workbooks.Add(true);
            Excel.Worksheet mySheet = new Excel.Worksheet();
            mySheet = excelwb.Worksheets["工作表1"];//引用第一張工作表
            int last = mySheet.Rows.Count;
            int lastUsedRow = last + 1;
            int col = 1;
            foreach (Control x in pictureBox1.Controls)
            {

                // mySheet.Rows[lastUsedRow].Add(x.Name, x.Width, x.Height, x.Location.X, x.Location.Y, x.Font);
               // Console.WriteLine("lastUsedRow" + lastUsedRow);
                /* switch (x.Name)
                 {
                     case "ProName":
                         col = 1;
                         break;
                     case "ProBrand":
                         col = 7;
                         break;
                     case "ProFormat":
                         col = 13;
                         break;
                     case "ProPrice":
                         col = 19;
                         break;
                     case "ProPromotion":
                         col = 25;
                         break;
                     case "ProBarcode":
                         col = 31;
                         break;
                     case "ProESLID":
                         col = 37;
                         break;
                 }*/
               // Console.WriteLine("col" + col);
                if (col == 0)
                {
                    mySheet.Cells[lastUsedRow, col] = x.Name;
                    col++;
                    mySheet.Cells[lastUsedRow, col] = x.Width;
                    col++;
                    mySheet.Cells[lastUsedRow, col] = x.Height;
                    col++;
                    mySheet.Cells[lastUsedRow, col] = x.Location.X;
                    col++;
                    mySheet.Cells[lastUsedRow, col] = x.Location.Y;
                    col++;
                    mySheet.Cells[lastUsedRow, col] = x.Font;
                  //  Console.WriteLine("Name" + x.Name + "width" + x.Width + x.Height + "textBox1.Location" + x.Location + "x.font" + x.Font);
                }

                col = 0;
            }
            mySheet.Cells[lastUsedRow, 43] = "TEST自訂名稱";
            //設置禁止彈出保存和覆蓋的詢問提示框
            mySheet.Application.DisplayAlerts = true;
            mySheet.Application.AlertBeforeOverwriting = true;


            //excel.ActiveWorkbook.SaveCopyAs(filename);
            excelwb.Save();
            mySheet = null;
            excelwb.Close();
            excelwb = null;
            excel.Quit();
            excel = null;
            //excel.Visible = false;
            //excel.Quit();//離開聯結 
        }

        private void ToolStrip1_ItemClicked(Object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void Labelcreate(string labeldata, string labelname,bool type)
        {
         //   Console.WriteLine("!!!!!!!!!!!!777777777777");
            Label LabelDemo = new Label();
            LabelDemo.Height = 20;
            LabelDemo.TextAlign = ContentAlignment.MiddleLeft;
            //LabelDemo.TextAlign = ContentAlignment.MiddleCenter;
            LabelDemo.Width = 40;
            LabelDemo.AutoSize = true;
            LabelDemo.Name = labelname;
            LabelDemo.Text = labeldata;


            if (type)
            {
                LabelDemo.Tag = "Header";
                if (labelname == "售價")
                {
                    LabelDemo.Location = new Point(153, 24);
                }
                else if(labelname == "促銷價")
                {
                    LabelDemo.Location = new Point(149, 24);
                }
                else
                {
                    LabelDemo.Location = new Point(10, 10);
                }
                
            }
            else
            {
                LabelDemo.Tag = "L";


                if (labelname == "品名(最多10字)")
                {
                    LabelDemo.Location = new Point(10, 10);
                }
                else if (labelname == "品牌")
                {
                    LabelDemo.Location = new Point(8, 44);
                }
                else if (labelname == "規格")
                {
                    LabelDemo.Location = new Point(8, 28);
                }
                else if (labelname == "售價")
                {
                    LabelDemo.Location = new Point(152, 35);
                    LabelDemo.Font = new Font("Calibri", 26, FontStyle.Bold);

                }
                else if (labelname == "促銷價")
                {

                    LabelDemo.Location = new Point(152, 35);
                    LabelDemo.Font = new Font("Calibri", 26, FontStyle.Bold);
                }
                //   Console.WriteLine("------------------" + LabelDemo.Tag);
            }


            pictureBox1.Controls.Add(LabelDemo);

            
            
            LabelDemo.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
            LabelDemo.MouseMove += new System.Windows.Forms.MouseEventHandler(Label_MouseMove);
            LabelDemo.Click += new System.EventHandler(this.LabelDemo_Click);



        }

        private void TextBoxcreate(string textdata, string textname)
        {
           // Console.WriteLine("!!!!!!!!!!!!777777777777");
            TextBox LabelDemo = new TextBox();
            LabelDemo.Height = 20;
            LabelDemo.TextAlign = HorizontalAlignment.Center;
            //LabelDemo.TextAlign = ContentAlignment.MiddleCenter;
            LabelDemo.Width = 40;
            LabelDemo.ReadOnly = true;
            LabelDemo.AutoSize = true;
            LabelDemo.Name = textname;
            LabelDemo.Text = textdata;

                LabelDemo.Tag = "T";
             //   Console.WriteLine("------------------" + LabelDemo.Tag);

            pictureBox1.Controls.Add(LabelDemo);
            LabelDemo.Location = new Point(59, 1);
            LabelDemo.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
            LabelDemo.MouseMove += new System.Windows.Forms.MouseEventHandler(TextBox_MouseMove);
            LabelDemo.Click += new System.EventHandler(this.TextBoxDemo_Click);



        }

        private void Label_MouseDown(object sender, MouseEventArgs e)
        {

            //Point position = e.GetPosition(pictureBox1);
          //  Console.WriteLine("-e.Location.X" + -e.Location.X + "-e.Location.Y" + -e.Location.Y);
            mouse_offset = new Point(-e.X, -e.Y);
            Label label = sender as Label;
            PictureBox pic = sender as PictureBox;
            if (e.Button == MouseButtons.Right)//按下右鍵
            {
                if (label != null) {
                    menu.Show(label, new Point(e.X, e.Y));//顯示右鍵選單
                                                          //建立選單
                    ContextMenuStrip contextMenuStrip = new ContextMenuStrip();
                    ToolStripMenuItem tsmiRemoveAll = new ToolStripMenuItem("刪除");
                    tsmiRemoveAll.Click += (obj, arg) =>
                    {
                        // dgv.Rows.Clear();
                        label.Dispose();
                    };
                    contextMenuStrip.Items.Add(tsmiRemoveAll);

                    contextMenuStrip.Show(label, new Point(e.X, e.Y));
                }

                if (pic != null)
                {
                    menu.Show(pic, new Point(e.X, e.Y));//顯示右鍵選單
                                                        //建立選單
                    ContextMenuStrip contextMenuStrip = new ContextMenuStrip();
                    ToolStripMenuItem tsmiRemoveAll = new ToolStripMenuItem("刪除");
                    tsmiRemoveAll.Click += (obj, arg) =>
                    {
                        // dgv.Rows.Clear();
                        pic.Dispose();
                    };
                    contextMenuStrip.Items.Add(tsmiRemoveAll);

                    contextMenuStrip.Show(pic, new Point(e.X, e.Y));
                }

            }
        }


        private void Label_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {

              //  Console.WriteLine("dataGridView2" + dataGridView2.CurrentCell.RowIndex + "dataGridView7" + dataGridView7.CurrentCell.RowIndex);
                if (!dataGridView2.Rows[0].Selected&&!dataGridView7.Rows[0].Selected) { 
                Label lab = (Label)sender;
              //  Console.WriteLine(" lab.Location" + lab.Location);
              //  Console.WriteLine("mouse_offset.X" + mouse_offset.X);
             //   Console.WriteLine("mouse_offset.Y" + mouse_offset.Y);
                Point mos = new Point(mouse_offset.X, mouse_offset.Y);
              //  Console.WriteLine("Control.MousePosition" + Control.MousePosition);
                Point mousePos = pictureBox1.PointToClient(Control.MousePosition);
                mousePos.Offset(mos.X, mos.Y);
                lab.Location = mousePos;
                }
            }
        }

        private void TextBox_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
              ///  Console.WriteLine("dataGridView2" + dataGridView2.CurrentCell.RowIndex+ "dataGridView7" +  dataGridView7.CurrentCell.RowIndex);
                if (!dataGridView2.Rows[0].Selected && !dataGridView7.Rows[0].Selected)
                {
                    TextBox lab = (TextBox)sender;
             //   Console.WriteLine(" lab.Location" + lab.Location);
              //  Console.WriteLine("mouse_offset.X" + mouse_offset.X);
               // Console.WriteLine("mouse_offset.Y" + mouse_offset.Y);
                Point mos = new Point(mouse_offset.X, mouse_offset.Y);
               // Console.WriteLine("Control.MousePosition" + Control.MousePosition);
                Point mousePos = pictureBox1.PointToClient(Control.MousePosition);
                mousePos.Offset(mos.X, mos.Y);
                lab.Location = mousePos;
                }
            }
        }
        private void picture_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
               // Console.WriteLine("dataGridView2" + dataGridView2.CurrentCell.RowIndex + "dataGridView7" + dataGridView7.CurrentCell.RowIndex);
                if (!dataGridView2.Rows[0].Selected && !dataGridView7.Rows[0].Selected)
                { 
                    PictureBox pic = (PictureBox)sender;
               // Console.WriteLine("mouse_offset.X" + mouse_offset.X);
             //   Console.WriteLine("mouse_offset.Y" + mouse_offset.Y);
                Point mos = new Point(mouse_offset.X, mouse_offset.Y);
               // Console.WriteLine("Control.MousePosition" + Control.MousePosition);
                Point mousePos = pictureBox1.PointToClient(Control.MousePosition);
                mousePos.Offset(mos.X, mos.Y);
                pic.Location = mousePos;
                }
            }
        }
        //新增項目
        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            string value = "Document 1";
            if (InputBox("combobox", "設定綁定欄位", "欄位:", ref value) == DialogResult.OK)
            {
                // System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
                // messageBoxCS.AppendFormat("{0} = {1}", "ClickedItem", e.ClickedItem);
                // messageBoxCS.AppendLine();
                //MessageBox.Show(messageBoxCS.ToString(), "ItemClicked Event");
                Labelcreate(selectIndex + "值", selectIndex.ToString(),false);
            }
        }

        public DialogResult InputBox(string state, string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;


            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);

            buttonOk.SetBounds(228, 100, 75, 23);
            buttonCancel.SetBounds(309, 100, 75, 23);

            label.AutoSize = true;

            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 200);

            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;
            if (state == "combobox")
            {
                ComboBox comboBox = new ComboBox();
                foreach (DataGridViewColumn column in this.dataGridView1.Columns)
                {
                  //  Console.WriteLine("----------------------------------------");
                    comboBox.Items.Add(column.HeaderText);
                }

                comboBox.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
                comboBox.SetBounds(12, 36, 372, 20);
                comboBox.Anchor = textBox.Anchor | AnchorStyles.Right;
                form.Controls.AddRange(new Control[] { label, comboBox, buttonOk, buttonCancel });
            }
            else if (state == "test")
            {
                ComboBox comboBox = new ComboBox();
                    
                foreach (DataGridViewColumn column in this.dataGridView1.Columns)
                {
                 //   Console.WriteLine("----------------------------------------");
                    comboBox.Items.Add(column.HeaderText);
                }

                comboBox.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
                comboBox.SetBounds(12, 36, 372, 20);
                comboBox.Anchor = textBox.Anchor | AnchorStyles.Right;
                form.Controls.AddRange(new Control[] { label, comboBox, buttonOk, buttonCancel });
            }
            else if (state == "dataGridView")
            {
                DataGridView dataGridView = new DataGridView();

                dataGridView.ColumnCount = 3;

                dataGridView.Columns[0].Name = "ESL";
                dataGridView.Columns[1].Name = "錯誤時間";
                dataGridView.Columns[2].Name = "錯誤訊息";
                for (int i = 0; i < ESLUpdaateFail.Count; i++) {
                    dataGridView.Rows.Add(ESLUpdaateFail[i][0], ESLUpdaateFail[i][1], ESLUpdaateFail[i][2]);
                }


                dataGridView.SetBounds(12, 36, 372, 80);
                dataGridView.Anchor = textBox.Anchor | AnchorStyles.Right;
                form.Controls.AddRange(new Control[] { label, dataGridView, buttonOk, buttonCancel });
            }
            else if (state == "textbox")
            {
                textBox.SetBounds(12, 36, 372, 20);
                textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
                ComboBox comboBox = new ComboBox();
                comboBox.Items.Add("2.13");
                comboBox.Items.Add("2.9");
                comboBox.Items.Add("4.2");
                comboBox.SelectedIndex = 0;
                comboBox.SelectedIndexChanged += new EventHandler(ComboBoxSize_SelectedIndexChanged);
                comboBox.SetBounds(12, 70, 372, 20);
                comboBox.Anchor = textBox.Anchor | AnchorStyles.Right;
                form.Controls.AddRange(new Control[] { label, textBox, comboBox, buttonOk, buttonCancel });
                textBox.TextChanged += new EventHandler(TextBox_TextChanged);
            }

            else if (state == "DateTimePickerSale")
            {
                DateTimePicker TimeS = new DateTimePicker();
                DateTimePicker TimeE = new DateTimePicker();
                Label TitimeELabel = new Label();

                BeaconTimeS = DateTime.Now.ToString("yyyy/MM/dd HH: mm: ss");
                TimeS.Format = DateTimePickerFormat.Custom;
                TimeS.CustomFormat = "yyyy/MM/dd HH:mm:ss";
                TimeS.ShowUpDown = true;
                TimeS.SetBounds(12, 36, 372, 20);
                TimeS.Anchor = TimeS.Anchor | AnchorStyles.Left;

                TimeS.ValueChanged += new EventHandler(TimeS_ValueChanged);

                TitimeELabel.SetBounds(9, 74, 372, 13);
                TitimeELabel.Text = "結束時間:";
                TitimeELabel.AutoSize = true;

                BeaconTimeE = DateTime.Now.ToString("yyyy/MM/dd HH: mm: ss");
                TimeE.Format = DateTimePickerFormat.Custom;
                TimeE.CustomFormat = "yyyy/MM/dd HH:mm:ss";
                TimeE.ShowUpDown = true;
                TimeE.SetBounds(12, 90, 372, 20);
                TimeE.Anchor = TimeE.Anchor | AnchorStyles.Right;
                TimeE.ValueChanged += new EventHandler(TimeE_ValueChanged);
                buttonOk.SetBounds(228, 130, 75, 23);
                buttonCancel.SetBounds(309, 130, 75, 23);


                form.Controls.AddRange(new Control[] { label, TitimeELabel, TimeS, TimeE, buttonOk, buttonCancel });
            }
            else if (state == "DateTimePicker")
            {
                DateTimePicker TimeS = new DateTimePicker();
                DateTimePicker TimeE = new DateTimePicker();
                ComboBox ComboBoxsales = new ComboBox();
                ComboBox ComboBoxdays = new ComboBox();
                Label TitimeELabel = new Label();
                Label ComboBoSalesELabel = new Label();
                Label ComboBoDaysELabel = new Label();
                List<Item> comboItemsList = new List<Item>();
                Button beaconClear = new Button();

                beaconClear.Text = "Beacon清除";
                beaconClear.DialogResult = DialogResult.Cancel;
                beaconClear.SetBounds(30, 150, 75, 23);
                beaconClear.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;

                for (int i=1;i<100;i++) {
                   // Item comboitem = new Item(i.ToString(), i);
                    ComboBoxsales.Items.Add(new Item(i.ToString(), i));
                    ComboBoxdays.Items.Add(new Item(i.ToString(), i));
                    // comboItemsList.Add(comboitem);
                }
                
                BeaconTimeS = DateTime.Now.ToString("yyyy/MM/dd HH: mm: ss");
              /*  TimeS.Format = DateTimePickerFormat.Custom;
                TimeS.CustomFormat = "yyyy/MM/dd HH:mm:ss";
                TimeS.ShowUpDown = true;
                TimeS.SetBounds(12, 36, 372, 20);
                TimeS.Anchor = TimeS.Anchor | AnchorStyles.Left;
               
                TimeS.ValueChanged += new EventHandler(TimeS_ValueChanged);

                TitimeELabel.SetBounds(9, 74, 372, 13);
                TitimeELabel.Text = "結束時間:";
                TitimeELabel.AutoSize = true;

                BeaconTimeE = DateTime.Now.ToString("yyyy/MM/dd HH: mm: ss");
                TimeE.Format = DateTimePickerFormat.Custom;
                TimeE.CustomFormat = "yyyy/MM/dd HH:mm:ss";
                TimeE.ShowUpDown = true;
                TimeE.SetBounds(12, 90, 372, 20);
                TimeE.Anchor = TimeE.Anchor | AnchorStyles.Right;
                TimeE.ValueChanged += new EventHandler(TimeE_ValueChanged);*/

                ComboBoxsales.SetBounds(52, 90, 75, 23); 
                ComboBoxdays.SetBounds(172, 90, 75, 23);
                ComboBoSalesELabel.SetBounds(12, 90, 372, 13);
                ComboBoSalesELabel.Text = "折數:";
                ComboBoSalesELabel.AutoSize = true;
                ComboBoDaysELabel.SetBounds(130, 90, 372, 13);
                ComboBoDaysELabel.Text = "天數:";
                ComboBoDaysELabel.AutoSize = true;
                ComboBoxsales.SelectedIndex = 0;
                ComboBoxdays.SelectedIndex = 0;
                ComboBoxdays.SelectedIndexChanged += new EventHandler(ComboBoDays_ValueChanged);
                ComboBoxsales.SelectedIndexChanged += new EventHandler(ComboBoxsales_ValueChanged);
                beaconClear.Click += new EventHandler(beaconClear_Click);
                buttonOk.SetBounds(228, 150, 75, 23);
                buttonCancel.SetBounds(309, 150, 75, 23);

                buttonOk.SetBounds(228, 150, 75, 23);
                form.Controls.AddRange(new Control[] { beaconClear, ComboBoxsales, ComboBoxdays, ComboBoSalesELabel, ComboBoDaysELabel, buttonOk, buttonCancel });
                //form.Controls.AddRange(new Control[] { label, TitimeELabel, TimeS, TimeE, ComboBoxsales, ComboBoxdays, ComboBoSalesELabel, ComboBoDaysELabel, buttonOk, buttonCancel });
            }
   
            DialogResult dialogResult = form.ShowDialog();
            dialogtext = textBox.Text;
            return dialogResult;
        }



        private void TextBox_TextChanged(object sender, EventArgs e)
        {
            TextBox tb = sender as TextBox;
            foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
            {
                if (dr7.Cells[1].Value != null && tb.Text == dr7.Cells[1].Value.ToString())
                {
                    MessageBox.Show("名稱不能重複");
                    tb.Text = "";
                    return;
                }
                dialogtext= tb.Text;
            }

            foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
            {
                if (dr2.Cells[1].Value != null && tb.Text == dr2.Cells[1].Value.ToString())
                {
                    MessageBox.Show("名稱不能重複");
                    tb.Text = "";
                    return;
                }
                dialogtext = tb.Text;
            }
        }


        private void TimeS_ValueChanged(object sender, EventArgs e)
        {
            DateTimePicker TimeS = sender as DateTimePicker;
           // String strDate = TimeS.Value.ToString("yyyy-MM-dd HH:mm:ss");
           // Console.WriteLine(strDate);
            BeaconTimeS = TimeS.Text;
        }

        private void TimeE_ValueChanged(object sender, EventArgs e)
        {
            DateTimePicker TimeE = sender as DateTimePicker;
          //  String strDate = TimeE.Value.ToString("yyyy-MM-dd HH:mm:ss");


            BeaconTimeE = TimeE.Text;
        }

        private void ComboBoDays_ValueChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            if (cb.SelectedItem.ToString() != null)
            { 
            Double item = Convert.ToDouble(cb.SelectedItem.ToString());

            DateTime localDate = DateTime.Now;
            BeaconTimeS = DateTime.Now.ToString("yyyy/MM/dd HH: mm: ss");
            DateTime localDateadd = localDate.AddDays(item);
            BeaconTimeE = localDateadd.ToString("yyyy/MM/dd HH: mm: ss");
                if (Convert.ToInt32(cb.SelectedItem.ToString()) < 10)
                    beacondays = "0" + cb.SelectedItem.ToString();
                else
                    beacondays = cb.SelectedItem.ToString();
            }
        }

        private void ComboBoxsales_ValueChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            if (cb.SelectedItem.ToString() != null)
            {
                if (Convert.ToInt32(cb.SelectedItem.ToString()) < 10)
                    beaconsales = "0" + cb.SelectedItem.ToString();
                else
                    beaconsales = cb.SelectedItem.ToString();
            }
        }

        private void beaconClear_Click(object sender, EventArgs e)
        {
            List<string> nullbeacon = new List<string>();
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[3].Value != null && (bool)dr.Cells[3].Value)
                {
                    dr.Cells[21].Value = DBNull.Value;
                    dr.Cells[22].Value = DBNull.Value;
                    dr.Cells[23].Value = DBNull.Value;
                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 21, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 22, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 23, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                    nullbeacon.Add(dr.Cells[5].Value.ToString());
                    productState(dr);
                }
            }

            if (nullbeacon.Count != 0)
                beacon_data_set(nullbeacon, "", "", "");
        }
        


        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            string item = cb.Text;
            if (item != null)
                selectIndex = item;
        }


        private void ComboBoxSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            string item = cb.Text;
            if (item != null)
                selectSize = item;
        }

        

        private void toolStripLabel4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先載入資料表");
                dataGridView1.Enabled = true;
                return;
            }
            if (label1.Text == "預設版型"|| label1.Text == "特價版型")
            {
                MessageBox.Show("預設格式無法修改");
                return;
            }
            else {
                //  2/1
                bool ESLStyleCover=false;
                foreach (DataGridViewRow dr2 in this.dataGridView2.Rows) {
                    if (label1.Text == dr2.Cells[1].Value.ToString())
                    {
                        ESLStyleCover = true;
                        Console.WriteLine("ESLStyleCover");
                        if (dr2.Cells[2].Value.ToString() == "V")
                        {
                            ESLStyleDataChange = true;

                        }
                    }
                }
                foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
                {
                    if (label1.Text == dr7.Cells[1].Value.ToString())
                    {
                        ESLStyleCover = true;
                        if (dr7.Cells[2].Value.ToString() == "V")
                        {
                            ESLSaleStyleDataChange = true;
                        }
                    }
                }


                testest = true;
               // string add_PicBox=""+","+"";
             //   DataTable dt = dataGridView5.DataSource as DataTable;

               // dt.Rows.Add(new object[] { mAP_Information.AP_Name, mAP_Information.AP_IP, "8899" });
              //  foreach (Control x in pictureBox1.Controls)
               // {
                    // mySheet.Rows[lastUsedRow].Add(x.Name, x.Width, x.Height, x.Location.X, x.Location.Y, x.Font);
                    //   Console.WriteLine("lastUsedRow" + lastUsedRow);
                    /* switch (x.Name)
                     {
                         case "ProName":
                             col = 1;
                             break;
                         case "ProBrand":
                             col = 7;
                             break;
                         case "ProFormat":
                             col = 13;
                             break;
                         case "ProPrice":
                             col = 19;
                             break;
                         case "ProPromotion":
                             col = 25;
                             break;
                         case "ProBarcode":
                             col = 31;
                             break;
                         case "ProESLID":
                             col = 37;
                             break;
                     //}*/
                    // Console.WriteLine("col" + col + "lastUsedRow" + lastUsedRow + "x.Tag.ToString()" + x.Tag.ToString());
                 /*   aaa.Add(x.Tag.ToString());
               add_PicBox = add_PicBox+"," +x.Tag.ToString();
                    aaa.Add(x.Name);
                    add_PicBox = add_PicBox + "," + x.Name;
                    aaa.Add(x.Text);
                    add_PicBox = add_PicBox + "," + x.Text;
                    aaa.Add(x.Width);
                    add_PicBox = add_PicBox + "," + x.Width;
                    aaa.Add(x.Height);
                    add_PicBox = add_PicBox + "," + x.Height;
                  add_PicBox = add_PicBox + "," + x.Location.X;
                  add_PicBox = add_PicBox + "," + x.Location.Y;
                  add_PicBox = add_PicBox + "," + x.Font.Name;
                  //    Console.WriteLine("Name" + x.Name + "width" + x.Width + x.Height + "textBox1.Location" + x.Location + "x.font" + x.Font + " x.ForeColor" + x.ForeColor.A + "," + x.ForeColor.R + "," + x.ForeColor.G + "," + x.ForeColor.B + "x.Font.Style" + x.Font.Style + "x.BackColor" + x.BackColor.A + "," + x.BackColor.R + "," + x.BackColor.G + "," + x.BackColor.B);
                  add_PicBox = add_PicBox + "," + x.Font.Size;
                  add_PicBox = add_PicBox + "," + x.Font.Style;
                  add_PicBox = add_PicBox + "," + x.ForeColor.A;
                  add_PicBox = add_PicBox + "," + x.ForeColor.R;
                  add_PicBox = add_PicBox + "," + x.ForeColor.G;
                  add_PicBox = add_PicBox + "," + x.ForeColor.B;
                  add_PicBox = add_PicBox + "," + x.BackColor.A;
                  add_PicBox = add_PicBox + "," + x.BackColor.R;
                  add_PicBox = add_PicBox + "," + x.BackColor.G;
                  add_PicBox = add_PicBox + "," + x.BackColor.B;

              }*/
                string filename = openExcelAddress;
                if (ESLStyleCover)
                {
                    Console.WriteLine("ESLStyleSave1");
                    int size = 0;
                    if (pictureBox1.Height == 104)
                        size = 0;
                    else if (pictureBox1.Height == 128)
                        size = 1;
                    else if (pictureBox1.Height == 300)
                        size = 2;

                    Console.WriteLine("ESLStyleCover");
                    mExcelData.ESLStyleCover(label1.Text,pictureBox1,excel,excelwb,mySheet,size);
                }
                else
                {
                    if (ESLStyleSave)
                    {
                        Console.WriteLine("ESLStyleSave1");
                        int size = 0;
                        if (pictureBox1.Height == 104)
                            size = 0;
                        else if (pictureBox1.Height == 128)
                            size = 1;
                        else if (pictureBox1.Height == 300)
                            size = 2;

                        mExcelData.dataGridView2Update(dataGridView2, label1.Text, filename, pictureBox1, excel, excelwb, mySheet, 0,size);
                    }
                    if (ESLSaleStyleSave)
                    {
                        Console.WriteLine("ESLStyleSave2");
                        int size = 0;
                        if (pictureBox1.Height == 104)
                            size = 0;
                        else if (pictureBox1.Height == 128)
                            size = 1;
                        else if (pictureBox1.Height == 300)
                            size = 2;

                        mExcelData.dataGridView2Update(dataGridView2, label1.Text, filename, pictureBox1, excel, excelwb, mySheet, 1,size);
                    }
                }

                dataGridView2.Columns.Clear();
                dataGridView2.DataSource = d;
                DataGridViewColumn dgvc2 = new DataGridViewCheckBoxColumn();
                dgvc2.Width = 60;
                dgvc2.Name = "選取";
                dgvc2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                this.dataGridView2.Columns.Insert(0, dgvc2);
                string tableName2 = "[工作表2$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql2 = "select * from " + tableName2+"WHERE 版型類型=0";//SQL查詢
                DataTable kk2 = mExcelData.GetExcelDataTable(filename, sql2);
                dataGridView2.DataSource = kk2;

                dataGridView7.Columns.Clear();
                dataGridView7.DataSource = d;
                DataGridViewColumn dgvc7 = new DataGridViewCheckBoxColumn();
                dgvc7.Width = 60;
                dgvc7.Name = "選取";
                dgvc7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                this.dataGridView7.Columns.Insert(0, dgvc7);
                string tableName7 = "[工作表2$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql7 = "select * from " + tableName7 + "WHERE 版型類型=1";//SQL查詢
                DataTable kk7 = mExcelData.GetExcelDataTable(filename, sql7);
                dataGridView7.DataSource = kk7;

                for (int ee = 0; ee < this.dataGridView7.ColumnCount; ee++)
                {
                    if (ee != 0 && ee != 1 && ee != 2)
                        this.dataGridView7.Columns[ee].Visible = false;
                }

                for (int ee = 0; ee < this.dataGridView2.ColumnCount; ee++)
                {
                    if (ee != 0 && ee != 1 && ee != 2)
                        this.dataGridView2.Columns[ee].Visible = false;
                }
                excel = new Excel.Application();
                excelwb = excel.Workbooks.Open(openExcelAddress);
                // excel.Application.Workbooks.Add(true);
                mySheet = new Excel.Worksheet();
                //excel.Visible = false;
                //excel.Quit();//離開聯結 
                if (ESLSaleStyleDataChange)
                {
                    ESLSaleFormat.Clear();
                    foreach (DataGridViewRow dr77 in this.dataGridView7.Rows)
                    {
                        if (dr77.Cells[2].Value != null && dr77.Cells[2].Value.ToString() == "V")
                        {

                            for (int i = 1; i < dr77.Cells.Count; i++)
                            {
                                if (dr77.Cells[i].Value!=null&&dr77.Cells[i].Value.ToString() != "")
                                {
                                    //           Console.WriteLine("HGEEGE");
                                    if (i == 1)
                                    {
                                        styleSaleName = dr77.Cells[1].Value.ToString();
                                    }
                                    if (i != 0 && i != 1 && i != 2)
                                    {

                                        ESLSaleFormat.Add(dr77.Cells[i].Value.ToString());

                                        //            Console.WriteLine(dr.Cells[i].Value.ToString());
                                    }
                                }
                                else
                                {
                                    break;
                                }

                            }

                        }
                    }
                }
                if (ESLStyleDataChange)
                {
                    ESLFormat.Clear();
                    foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                    {

                        
                        if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V")
                        {
                            for (int i = 1; i < dr.Cells.Count; i++)
                            {
                                if (dr.Cells[i].Value!=null&&dr.Cells[i].Value.ToString() != "")
                                {
                                    if (i == 1)
                                    {
                                        styleName = dr.Cells[1].Value.ToString();
                                    }
                                    if (i != 0 && i != 1 && i != 2)
                                    {

                                        ESLFormat.Add(dr.Cells[i].Value.ToString());

                                    }
                                }
                                else
                                {
                                    break;
                                }

                            }

                        }
                    }
                }

                testest = false;
            }

        }

        private void toolStripLabel5_Click(object sender, EventArgs e)
        {
            string value = "Document 1";
            if (InputBox("textbox", "設定綁定欄位", "欄位:", ref value) == DialogResult.OK)
            {
                // System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
                // messageBoxCS.AppendFormat("{0} = {1}", "ClickedItem", e.ClickedItem);
                // messageBoxCS.AppendLine();
                //MessageBox.Show(messageBoxCS.ToString(), "ItemClicked Event");
            //    Console.WriteLine("pictureBox1.Image" + pictureBox1.Image);
                pictureBox1.Image = null;
                pictureBox1.Image = null;
                pictureBox1.Image = null;
                for (int i = 0; i < 3; i++) {
                    foreach (Control x in pictureBox1.Controls)
                    {
                        Console.WriteLine("x.Name" + x.Name);
                        x.Dispose();
                    }
                }

                label1.Text = dialogtext;
            }
        }

        private void toolStripLabel3_Click(object sender, EventArgs e)
        {
            string value = "Document 1";
            if (InputBox("combobox", "設定綁定欄位", "欄位:", ref value) == DialogResult.OK)
            {
                // System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
                // messageBoxCS.AppendFormat("{0} = {1}", "ClickedItem", e.ClickedItem);
                // messageBoxCS.AppendLine();
                //MessageBox.Show(messageBoxCS.ToString(), "ItemClicked Event");
                codecreate(selectIndex, selectIndex, true);
            }
        }

        private void codecreate(string data, string name, bool state) {

            pictureboxBarcode pictureBox2 = new pictureboxBarcode();
            pictureBox1.Controls.Add(pictureBox2);
            if (state)
            {
                // Set the size of the PictureBox control.
                pictureBox2.Size = new System.Drawing.Size(160, 20);

                //Set the SizeMode to center the image.
                pictureBox2.SizeMode = PictureBoxSizeMode.CenterImage;
                Bitmap bar = new Bitmap(120, 20);
                BarcodeWriter barcode_w = new BarcodeWriter();       // 建立條碼物件
                barcode_w.Format = BarcodeFormat.CODE_93;            // 條碼類別.
                barcode_w.Options.Width = 120;
                barcode_w.Options.Height = 20;
                barcode_w.Options.PureBarcode = true;               // 顯示條碼字串
                bar = barcode_w.Write("93213450A0BB");
                pictureBox2.Image = bar;
                pictureBox2.Name = name;
                pictureBox2.Location = new Point(9,78);
                //更改項目@#
                ///pictureBox2.Text =;
                pictureBox2.Tag = "B";
                pictureBox2.barcodedata = "93213450A0BB";
                pictureBox2.SizeChanged += new System.EventHandler(this.pictureBox2B_SizeChanged);
                //pictureBox2.AutoSize = true;


            }
            else {
                // Set the size of the PictureBox control.
                pictureBox2.Size = new System.Drawing.Size(35, 35);

                //Set the SizeMode to center the image.
                pictureBox2.SizeMode = PictureBoxSizeMode.CenterImage;
                Bitmap bqr = new Bitmap(35, 35);
                BarcodeWriter qr = new BarcodeWriter();       // 建立條碼物件
                qr.Format = BarcodeFormat.QR_CODE;
                qr.Options.Width = 35;
                qr.Options.Height = 35;
                qr.Options.Margin = 0;
                if (!data.Equals(""))
                {
                    bqr = qr.Write("http://www.smartchip.com.tw");
                }
                pictureBox2.Image = bqr;
                pictureBox2.Name = name;
                pictureBox2.Location = new Point(114 ,45);

                pictureBox2.Tag = "Q";
                pictureBox2.barcodedata = "http://www.smartchip.com.tw";
                pictureBox2.SizeChanged += new System.EventHandler(this.pictureBox2Q_SizeChanged);
                // pictureBox2.AutoSize = true;
            }


            pictureBox2.Text = data;


            //pictureBox2.Location = new Point(10, 10);
            pictureBox2.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
            pictureBox2.MouseMove += new System.Windows.Forms.MouseEventHandler(picture_MouseMove);
            pictureBox2.Click += new System.EventHandler(this.PictureBoxDemo_Click);
        }

        private void toolStripLabel6_Click(object sender, EventArgs e)
        {
            string value = "Document 1";
            if (InputBox("combobox", "設定綁定欄位", "欄位:", ref value) == DialogResult.OK)
            {
                // System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
                // messageBoxCS.AppendFormat("{0} = {1}", "ClickedItem", e.ClickedItem);
                // messageBoxCS.AppendLine();
                //MessageBox.Show(messageBoxCS.ToString(), "ItemClicked Event");
                codecreate(selectIndex, selectIndex, false);
            }
        }

        private void dataGridView2_CurrentCellChanged(object sender, EventArgs e)
        {

            ///1/20-----------------------------------------
            /* DataGridView dg2 = sender as DataGridView;
             foreach (DataGridViewRow dr4 in dg2.Rows) {

             }*/
            foreach (DataGridViewRow dr7 in dataGridView7.Rows)
            {
                dr7.Cells[1].Selected = false;
            }
                bbb(datagridview2curr,dataGridView2);
        }

        private void datagridview2DataUpdate()
        {
            string tableName2 = "[工作表2$]";//在頁簽名稱後加$，再用中括號[]包起來
            string sql2 = "select * from " + tableName2;//SQL查詢
            DataTable kk2 = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql2);
            dataGridView2.DataSource = kk2;
            DataGridViewCheckBoxColumn dgvc = new DataGridViewCheckBoxColumn();
            dgvc.Width = 50;
            dgvc.Name = "選取";
            this.dataGridView2.Columns.Insert(0, dgvc);

            this.dataGridView2.RowsDefaultCellStyle.BackColor = Color.Bisque;
            this.dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
            this.dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            for (int ee = 0; ee < this.dataGridView2.ColumnCount; ee++)
            {
                if (ee != 0)
                    this.dataGridView2.Columns[ee].Visible = false;
            }
        }
        //12/14------------轉換版型view
        private void bbb(int dataGrid2Current,DataGridView dataGridView) {
        //    Console.WriteLine("11111111");
         //   if (!testest)
          //  {
                
             //   Console.WriteLine("BBB-------------");
                for (int i = 0; i < picturelabel; i++)
                {
                    foreach (Control x in pictureBox1.Controls)
                    {
                  //      Console.WriteLine("x.Name" + x.Name);
                        x.Dispose();
                    }

                }
                picturelabel = 0;
                if (dataGrid2Current > 1&& dataGridView.CurrentCellAddress.Y != -1)
                {

                    int datacount = 0;
                    int currentRow = dataGridView.CurrentCell.RowIndex;
                    int currentColumn = dataGridView.Columns.Count;
                    foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                    {

                            dr8.Cells[0].Value = false;
                    }
                    label1.Text = dataGridView.Rows[currentRow].Cells[1].Value.ToString();
                 //   Console.WriteLine("currentColumn-------------" + currentColumn);
                    for (int g = 0; g < currentColumn; g++)
                    {
                        if (g != 0)
                        {
                            if (dataGridView.Rows[currentRow].Cells[g].Value == null)
                                break;

                            if(dataGridView.Rows[currentRow].Cells[4].Value.ToString()== "0")
                            {
                                pictureBox1.BackColor = Color.White;
                                pictureBox1.Size = new Size(212, 104);
                                pictureBox1.Location = new Point(235, 81);
                                panel1.Controls.Add(pictureBox1);
                            }
                            else if (dataGridView.Rows[currentRow].Cells[4].Value.ToString() == "1")
                            {
                                pictureBox1.BackColor = Color.White;
                                pictureBox1.Size = new Size(296, 126);
                                pictureBox1.Location = new Point(151, 59);
                                panel1.Controls.Add(pictureBox1);
                            }
                            else if (dataGridView.Rows[currentRow].Cells[4].Value.ToString() == "2")
                            {
                                pictureBox1.BackColor = Color.White;
                                pictureBox1.Size = new Size(400, 300);
                                pictureBox1.Location = new Point(88, 5);
                                panel1.Controls.Add(pictureBox1);
                            }
                        if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "L" || dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "Header")
                            {

                                Label LabelDemo = new Label();
                                LabelDemo.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                      //          Console.WriteLine("LabelDemo.Name" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                LabelDemo.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                               
                                foreach (DataGridViewRow dr8 in dataGridView8.Rows) {

                                    if (LabelDemo.Name == dr8.Cells[1].Value.ToString()) {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                g = g + 2;
                        //        Console.WriteLine("LabelDemo.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                LabelDemo.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                LabelDemo.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                LabelDemo.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                //LabelDemo.Name = labelname;
                                //  LabelDemo.Text = labeldata;
                                pictureBox1.Controls.Add(LabelDemo);
                                LabelDemo.AutoSize = true;
                                LabelDemo.TextAlign = ContentAlignment.MiddleCenter;
                                LabelDemo.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                LabelDemo.TextAlign = ContentAlignment.MiddleCenter;
                                g = g + 2;
                                LabelDemo.Font = new Font(dataGridView.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Regular"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Regular);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Bold"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Bold);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Italic"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Italic);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Underline"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Underline);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Strikeout"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Strikeout);
                            // if (Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g + 2].Value) == 3)
                            // LabelDemo.Font = new Font(dataGridView2.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g+1].Value),(FontStyle)dataGridView2.Rows[currentRow].Cells[g + 2].Value);
                                g = g + 3;
                                LabelDemo.ForeColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                g = g + 4;
                                LabelDemo.BackColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                LabelDemo.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                LabelDemo.MouseMove += new System.Windows.Forms.MouseEventHandler(Label_MouseMove);
                                LabelDemo.Click += new System.EventHandler(this.LabelDemo_Click);

                            }
                            if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "T")
                            {

                                TextBox LabelDemo = new TextBox();
                                LabelDemo.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                            //    Console.WriteLine("LabelDemo.Name" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                LabelDemo.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (LabelDemo.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                g = g + 2;
                            //    Console.WriteLine("LabelDemo.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                LabelDemo.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                LabelDemo.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                LabelDemo.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                //LabelDemo.Name = labelname;
                                //  LabelDemo.Text = labeldata;
                                pictureBox1.Controls.Add(LabelDemo);
                                LabelDemo.AutoSize = true;
                                LabelDemo.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                LabelDemo.TextAlign = HorizontalAlignment.Center;
                                g = g + 2;
                                LabelDemo.Font = new Font(dataGridView.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Regular"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Regular);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Bold"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Bold);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Italic"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Italic);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Underline"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Underline);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Strikeout"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Strikeout);
                            // if (Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g + 2].Value) == 3)
                            // LabelDemo.Font = new Font(dataGridView2.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g+1].Value),(FontStyle)dataGridView2.Rows[currentRow].Cells[g + 2].Value);
                                g = g + 3;
                                LabelDemo.ForeColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                g = g + 4;
                                LabelDemo.BackColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                LabelDemo.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                LabelDemo.MouseMove += new System.Windows.Forms.MouseEventHandler(TextBox_MouseMove);
                                LabelDemo.Click += new System.EventHandler(this.TextBoxDemo_Click);

                            }
                            else if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "B")
                            {
                                //pictureBox2.AutoSize = true;
                                Regex NumandEG = new Regex("[^A-Za-z0-9]");
                                bool type;
                                string barcodevalue;
                                pictureboxBarcode pictureBox2 = new pictureboxBarcode();
                                pictureBox2.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                              //  Console.WriteLine(" pictureBox2.Name" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                pictureBox2.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (pictureBox2.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                g = g + 2;
                             //   Console.WriteLine(" pictureBox2.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                pictureBox2.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                type= NumandEG.IsMatch(dataGridView.Rows[currentRow].Cells[g].Value.ToString());
                                barcodevalue = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                pictureBox1.Controls.Add(pictureBox2);
                             //   Console.WriteLine("barcode_w.Options.Width" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                               // Console.WriteLine("barcode_w.Options.Height" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                pictureBox2.Size = new System.Drawing.Size(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                pictureBox2.SizeMode = PictureBoxSizeMode.CenterImage;
                                Bitmap bar = new Bitmap(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                BarcodeWriter barcode_w = new BarcodeWriter();       // 建立條碼物件
                             //   if (type)
                              //  {
                                    barcode_w.Format = BarcodeFormat.CODE_93;
                               // }
                          //      else {
                           //         barcode_w.Format = BarcodeFormat.EAN_13;            // 條碼類別.
                            //    }
                                
                                barcode_w.Options.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                barcode_w.Options.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                barcode_w.Options.PureBarcode = true;               // 顯示條碼字串
                             //   Console.WriteLine("pictureBox2.Location " + dataGridView2.Rows[currentRow].Cells[g].Value.ToString() + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                pictureBox2.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                g = g + 12;
                                bar = barcode_w.Write("4253786521345");
                                pictureBox2.barcodedata = "4253786521345";
                                pictureBox2.Image = bar;
                                pictureBox2.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                pictureBox2.MouseMove += new System.Windows.Forms.MouseEventHandler(picture_MouseMove);
                                pictureBox2.Click += new System.EventHandler(this.PictureBoxDemo_Click);
                                pictureBox2.SizeChanged += new System.EventHandler(this.pictureBox2B_SizeChanged);
                            // pictureBox2.SizeChanged += new System.EventHandler(this.pictureBox2Q_SizeChanged);
                        }
                            else if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "Q")
                            {
                                pictureboxBarcode pictureBox2 = new pictureboxBarcode();
                                pictureBox2.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                pictureBox2.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (pictureBox2.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                g = g + 2;
                            //    Console.WriteLine(" pictureBox2.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                pictureBox2.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                pictureBox1.Controls.Add(pictureBox2);
                                pictureBox2.Size = new System.Drawing.Size(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                pictureBox2.SizeMode = PictureBoxSizeMode.CenterImage;
                                pictureBox2.barcodedata = "http://www.smartchip.com.tw";
                                Bitmap bqr = new Bitmap(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                BarcodeWriter qr = new BarcodeWriter();       // 建立條碼物件
                                qr.Format = BarcodeFormat.QR_CODE;
                                qr.Options.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                qr.Options.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                qr.Options.Margin = 0;
                                pictureBox2.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                g = g + 12;
                                bqr = qr.Write("http://www.smartchip.com.tw");
                            
                                pictureBox2.Image = bqr;
                                pictureBox2.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                pictureBox2.MouseMove += new System.Windows.Forms.MouseEventHandler(picture_MouseMove);
                                pictureBox2.Click += new System.EventHandler(this.PictureBoxDemo_Click);
                                pictureBox2.SizeChanged += new System.EventHandler(this.pictureBox2Q_SizeChanged);
                        }
                            picturelabel++;
                        }
                    }
                   // Console.WriteLine("currentRow" + currentRow);
                }
                else
                {
                //    Console.WriteLine("dataGrid2Current" + dataGrid2Current);
                    datagridview2curr++;
                }
         //   }
         /*   else
            {
                Console.WriteLine("AAA----------");
                testest = false;
            }*/
            
        }

        private void toolStripLabel7_Click(object sender, EventArgs e)
        {
            Labelcreate("TEXT", "TEXT",true);
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
          //  Console.WriteLine("WWWW");
            /*  if (e.KeyCode == Keys.Delete)
              {
                  Console.WriteLine("pppppppppppppp");
                  foreach (Control c in this.pictureBox1.Controls)
                  {
                      if (clicklabeldel== c.Name) {
                          this.Controls.Remove(c);
                      }
                  }

              }*/
        }

        private void button5_Click(object sender, EventArgs e)
        {
            testest = true;
            deldataview2no.Clear();
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }
            int relrowno = 0;


            List<DataGridViewRow> toDelete = new List<DataGridViewRow>();


            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                if (dr.Cells[0].Value != null && dr.Cells[0].Value.ToString() == "True")
                {
                    if (dr.Index == 0)
                    {
                        MessageBox.Show("基本預設格式無法刪除");
                        return;
                    }
                    deldataview2no.Add(dr.Cells[0].RowIndex);
                    mExcelData.dataviewdel(dataGridView2, deldataview2no, "工作表2", openExcelAddress, excel, excelwb, mySheet);
                    dataGridView2.Rows.Remove(dr);
                    //  toDelete.Add(dr);
                    break;
                }
            }

            DialogResult result = MessageBox.Show("版型是否刪除?", "刪除", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            { 
                foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                if (dr.Cells[0].Value != null && dr.Cells[0].Value.ToString() == "True")
                {
                        if (dr.Index == 0) {
                            MessageBox.Show("基本預設格式無法刪除");
                            return;
                        }
                    deldataview2no.Add(dr.Cells[0].RowIndex);
                    mExcelData.dataviewdel(dataGridView2, deldataview2no, "工作表2", openExcelAddress, excel, excelwb, mySheet);
                   dataGridView2.Rows.Remove(dr);
                        //  toDelete.Add(dr);
                        break;
                }
            }

                foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
                {
                    if (dr7.Cells[0].Value != null && dr7.Cells[0].Value.ToString() == "True")
                    {
                        if (dr7.Index == 0)
                        {
                            MessageBox.Show("基本預設格式無法刪除");
                            return;
                        }
                        deldataview2no.Add(dr7.Cells[0].RowIndex);
                        mExcelData.dataviewdel(dataGridView7, deldataview2no, "工作表2", openExcelAddress, excel, excelwb, mySheet);
                        dataGridView7.Rows.Remove(dr7);
                        //  toDelete.Add(dr);
                        break;
                    }
                }

                /*  foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                  {
                      testest = true;
                      if (dr.Cells[0].Value != null && dr.Cells[0].Value.ToString() == "True")
                      {

                          deldataview2no.Add(dr.Cells[0].RowIndex+1- relrowno);
                          relrowno++;

                          dataGridView2.Rows.Remove(dr);
                      }
                  }*/

                testest = false;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr = MessageBox.Show(this, "確定退出？", "退出視窗通知", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr != DialogResult.Yes)
            {

               
                e.Cancel = true;
            }
            else {
                if (openExcelAddress != null) {
                    excelwb.Save();
                    mySheet = null;
                    excelwb.Close();
                    excelwb = null;
                    excel.Quit();
                    excel = null; 
                }
                //  mExcelData.AllDataGridviewSave(dataGridView1, dataGridView2, dataGridView4, dataGridView5, false, openExcelAddress);
            }
    }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            /*   DataGridView dgv =  sender as  DataGridView;
               DataGridViewRow dr;
                int currentRow = dgv.CurrentCell.RowIndex;
                int currentColumn = dgv.Columns.Count;
                int col = dgv.HitTest(e.X, e.Y).ColumnIndex;
                int row = dgv.HitTest(e.X, e.Y).RowIndex;*/
            if (e.ColumnIndex == 0)
            {
                return;
            }
            mExcelData.dataGridViewRowCellUpdate(dataGridView4, e.ColumnIndex,e.RowIndex, false, openExcelAddress, excel, excelwb, mySheet);
            if (CountESLAll.Text != (dataGridView4.Rows.Count-1).ToString()) {
                CountESLAll.Text = (dataGridView4.Rows.Count - 1).ToString();
                int currentRow = dataGridView4.CurrentCell.RowIndex;
                dataGridView4.Rows[currentRow].Cells[6].Value = "未綁定";
                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, e.RowIndex, false, openExcelAddress, excel, excelwb, mySheet);
                //  Console.WriteLine("---------------------"+ dataGridView4.Rows[currentRow].Cells[1].Value);
            }
               
            
            //mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
        }

        private void button6_Click(object sender, EventArgs e)
        {

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }

            if (!testest) {
            int relrowno = 0;
                string pp="";
            List<DataGridViewRow> toDelete = new List<DataGridViewRow>();
            List<int> deldataview4no = new List<int>();
                foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                {
                    if (dr.Cells[0].Value != null && dr.Cells[0].Value.ToString() == "True")
                    {
                        
                    }
                }


                    DialogResult result = MessageBox.Show("該欄位是否刪除?", "Confirmation", MessageBoxButtons.YesNoCancel);
            if (result == DialogResult.Yes)
            {
                foreach (DataGridViewRow dr in this.dataGridView4.Rows)
            {
                if (dr.Cells[0].Value != null && dr.Cells[0].Value.ToString() == "True")
                {
                            if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() == "未綁定")
                            {

                            }
                            else
                            {
                                if (pp == "")
                                {
                                    pp = dr.Cells[1].Value.ToString();
                                }
                                else
                                {
                                    pp= pp+","+ dr.Cells[1].Value.ToString();
                                }
                            }

                           
                    deldataview4no.Add(dr.Cells[0].RowIndex + 2 - relrowno);
                    relrowno++;
                    toDelete.Add(dr);
                }
            }

                    if (pp != "")
                    {
                        MessageBox.Show(pp + "已綁定請先回歸原廠或下架");
                        return;
                    }


                    foreach (DataGridViewRow row in toDelete)
                {

               /* foreach (DataGridViewRow dr1 in this.dataGridView1.Rows)
                {
                    if (row.Cells[1].Value == dr1.Cells[2].Value) {
                        
                    }
                }*/
                   
                    dataGridView4.Rows.Remove(row);
                    CountESLAll.Text = (Convert.ToInt32(CountESLAll.Text)-1).ToString();
                }
                mExcelData.dataviewdel(dataGridView4,deldataview4no, "工作表3",openExcelAddress, excel, excelwb, mySheet);

            }
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void excelinputstart() {
           // Console.WriteLine("ExcelInput_Click");
            MacAddressList.Clear();
            BeaconList.Clear();
            PageList.Clear();

            dataGridView1.Columns.Clear();
            DataTable d = new DataTable();
            dataGridView1.DataSource = d;

            dataGridView2.Columns.Clear();
            dataGridView2.DataSource = d;

            dataGridView4.Columns.Clear();
            dataGridView4.DataSource = d;
            DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
            dgvc.Width = 60;
            dgvc.Name = "選取";
            dgvc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewColumn dgvc2 = new DataGridViewCheckBoxColumn();
            dgvc2.Width = 60;
            dgvc2.Name = "選取";
            dgvc2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewColumn dgvc3 = new DataGridViewCheckBoxColumn();
            dgvc3.Width = 60;
            dgvc3.Name = "選取";
            dgvc3.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            this.dataGridView1.Columns.Insert(0, dgvc);
            this.dataGridView2.Columns.Insert(0, dgvc2);
            this.dataGridView4.Columns.Insert(0, dgvc3);
            dataGridView1.MouseDown += new MouseEventHandler(dataGridView1_MouseDown);


            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            //  mSmcEsl.DisConnectBleDevice();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //  img1 = ImageDecoder.DecodeFromFile(openFileDialog1.FileName);
                //MessageBox.Show(openFileDialog1.FileName );
                string tableName = "[工作表1$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql = "select * from " + tableName;//SQL查詢
                DataTable kk = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql);
                string tableName2 = "[工作表2$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql2 = "select * from " + tableName2;//SQL查詢
                DataTable kk2 = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql2);
                string tableName3 = "[工作表3$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql3 = "select * from " + tableName3;//SQL查詢
                DataTable kk3 = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql3);
                string tableName4 = "[工作表4$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql4 = "select * from " + tableName4;//SQL查詢
                DataTable kk4 = mExcelData.GetExcelDataTable(openFileDialog1.FileName, sql4);
                
                
                dataGridView2.DataSource = kk2;
                dataGridView4.DataSource = kk3;
                dataGridView5.DataSource = kk4;
                dataGridView1.DataSource = kk;
                foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                {
                    if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() != "未綁定")
                        BindESL.Text = (Convert.ToInt32(BindESL.Text) + 1).ToString();
                }

               /* foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                {
                    if (dr.Cells[1].Value != null && dr.Cells[2].Value != null)
                    {
                        mSmcEsl.setUdpClient(dr.Cells[1].Value.ToString(),Convert.ToInt32(dr.Cells[2].Value));
                    }
                        
                }*/
                CountESLAll.Text = (kk3.Rows.Count-2).ToString();
                productAll.Text = kk.Rows.Count.ToString();
                dgvc = new DataGridViewCheckBoxColumn();
                dgvc.Width = 50;
                dgvc.Name = "Beacon選取";
                this.dataGridView1.Columns.Insert(3, dgvc);


                this.dataGridView4.RowsDefaultCellStyle.BackColor = Color.Bisque;
                this.dataGridView4.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView4.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

                this.dataGridView2.RowsDefaultCellStyle.BackColor = Color.Bisque;
                this.dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                //this.dataGridView2.AllowUserToAddRows = false;
                this.dataGridView1.RowsDefaultCellStyle.BackColor = Color.Bisque;
                this.dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                rownoinsertcount = dataGridView1.RowCount;
              //  Console.WriteLine("this.dataGridView1.ColumnCount" + this.dataGridView1.ColumnCount);
                this.dataGridView1.Columns[1].ReadOnly = true;
                //this.dataGridView1.Columns[2].ReadOnly = true;
                this.dataGridView1.Columns[12].ReadOnly = true;
                this.dataGridView1.Columns[13].ReadOnly = true;
                this.dataGridView1.Columns[15].ReadOnly = true;
                this.dataGridView1.Columns[16].ReadOnly = true;
                this.dataGridView1.Columns[17].ReadOnly = true;
                this.dataGridView4.Columns[2].ReadOnly = true;
                this.dataGridView4.Columns[3].ReadOnly = true;
                this.dataGridView4.Columns[4].ReadOnly = true;
                this.dataGridView4.Columns[5].ReadOnly = true;
                this.dataGridView4.Columns[6].ReadOnly = true;
                //this.dataGridView1.Columns[12].Visible = false;
                //this.dataGridView1.Columns[18].Visible = false;
                for (int i = 0; i < this.dataGridView1.ColumnCount; i++)
                {
                    if (i != 0)
                        headertextall = headertextall + ",";
                    headertextall = headertextall + this.dataGridView1.Columns[i].Name;
                }

             //   Console.WriteLine("headertextall" + headertextall);
                for (int ee = 0; ee < this.dataGridView2.ColumnCount; ee++)
                {
                    if (ee != 0 && ee != 1)
                        this.dataGridView2.Columns[ee].Visible = false;
                }
            }
        }

        private void toolStripLabel8_Click(object sender, EventArgs e)
        {
            string value = "Document 1";
            if (InputBox("combobox", "設定綁定欄位", "欄位:", ref value) == DialogResult.OK)
            {
                // System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
                // messageBoxCS.AppendFormat("{0} = {1}", "ClickedItem", e.ClickedItem);
                // messageBoxCS.AppendLine();
                //MessageBox.Show(messageBoxCS.ToString(), "ItemClicked Event");
                TextBoxcreate(selectIndex + "值", selectIndex.ToString());
            }
        }

        private void button4_Click_2(object sender, EventArgs e)
        {

            string value = "Document 1";
            if (InputBox("dataGridView", "設定綁定欄位", "欄位:", ref value) == DialogResult.OK)
            {
                
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            
            if (!autoMateESL)
            {
                foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                {
                    dr4.Cells[8].Value = "";
                }

                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                {

                    dr5.Cells[5].Value = 0;

                    if (dr5.Cells[1].Value != null && dr5.Cells[1].Value.ToString() == "未指定")
                        dr5.Cells[5].Value =dataGridView4.RowCount - 2;
                }

                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    kvp.Value.mSmcEsl.startScanBleDevice();
                    Console.WriteLine("BBA");
                }


                autoMateESL = true;
                button9.Text = "停止配對";
                button9.ForeColor = Color.Red;
            }
            else
            {
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {

                    kvp.Value.mSmcEsl.stopScanBleDevice();
                }
              //  mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                autoMateESL = false;
                button9.Text = "重置AP配對";
                button9.ForeColor = Color.Black;
            }
        }

        private void dataGridView5_CurrentCellChanged(object sender, EventArgs e)
        {
            if (this.dataGridView5.CurrentCellAddress.Y == -1)
                return;


            ccctrue = true;
      //      Console.WriteLine("rrrrrrrrrrrr");
         //   ccc(3);
        }

   /*     private void  ccc(int datagridview3count) {
           
         //   Console.WriteLine("QQQQQQ"+ datagridview3curr);
            if (datagridview3count > 2&& this.dataGridView5.CurrentCellAddress.Y != -1)
            {
             //   Console.WriteLine("fhfghfh");
                
                dataGridView6.Columns.Clear();
                leftmosueESL.Clear();
                dataGridView6.ColumnCount = 3;
                DataTable bd = new DataTable();
                DataGridViewColumn dgvc = new DataGridViewCheckBoxColumn();
                dgvc.Width = 60;
                dgvc.Name = "選取";
                dgvc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                this.dataGridView6.Columns.Insert(0, dgvc);

                // dataGridView1.MouseDown += new MouseEventHandler(dataGridView1_MouseDown);
                dataGridView6.Columns[1].Name = "ESLID";
                dataGridView6.Columns[2].Name = "RSSI";
                dataGridView6.Columns[3].Name = "配對AP";
                dataGridView6.Columns[1].ReadOnly = true;
                dataGridView6.Columns[2].ReadOnly = true;
                dataGridView6.Columns[3].ReadOnly = true;
                this.dataGridView6.RowsDefaultCellStyle.BackColor = Color.Bisque;
                this.dataGridView6.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
                this.dataGridView6.AllowUserToAddRows = false;
                this.dataGridView6.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                //  img1 = ImageDecoder.DecodeFromFile(openFileDialog1.FileName);
                //MessageBox.Show(openFileDialog1.FileName );
                int datacount = 0;
                int currentRow = dataGridView5.CurrentCell.RowIndex;
                int currentColumn = dataGridView5.Columns.Count;
                this.dataGridView6.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                {
                   // Console.WriteLine(dr.Cells[8].Value.ToString()+ dataGridView5.Rows[currentRow].Cells[1].Value.ToString());
                    if (dr.Cells[8].Value != null && dr.Cells[8].Value.ToString() == dataGridView5.Rows[currentRow].Cells[2].Value.ToString())
                    {
             //           Console.WriteLine("-+-+");
                        dataGridView6.Rows.Add(false,dr.Cells[1].Value, dr.Cells[4].Value, dr.Cells[8].Value);
                    }

                }
                ccctrue = false;

            }
            else {
          //      Console.WriteLine("WHY");
                datagridview3curr++;

            }
            }*/

        private void dataGridView6_CurrentCellChanged(object sender, EventArgs e)
        {
          //  Console.WriteLine("66666666"+ dataGridView6.RowCount);
          //  if (!ccctrue&& this.dataGridView6.CurrentCellAddress.Y!=-1) {
                
          //  int currentRow = dataGridView6.CurrentCell.RowIndex;
          //  int currentColumn = dataGridView6.Columns.Count;
          //  APESLID.Text = dataGridView6.Rows[currentRow].Cells[1].Value.ToString();
          //  APRSSI.Text = dataGridView6.Rows[currentRow].Cells[2].Value.ToString();
          //  comboBox1.SelectedItem = dataGridView6.Rows[currentRow].Cells[3].Value.ToString();
         //   }
            //   comboBox1
        }

        private void button8_Click(object sender, EventArgs e)
        {

            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {

                kvp.Value.mSmcEsl.stopScanBleDevice();
            }
            System.Threading.Thread.Sleep(100);
            listcount = 0;
            PageList.Clear();
            CheckESLOnly = true;
         /*   foreach (DataGridViewRow dr in this.dataGridView6.Rows) {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    Page1 mPageC = new Page1();
                    mPageC.BleAddress = dr.Cells[1].Value.ToString();
                    mPageC.APLink = dr.Cells[2].Value.ToString();
               //     Console.WriteLine("mPageC.BleAddress" + mPageC.BleAddress);
                    PageList.Add(mPageC);
                }
               
            }*/
            if (PageList.Count > 0) { 
            Page1 mPage1 = PageList[listcount];
            //    Console.WriteLine("PageList[listcount]"+PageList[listcount].BleAddress);
             foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    if (kvp.Key.Contains(mPage1.APLink))
                    {
                        //         Console.WriteLine("kvp.Key" + kvp.Key);
                      //  ConnectBleTimeOut.Start();
                        kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                    }
                }

                richTextBox1.Text = richTextBox1.Text + mPage1.BleAddress + "  嘗試連線中請稍候... \r\n";
            }


        }


        //ESL手動分配AP功能
     /*   private void button7_Click(object sender, EventArgs e)
        {

            List<DataGridViewRow> removeRow = new List<DataGridViewRow>();
            foreach (DataGridViewRow dr in this.dataGridView6.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {

                    Console.WriteLine("dr.Index"+ dr.Index);
                    foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                    {
                        if (dr4.Cells[1].Value != null&&dr.Cells[1].Value.ToString()==dr4.Cells[1].Value.ToString()){
                           dr4.Cells[8].Value = comboBox1.SelectedItem.ToString();
                            mExcelData.dataGridViewRowCellUpdate(dataGridView4, 8, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                        }
                    }
                    removeRow.Add(dr);

               
                }
                    

            }
            for(int i=0;i< removeRow.Count; i++)
            {

                dataGridView6.Rows.Remove(removeRow[i]);

            }
            mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);


        }
        */
        private void BindESL_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView5_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Console.WriteLine("dataGridView5_"+ e.ColumnIndex+", "+e.RowIndex);
            if(e.ColumnIndex!=0)
            mExcelData.dataGridViewRowCellUpdate(dataGridView5, e.ColumnIndex, e.RowIndex, false, openExcelAddress, excel, excelwb, mySheet);

         /*   if (comboBox1.Items.Count + 1 != dataGridView5.RowCount)
            {
                comboBox1.Items.Clear();
                foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                {
                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() != "")
                        comboBox1.Items.Add(dr.Cells[2].Value.ToString());
                }
            }*/
            //mExcelData.DataGridview5Update(dataGridView5,false,openExcelAddress);
        }

        private void UpdateESLDen_Click(object sender, EventArgs e)
        {

        }

        private void APLink_Click(object sender, EventArgs e)
        {
            APStart = true;
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }

            EslUdpTest.Tools tool = new EslUdpTest.Tools();
            //Tools tool = new Tools();
            tool.onApScanEvent += new EventHandler(AP_Scan);
            tool.SNC_GetAP_Info();
            //  Console.WriteLine("---");
            setLocalTime();
            //datagridview1curr = 2;
            datagridview1curr = 2;
        }

        private void setLocalTime() {
            int yy = int.Parse(String.Format("{0}", DateTime.Now.ToString("yy")));
            int MM = int.Parse(String.Format("{0}", DateTime.Now.ToString("MM")));
            int dd = int.Parse(String.Format("{0}", DateTime.Now.ToString("dd")));
            int HH = int.Parse(String.Format("{0}", DateTime.Now.ToString("HH")));
            int mm = int.Parse(String.Format("{0}", DateTime.Now.ToString("mm")));
            int ss = int.Parse(String.Format("{0}", DateTime.Now.ToString("ss")));
            String dateString = String.Format("{0}", DateTime.Now.ToString("MM/dd/yyyy"));
            DateTime Week = DateTime.Parse(dateString, CultureInfo.InvariantCulture);
            //mSmcEsl.setRTCTime(yy, MM, dd, (int)Week.DayOfWeek, HH, mm, ss);
            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {
                
                    kvp.Value.mSmcEsl.setRTCTime(yy, MM, dd, (int)Week.DayOfWeek, HH, mm, ss);
            }
        }

        private void setBeaconTime(string APIP) {
            int yy = int.Parse(Convert.ToDateTime(BeaconTimeS).ToString("yy"));
            int MM = int.Parse(Convert.ToDateTime(BeaconTimeS).ToString("MM"));
            int dd = int.Parse(Convert.ToDateTime(BeaconTimeS).ToString("dd"));
            int HH = int.Parse(Convert.ToDateTime(BeaconTimeS).ToString("HH"));
            int mm = int.Parse(Convert.ToDateTime(BeaconTimeS).ToString("mm"));

            //  mSmcEsl.setBeaconTime(yy, MM, dd, HH, mm, eyy, eMM, edd, eHH, emm);

                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    if (kvp.Key.Contains(APIP))
                    {
                        kvp.Value.mSmcEsl.setBeaconTime(yy, MM, dd, HH, mm, 99, 12, 31, 23, 59);
                    }
                }

        }

      /*  private void button10_Click(object sender, EventArgs e)
        {
            listcount = 0;
            checkESLV.Clear();

            UpdateESLDen.Text = "0";
            updateESLper.Text = "0";
            if (!APStart)
            {
                MessageBox.Show("請先連接AP");
               // dataGridView1.Enabled = true;
                return;
            }

            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {
                kvp.Value.mSmcEsl.stopScanBleDevice();
            }
            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {
                    UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                    Page mage =new Page();
                    mage.No = dr.Index.ToString();
                    mage.APID = dr.Cells[8].Value.ToString();
                    mage.ESLID = dr.Cells[1].Value.ToString();
                    checkESLV.Add(mage);
                }

            }

            if (checkESLV.Count != 0) {

                foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                {
                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == checkESLV[listcount].APID)
                    {
                        if (dr.Cells[4].Value.ToString() == "")
                        {
                                MessageBox.Show(checkESLV[listcount].ESLID + "該ESL綁定AP未啟用");
                                
                            return;
                        }
                    }

                }
                checkV = true;
                int numVal = Convert.ToInt32(checkESLV[listcount].No);
                dataGridView4.Rows[numVal].Cells[0].Selected = true;
                DisConnectTimer.Interval=10000;
                DisConnectTimer.Start();
                richTextBox1.Text = checkESLV[listcount].ESLID + "  嘗試連線中請稍候... \r\n";
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {
                        if (kvp.Key.Contains(checkESLV[listcount].APID))
                        {
                            macaddress = checkESLV[listcount].ESLID;
                            kvp.Value.mSmcEsl.ConnectBleDevice(checkESLV[listcount].ESLID);
                         //   System.Threading.Thread.Sleep(2000);
                          //  kvp.Value.mSmcEsl.ReadEslBattery();
                          //  System.Threading.Thread.Sleep(4000);
                           //  kvp.Value.mSmcEsl.DisConnectBleDevice();
                        }
                    }
                }


        }*/

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
            {
                dr7.Selected = false;
            }
            if (e.ColumnIndex==0) {
                foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
                {
                    dr2.Cells[0].Value = false;
                }
                foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
                {
                    dr7.Cells[0].Value = false;
                }
                dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = true;
            }
        }


        private void button11_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in this.dataGridView1.Rows) {


               
                if (dr.Cells[6].RowIndex == dataGridView1.Rows.Count - 1) {
                    break;
                }
                if (dr.Cells[6].Value!=null&&dr.Cells[6].Value.ToString().Contains(textBox2.Text))
                {
                    dr.Visible = true;
                }
                else
                {
                    dataGridView1.CurrentCell = null;
                    dr.Visible = false;
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {

            dataGridView4.CurrentCell = null;
            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
            {
            //    Console.WriteLine("dr.Index" + dr.Index);
                
                if (dr.Cells[1].RowIndex == dataGridView4.Rows.Count - 1)
                {
                    break;
                }

                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString().Contains(textBox3.Text))
                {
                    dr.Visible = true;
                }
                else
                {
                    dr.Visible = false;
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }

            if (!APStart)
            {
                MessageBox.Show("請先連接AP");
                dataGridView1.Enabled = true;
                return;
            }


            string value = "Document 1";
            if (InputBox("DateTimePickerSale", "特價時間設定", "開始時間:", ref value) == DialogResult.OK)
            {

                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    if (dr.Cells[3].Value != null && (bool)dr.Cells[3].Value)
                    {
                        dr.Cells[19].Value = BeaconTimeS;
                        dr.Cells[20].Value = BeaconTimeE;
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                    }
                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }

            string value = "";
            if (InputBox("textbox", "新增預設版型", "版型名稱:", ref value) == DialogResult.OK)
            {
                foreach (DataGridViewRow dr7 in dataGridView7.Rows)
                {
                    dr7.Selected = false;

                }
                foreach (DataGridViewRow dr2 in dataGridView2.Rows)
                {
                    dr2.Selected = false;

                }

                if (dialogtext == "")
                {
                    MessageBox.Show("不能為空值");
                    return;
                }

                DataGridView dataGridView = dataGridView2;

                pictureBox1.Image = null;
                pictureBox1.Image = null;
                pictureBox1.Image = null;
                for (int i = 0; i < picturelabel; i++)
                {
                    foreach (Control x in pictureBox1.Controls)
                    {
                        //      Console.WriteLine("x.Name" + x.Name);
                        x.Dispose();
                    }

                }
                Console.WriteLine("selectSize:" + selectSize);
                if (selectSize == "2.13")
                {
                    pictureBox1.BackColor = Color.White;
                    pictureBox1.Size = new Size(212, 104);
                    pictureBox1.Location = new Point(235, 81);
                    panel1.Controls.Add(pictureBox1);
                }
                else if (selectSize == "2.9")
                {
                    pictureBox1.BackColor = Color.White;
                    pictureBox1.Size = new Size(296, 128);
                    pictureBox1.Location = new Point(151, 59);
                    panel1.Controls.Add(pictureBox1);
                }
                else if (selectSize == "4.2")
                {
                    pictureBox1.BackColor = Color.White;
                    pictureBox1.Size = new Size(400, 300);
                    pictureBox1.Location = new Point(88, 5);
                    panel1.Controls.Add(pictureBox1);
                }

                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                {
                        dr8.Cells[0].Value = false;

                }
                label1.Text = dialogtext;
                ESLStyleSave = true;
                ESLSaleStyleSave = false;
                picturelabel = 0;
                if (2 > 1 && dataGridView.CurrentCellAddress.Y != -1)
                {

                    int datacount = 0;
                    int currentRow = 0;
                    int currentColumn = dataGridView.Columns.Count;
                    //   Console.WriteLine("currentColumn-------------" + currentColumn);
                    for (int g = 0; g < currentColumn; g++)
                    {
                        if (g != 0)
                        {
                            if (dataGridView.Rows[currentRow].Cells[g].Value == null)
                                break;
                            if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "L" || dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "Header")
                            {

                                Label LabelDemo = new Label();
                                LabelDemo.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                //          Console.WriteLine("LabelDemo.Name" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                LabelDemo.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                g = g + 2;

                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (LabelDemo.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                //        Console.WriteLine("LabelDemo.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                LabelDemo.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                LabelDemo.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                LabelDemo.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                //LabelDemo.Name = labelname;
                                //  LabelDemo.Text = labeldata;
                                pictureBox1.Controls.Add(LabelDemo);
                                LabelDemo.AutoSize = true;
                                LabelDemo.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                LabelDemo.TextAlign = ContentAlignment.MiddleCenter;
                                g = g + 2;
                                LabelDemo.Font = new Font(dataGridView.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Regular"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style| FontStyle.Regular);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Bold"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Bold);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Italic"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Italic);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Underline"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Underline);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Strikeout"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Strikeout);
                                // if (Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g + 2].Value) == 3)
                                // LabelDemo.Font = new Font(dataGridView2.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g+1].Value),(FontStyle)dataGridView2.Rows[currentRow].Cells[g + 2].Value);
                                g = g + 3;
                                LabelDemo.ForeColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                g = g + 4;
                                LabelDemo.BackColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                LabelDemo.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                LabelDemo.MouseMove += new System.Windows.Forms.MouseEventHandler(Label_MouseMove);
                                LabelDemo.Click += new System.EventHandler(this.LabelDemo_Click);

                            }
                            if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "T")
                            {

                                TextBox LabelDemo = new TextBox();
                                LabelDemo.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                //    Console.WriteLine("LabelDemo.Name" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                LabelDemo.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                g = g + 2;

                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (LabelDemo.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                //    Console.WriteLine("LabelDemo.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                LabelDemo.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                LabelDemo.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                LabelDemo.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                //LabelDemo.Name = labelname;
                                //  LabelDemo.Text = labeldata;
                                pictureBox1.Controls.Add(LabelDemo);
                                LabelDemo.AutoSize = true;
                                LabelDemo.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                LabelDemo.TextAlign = HorizontalAlignment.Center;
                                g = g + 2;
                                LabelDemo.Font = new Font(dataGridView.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Regular"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Regular);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Bold"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Bold);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Italic"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Italic);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Underline"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Underline);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Strikeout"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Strikeout);
                                // if (Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g + 2].Value) == 3)
                                // LabelDemo.Font = new Font(dataGridView2.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g+1].Value),(FontStyle)dataGridView2.Rows[currentRow].Cells[g + 2].Value);
                                g = g + 3;
                                LabelDemo.ForeColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                g = g + 4;
                                LabelDemo.BackColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                LabelDemo.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                LabelDemo.MouseMove += new System.Windows.Forms.MouseEventHandler(TextBox_MouseMove);
                                LabelDemo.Click += new System.EventHandler(this.TextBoxDemo_Click);

                            }
                            else if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "B")
                            {
                                //pictureBox2.AutoSize = true;
                                Regex NumandEG = new Regex("[^A-Za-z0-9]");
                                bool type;
                                string barcodevalue;
                                pictureboxBarcode pictureBox2 = new pictureboxBarcode();
                                pictureBox2.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                //  Console.WriteLine(" pictureBox2.Name" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                pictureBox2.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                g = g + 2;

                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (pictureBox2.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                //   Console.WriteLine(" pictureBox2.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                pictureBox2.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                type = NumandEG.IsMatch(dataGridView.Rows[currentRow].Cells[g].Value.ToString());
                                barcodevalue = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                pictureBox1.Controls.Add(pictureBox2);
                                //   Console.WriteLine("barcode_w.Options.Width" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                // Console.WriteLine("barcode_w.Options.Height" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                pictureBox2.Size = new System.Drawing.Size(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                pictureBox2.SizeMode = PictureBoxSizeMode.CenterImage;
                                Bitmap bar = new Bitmap(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                BarcodeWriter barcode_w = new BarcodeWriter();       // 建立條碼物件
                                                                                     //   if (type)
                                                                                     //  {
                                barcode_w.Format = BarcodeFormat.CODE_93;
                                // }
                                //      else {
                                //         barcode_w.Format = BarcodeFormat.EAN_13;            // 條碼類別.
                                //    }

                                barcode_w.Options.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                barcode_w.Options.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                barcode_w.Options.PureBarcode = true;               // 顯示條碼字串
                                                                                    //   Console.WriteLine("pictureBox2.Location " + dataGridView2.Rows[currentRow].Cells[g].Value.ToString() + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                pictureBox2.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                g = g + 12;
                                bar = barcode_w.Write("4253786521345");
                                pictureBox2.barcodedata = "4253786521345";
                                pictureBox2.Image = bar;
                                pictureBox2.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                pictureBox2.MouseMove += new System.Windows.Forms.MouseEventHandler(picture_MouseMove);
                                pictureBox2.Click += new System.EventHandler(this.PictureBoxDemo_Click);
                                pictureBox2.SizeChanged += new System.EventHandler(this.pictureBox2B_SizeChanged);

                            }
                            else if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "Q")
                            {
                                pictureboxBarcode pictureBox2 = new pictureboxBarcode();
                                pictureBox2.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                pictureBox2.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                g = g + 2;

                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (pictureBox2.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                //    Console.WriteLine(" pictureBox2.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                pictureBox2.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                pictureBox1.Controls.Add(pictureBox2);
                                pictureBox2.Size = new System.Drawing.Size(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                pictureBox2.SizeMode = PictureBoxSizeMode.CenterImage;
                                Bitmap bqr = new Bitmap(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                BarcodeWriter qr = new BarcodeWriter();       // 建立條碼物件
                                qr.Format = BarcodeFormat.QR_CODE;
                                qr.Options.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                qr.Options.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                qr.Options.Margin = 0;
                                pictureBox2.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                g = g + 12;
                                bqr = qr.Write("http://www.smartchip.com.tw");
                                pictureBox2.barcodedata = "http://www.smartchip.com.tw";
                                pictureBox2.Image = bqr;
                                pictureBox2.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                pictureBox2.MouseMove += new System.Windows.Forms.MouseEventHandler(picture_MouseMove);
                                pictureBox2.Click += new System.EventHandler(this.PictureBoxDemo_Click);
                                pictureBox2.SizeChanged += new System.EventHandler(this.pictureBox2Q_SizeChanged);
                            }
                            picturelabel++;
                        }
                    }
                    // Console.WriteLine("currentRow" + currentRow);
                }
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }


            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {

                kvp.Value.mSmcEsl.stopScanBleDevice();
            }
            if (!testest)
            {
                string eslAPNoSetMsg = null;
                nullMsg = null;
            removeESLingstate = true;
            UpdateESLDen.Text = "0";
            updateESLper.Text = "0";
            stopwatch.Reset();
            stopwatch.Start();
            PageList.Clear();
            //UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
            listcount = 0;
            dataGridView1.Enabled = false;
            dataGridView1.ClearSelection();
            // mSmcEsl.DisConnectBleDevice();
           // Console.WriteLine("PageList" + PageList);
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先載入資料表");
                dataGridView1.Enabled = true;
                return;
            }

            if (!APStart)
            {
                MessageBox.Show("請先連接AP");
                dataGridView1.Enabled = true;
                return;
            }

        /*    if (styleName == null)
            {
                if (dataGridView2.Rows.Count > 1)
                {
                    foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                    {
               //         Console.WriteLine(dr.Cells[1].RowIndex + dr.Cells[1].Value.ToString());

                        if (dr.Cells[1].RowIndex == 0)
                        {

                            for (int i = 0; i < dr.Cells.Count; i++)
                            {

                                if (dr.Cells[i].Value != null && dr.Cells[i].Value.ToString() != "")
                                {
                        //            Console.WriteLine("HGEEGE");
                                    if (i == 1)
                                    {
                                        styleName = dr.Cells[1].Value.ToString();
                                    }

                                    if (i != 0 && i != 1)
                                    {

                                        ESLFormat.Add(dr.Cells[i].Value.ToString());

                                  //      Console.WriteLine(dr.Cells[i].Value.ToString());
                                    }
                                }
                            }

                        }
                    }
                }
            }*/

            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
            {
                if (dr.Cells[0].Value != null && (bool)dr.Cells[0].Value)
                {

                    //   dr.Cells[16].Value = "X";

                    if (dr.Cells[1].Value.ToString() != "")
                    {
                        //   Console.WriteLine("dr.Cells[6].Value.ToString()" + dr.Cells[6].Value.ToString());
                        UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                        Page1 mPageC = new Page1();
                        mPageC.no = (dr.Index+1).ToString();
                        mPageC.usingAddress = dr.Cells[1].Value.ToString();
                        mPageC.APLink = dr.Cells[8].Value.ToString();
                        mPageC.TimerConnect = new System.Windows.Forms.Timer();
                        mPageC.TimerConnect.Interval = (30 * 1000);
                        mPageC.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                        mPageC.TimerSeconds = new Stopwatch();
                        mPageC.actionName = "reset";
                            foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                            {
                                if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == mPageC.usingAddress)
                                {
                                    if (drAP.Cells[8].Value.ToString() == "")
                                    {
                                        MessageBox.Show("請先配對ESL IP");
                                        break;
                                    }
                                    mPageC.APLink = drAP.Cells[8].Value.ToString();
                                    break;
                                }
                            }

                            foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                            {
                                if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPageC.APLink)
                                {
                                    if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                    {
                                        PageList.Add(mPageC);
                                    }
                                    else
                                    {
                                        if (eslAPNoSetMsg == null)
                                            eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                        if (eslAPNoSetMsg != null)
                                            eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                    }
                                }
                            }
                        }
                    else
                    {
                        if (nullMsg == null)
                        {
                            nullMsg = dr.Cells[1].Value.ToString();
                        }
                        else
                        {
                            nullMsg = nullMsg + "," + dr.Cells[1].Value.ToString();
                        }
                    }

                    //    dr.Cells[12].Value = "";


                }
            }
            
            //var sss = PageList.Distinct(x => x.APLink);

            if (nullMsg != null)
            {
              //  MessageBox.Show("勾選" + nullMsg + "未綁定電子標籤");
                PageList.Clear();
                dataGridView1.Enabled = true;
                nullMsg = null;
                datagridview1curr = 2;
                aaa(1, false, 0);
                DialogResult dialogResult = MessageBox.Show("勾選" + nullMsg + "未綁定電子標籤" + "\r\n" + "是否繼續綁定", "未綁定", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }
                if (eslAPNoSetMsg != null)
                {

                    //MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用");

                    dataGridView1.Enabled = true;
                    DialogResult dialogResult = MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用" + "\r\n" + "是否繼續執行", "AP未启用", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                    else if (PageList.Count==1)
                    {
                        return;
                    }
                    //return;
                }


                //----------------1/15
                // ------------------明天 初始修改
                if (PageList.Count !=0)
            {
                testest = true;
                onlockedbutton(testest);
                reset = true;
                UpdateESLDen.Text = PageList.Count.ToString();
                List<string> RunAPList = new List<string>();
                ProgressBarVisible(PageList.Count);
                List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).ToList();
                foreach (Page1 p in list)
                {
                    RunAPList.Add(p.APLink);

                }

                Page1 mPage1 = PageList[listcount];
                if (mPage1.usingAddress != "")
                {

                   

                   
                     for (int a = 0; a < RunAPList.Count; a++)
                       {
                        for (int i = 0; i < PageList.Count; i++)
                        {
                            if (PageList[i].APLink == RunAPList[a]) {
                                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                                {
                        //PageList[i].APLink
                        if (kvp.Key.Contains(PageList[i].APLink))
                                    {
                                       // listcount++;
                                        Bitmap bmp = mElectronicPriceData.writeIDimage(PageList[i].usingAddress);
                                        pictureBoxPage1.Image = bmp;
                                        int numVal = Convert.ToInt32(PageList[i].no) - 1;
                                        //  Console.WriteLine("mPage1.no" + mPage1.no);
                                        dataGridView4.Rows[numVal].Selected = true;
                                        dataGridView4.Rows[numVal].Cells[2].Value = DateTime.Now.ToString();
                                        Console.WriteLine("ININ"+ PageList[i].usingAddress+ PageList[i].APLink);
                                     //   kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                     //   kvp.Value.mSmcEsl.writeESLDataBuffer(PageList[i].usingAddress,0);
                                        pictureBoxPage1.Image = bmp;
                                        //System.Threading.Thread.Sleep(100);
                                        //    SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                        //  mSmcEsl.UpdataESLDataFromBuffer(PageList[i].usingAddress, 0, 3);
                                        richTextBox1.Text = richTextBox1.Text +PageList[i].usingAddress + "  嘗試連線中請稍候... \r\n";
                                        /*    mPage1.TimerConnect = new System.Windows.Forms.Timer();
                                            mPage1.TimerConnect.Interval = (30 * 1000);
                                            mPage1.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                            mPage1.TimerSeconds = new Stopwatch();*/
                                            mPage1.TimerSeconds.Start();
                                            mPage1.TimerConnect.Start();
                                            System.Threading.Thread.Sleep(1000);
                                        /*   if (PageList[listcount + 1].APLink  !=mPage1.APLink) {
                                               kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                               kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress);
                                               //System.Threading.Thread.Sleep(100);
                                               SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                               mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 8);
                                           }*/
                                    }
                               }
                                break;
                            }
                            

                        }
                    }
                   
                    // mSmcEsl.TransformImageToData(bmp);
                  

                    //   mSmcEsl.ConnectBleDevice(mPage1.usingAddress);
                    //  mSmcEsl.WriteESLData(mPage1.usingAddress);
                 //   Console.WriteLine("listcount" + listcount);
                    macaddress = PageList[listcount].usingAddress;
                  //  richTextBox1.Text = mPage1.usingAddress + "  嘗試連線中請稍候... \r\n";


                }
                else
                {
                    MessageBox.Show("該商品" + mPage1.product_name + "未裝置電子標籤");
                    dataGridView1.Enabled = true;
                    reset = false;
                }
            }
            }
            else
            {
                MessageBox.Show("ESL更新中請稍後", "更新中");
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Console.WriteLine("PageOneSelectAll_Click");
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }
            foreach (DataGridViewRow dr in this.dataGridView4.Rows)
            {

                if (dr.Index != dataGridView4.Rows.Count - 1)
                {
               
                    if (pageOneAll == false)
                    {
                        dr.Cells[0].Value = true;
                    }
                    else
                    {
                        dr.Cells[0].Value = false;
                    }
                }
            }
            pageOneAll = !pageOneAll;
        }

        private void checkESLRSSI_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }

            if (!APStart)
            {
                MessageBox.Show("請先連接AP");
               // dataGridView1.Enabled = true;
                return;
            }


            foreach (DataGridViewRow dr in this.dataGridView4.Rows) {
                if (dr.Index % 2 == 1)
                {
                    dr.Cells[1].Style.BackColor = Color.Beige;
                    dr.Cells[4].Style.BackColor = Color.Beige;
                    dr.Cells[7].Style.BackColor = Color.Beige;
                }
                else
                {
                    dr.Cells[1].Style.BackColor = Color.Bisque;
                    dr.Cells[4].Style.BackColor = Color.Bisque;
                    dr.Cells[7].Style.BackColor = Color.Bisque;
                }
                dr.DefaultCellStyle.ForeColor = Color.Black;
                dr.Cells[4].Value = "";
                dr.Cells[5].Value = "";
            }
            realESL.Text="0";
           
            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {
                kvp.Value.mSmcEsl.startScanBleDevice();
                Console.WriteLine("BBA");
            }
            checkESLRSSIClick = true;
            CheckESLStateTimer.Interval = 10000;
            CheckESLStateTimer.Start();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Console.WriteLine("Con");
            string eslAPNoSetMsg =null;
            if (e.ColumnIndex == 0)
            {

                if (!APStart)
                {
                    MessageBox.Show("請先連接AP"+ textBeforeEdit);
                    DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)dataGridView1.Rows[e.RowIndex].Cells[0];
                    if(chk.TrueValue== chk.Value)
                    {
                        chk.Value = chk.FalseValue;
                        chk.ReadOnly = true;
                    }
                    else
                        chk.Value = chk.TrueValue;

                    dataGridView1.Enabled = true;
                    return;
                }


                if (dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor == Color.Gray)
                {
                    return;
                }

           /*     if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                Console.WriteLine("(bool)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value"+ (bool)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);

                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null && !(bool)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value)
                {
                    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = false;
                    return;
                }*/
                else {

                 /*   foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {

                        kvp.Value.mSmcEsl.stopScanBleDevice();
                    }*/

                    DataGridViewRow drR = dataGridView1.Rows[e.RowIndex];

                    Page1 mPage = new Page1();
                    drR.Selected = false;
                    int aaaa = drR.Cells[1].Value.ToString().Length;
                    if (aaaa > 14)
                    {


                        Console.WriteLine("e.ColumnIndex" + e.ColumnIndex+ "e.RowIndex"+ e.RowIndex);
                        string[] drrow = drR.Cells[12].Value.ToString().Split(',');
                        drR.Cells[16].Value = "V";
                        for (int bb = 0; bb < drrow.Length; bb++)
                        {
                            Page1 mPageC = new Page1();
                            Console.WriteLine("usingAddress" + drrow[bb]);
                            Console.WriteLine("MT WRITE QQW" + drrow[bb]);
                            UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                            mPageC.no = (drR.Index + 1).ToString();
                            mPageC.BleAddress = drrow[bb];
                            mPageC.barcode = drR.Cells[5].Value.ToString();
                            mPageC.product_name = drR.Cells[6].Value.ToString();
                            mPageC.Brand = drR.Cells[7].Value.ToString();
                            mPageC.specification = drR.Cells[8].Value.ToString();
                            mPageC.price = drR.Cells[9].Value.ToString();
                            mPageC.Special_offer = drR.Cells[10].Value.ToString();
                            mPageC.Web = drR.Cells[11].Value.ToString();
                            mPageC.TimerConnect = new System.Windows.Forms.Timer();
                            mPageC.TimerConnect.Interval = (30 * 1000);
                            mPageC.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                            mPageC.TimerSeconds = new Stopwatch();
                            mPageC.HeadertextALL = headertextall;
                            mPageC.usingAddress = drrow[bb];
                            mPageC.actionName = "down";

                            foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                            {
                                if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == drrow[bb])
                                {
                                    if (drAP.Cells[8].Value.ToString() == "")
                                    {
                                        MessageBox.Show("請先配對ESL IP");
                                        break;
                                    }
                                    mPageC.APLink = drAP.Cells[8].Value.ToString();
                                    break;
                                }
                            }
                            PageList.Add(mPageC);
                        }
                    }
                    else
                    {
                        //  MessageBox.Show(" " + ((DataRowView)(dr.DataBoundItem)).Row.ItemArray[0] + " 被選取了！");
                        //字體   品名  品牌  規格  價格  特價  條碼  Qr
                        Console.WriteLine("AAAAAAAAAA");

                        if (drR.Cells[1].Value.ToString() == "")
                        {
                            mPage.no = (rownoinsertcount + rownonullinsert).ToString();
                            rownonullinsert++;
                        }
                        else
                        {
                            mPage.no = (drR.Index + 1).ToString();

                        }
                        UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();

                        Console.WriteLine("usingAddress" + drR.Cells[1].Value.ToString());
                        mPage.BleAddress = drR.Cells[1].Value.ToString();
                        mPage.barcode = drR.Cells[5].Value.ToString();
                        mPage.product_name = drR.Cells[6].Value.ToString();
                        mPage.Brand = drR.Cells[7].Value.ToString();
                        mPage.specification = drR.Cells[8].Value.ToString();
                        mPage.price = drR.Cells[9].Value.ToString();
                        mPage.Special_offer = drR.Cells[10].Value.ToString();
                        mPage.Web = drR.Cells[11].Value.ToString();
                        mPage.HeadertextALL = headertextall;
                        mPage.usingAddress = drR.Cells[12].Value.ToString();
                        mPage.TimerConnect = new System.Windows.Forms.Timer();
                        mPage.TimerConnect.Interval = (30 * 1000);
                        mPage.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                        mPage.TimerSeconds = new Stopwatch();
                        mPage.actionName = "down";
                        foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                        {
                            if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == drR.Cells[12].Value.ToString())
                            {
                                mPage.APLink = drAP.Cells[8].Value.ToString();
                                break;
                            }
                        }
                        drR.Cells[1].Value.ToString();
                        drR.Cells[16].Value = "V";
                        PageList.Add(mPage);
                        Console.WriteLine("BBBBBBBBBB");
                    }

                    progressBar1.Maximum = PageList.Count * 10;
                    //     Console.WriteLine("checkClick"+ checkClick);
                    if (!checkClick)
                    {
                        foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                        {

                            kvp.Value.mSmcEsl.stopScanBleDevice();
                        }
                        listcount = 0;
                        checkClick = true;
                        down = true;
                        testest = true;
                        SendData.Enabled = false;
                        button2.Enabled = false;
                        button15.Enabled = false;
                        button19.Enabled = false;
                        ProgressBarVisible(PageList.Count);
                        stopwatch.Reset();
                        stopwatch.Start();
                    }
                        List<string> RunAPList = new List<string>();
                        List<string> newAPList = new List<string>();
                    List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).Where(p => p.UpdateState == null).ToList();
                        foreach (Page1 p in list)
                        {
                            RunAPList.Add(p.APLink);
                        Console.WriteLine("RunAPList" + p.APLink + p.BleAddress);
                        }

                    for (int a=0;a< OldRunAPList.Count;a++)
                    {
                        Console.WriteLine("OldRunAPList" + OldRunAPList[a]);

                    }

                    Console.WriteLine("RunAPList.Except(OldRunAPList).Count()" + RunAPList.Except(OldRunAPList).Count());   
                        if (RunAPList.Count == OldRunAPList.Count && RunAPList.Except(OldRunAPList).Count()==0)
                        {
                            Console.WriteLine("一樣");
                        }
                        else
                        {
                            newAPList = RunAPList.Except(OldRunAPList).ToList();
                        Console.WriteLine("我們不一樣"+ newAPList);

                        for (int a = 0; a < newAPList.Count; a++)
                        {
                            Console.WriteLine("newAPList"+ newAPList[a]);
                            for (int i = 0; i < PageList.Count; i++)
                            {
                                if (PageList[i].APLink == newAPList[a])
                                {
                                    Page1 mPage1 = PageList[i];
                                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                                    {

                                        if (kvp.Key.Contains(mPage1.APLink))
                                        {


                                          //  int Blcount = mPage1.BleAddress.Length;
                                          //  Bitmap bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                                            /*  Bitmap bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                                        mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                         mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);*/
                                            int numVal = Convert.ToInt32(mPage1.no) - 1;
                                          //  Console.WriteLine("mPage1.no" + mPage1.no);
                                            dataGridView1.Rows[numVal].Cells[0].Selected = true;
                                            aaa(datagridview1curr, true, numVal);
                                            dataGridView1.Rows[numVal].Cells[17].Value = DateTime.Now.ToString();
                                            //   pictureBoxPage1.Image = bmp;

                                             Console.WriteLine("ININ"+ mPage1.BleAddress);
                                            deviceIPData = mPage1.APLink;
                                          //  ConnectBleTimeOut.Start();
                                            kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                                          /*  mPage1.TimerConnect = new System.Windows.Forms.Timer();
                                            mPage1.TimerConnect.Interval = (30 * 1000);
                                            mPage1.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                            mPage1.TimerSeconds = new Stopwatch();*/
                                            mPage1.TimerSeconds.Start();
                                            mPage1.TimerConnect.Start();
                                            //kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.usingAddress);
                                            // kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                            ///   kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress,0);
                                            //System.Threading.Thread.Sleep(1000);
                                            //   EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                            //  mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3,0);
                                            //  pictureBoxPage1.Image = bmp;
                                            richTextBox1.Text = richTextBox1.Text + PageList[i].usingAddress + "  嘗試連線中請稍候... \r\n";
                                           // System.Threading.Thread.Sleep(1000);
                                        }
                                    }
                                    break;
                                }
                            }

                        }

                    }

                        if (eslAPNoSetMsg != null)
                        {
                           // tt = 1;
                           // MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用");

                            dataGridView1.Enabled = true;
                        DialogResult dialogResult = MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用" + "\r\n" + "是否繼續執行", "AP未啟用", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            //do something
                        }
                        else if (dialogResult == DialogResult.No)
                        {
                            return;
                        }
                        //return;
                    }
                    
                 //   Console.WriteLine("OldRunAPList:" + OldRunAPList.Count + "RunAPList:" + RunAPList.Count);
                    OldRunAPList = RunAPList;
                   // Console.WriteLine("OldRunAPList:"+ OldRunAPList .Count+ "RunAPList:" + RunAPList.Count);

                   // for (int a = 0; a < RunAPList.Count; a++)
                 //   {
                  //      Console.WriteLine("aaaaaaaaaaaaaaaaa"+a);
                 //       OldRunAPList.Add(RunAPList[a]);
                 //   }
                    
                   // }
                    
                    //  mSmcEsl.TransformImageToData(bmp);
                    //  mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                  ///  macaddress = PageList[listcount].BleAddress;
                }
               
                //  int a = dataGridView1.RowCount - 2;
                
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (!autoNullESLMate)
            {
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    kvp.Value.mSmcEsl.startScanBleDevice();
                    Console.WriteLine("BBA");
                }
                autoNullESLData.Clear();
                foreach (DataGridViewRow dr4 in this.dataGridView4.Rows) {
                    if (dr4.Cells[1].Value != null && dr4.Cells[8].Value != null && dr4.Cells[8].Value.ToString() == "")
                        autoNullESLData.Add(dr4.Cells[1].Value.ToString());
                }
                autoNullESLMate = true;
                button17.Text = "停止配對";
                button17.ForeColor = Color.Red;
            }
            else
            {
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {

                    kvp.Value.mSmcEsl.stopScanBleDevice();
                }
                //mExcelData.DataGridview4Update(dataGridView4, false, openExcelAddress);
                autoNullESLMate = false;
                button17.Text = "自動AP配對";
                button17.ForeColor = Color.Black;
            }
        }

        private void scancode_KeyPress(object sender, KeyPressEventArgs e)
        {
            Console.WriteLine("KeyPress"+ e.KeyChar);
            DialogResult dialogResult = new DialogResult();

            if ((e.KeyChar >= 'a' && e.KeyChar <= 'z') || (e.KeyChar >= 'A' && e.KeyChar <= 'Z') || e.KeyChar.CompareTo('0') > 0 || e.KeyChar.CompareTo('9') < 0)
            {
                Console.WriteLine("ININKeyPress" + e.KeyChar);
            }
            else
            {
                scanstate.Text = "輸入法請切換英文。";
                scanstate.ForeColor = Color.Red;
                return;
            }
                
            if (e.KeyChar == 13) {

                    CheckSyntaxAndReport();
            }
        }

        private void dataGridView7_CurrentCellChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr2 in dataGridView2.Rows)
            {
                dr2.Cells[1].Selected = false;
            }
            bbb(datagridview2curr, dataGridView7);
        }

        private void dataGridView7_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {


            foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
            {
                dr2.Selected=false;
            }
            
            if (e.ColumnIndex == 0)
            {
                foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
                {
                    dr2.Cells[0].Value = false;
                }
                foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
                {
                    dr7.Cells[0].Value = false;
                }
                dataGridView7.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = true;
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }

            string value = "";
            if (InputBox("textbox", "新增特價版型", "版型名稱:", ref value) == DialogResult.OK)
            {

                if (dialogtext == "")
                {
                    MessageBox.Show("不能為空值");
                    return;
                }

                DataGridView dataGridView = dataGridView7;
                
                pictureBox1.Image = null;
                pictureBox1.Image = null;
                pictureBox1.Image = null;
                for (int i = 0; i < picturelabel; i++)
                {
                    foreach (Control x in pictureBox1.Controls)
                    {
                        //      Console.WriteLine("x.Name" + x.Name);
                        x.Dispose();
                    }

                }


                if (selectSize == "2.13")
                {
                    pictureBox1.BackColor = Color.White;
                    pictureBox1.Size = new Size(212, 104);
                    pictureBox1.Location = new Point(235, 81);
                    panel1.Controls.Add(pictureBox1);
                }
                else if (selectSize == "2.9")
                {
                    pictureBox1.BackColor = Color.White;
                    pictureBox1.Size = new Size(296, 128);
                    pictureBox1.Location = new Point(151, 59);
                    panel1.Controls.Add(pictureBox1);
                }
                else if (selectSize == "4.2")
                {
                    pictureBox1.BackColor = Color.White;
                    pictureBox1.Size = new Size(400, 300);
                    pictureBox1.Location = new Point(88, 5);
                    panel1.Controls.Add(pictureBox1);
                }


                foreach (DataGridViewRow dr7 in dataGridView7.Rows)
                {
                    dr7.Selected = false;

                }
                foreach (DataGridViewRow dr2 in dataGridView2.Rows)
                {
                    dr2.Selected = false;

                }

                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                {
                        dr8.Cells[0].Value = false;
                }
                label1.Text = dialogtext;
                ESLStyleSave = false;
                ESLSaleStyleSave = true;
                picturelabel = 0;
                if (2 > 1 && dataGridView.CurrentCellAddress.Y != -1)
                {

                    int datacount = 0;
                    int currentRow = 0;
                    int currentColumn = dataGridView.Columns.Count;
                    //   Console.WriteLine("currentColumn-------------" + currentColumn);
                    for (int g = 0; g < currentColumn; g++)
                    {
                        if (g != 0)
                        {
                            if (dataGridView.Rows[currentRow].Cells[g].Value == null)
                                break;
                            if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "L" || dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "Header")
                            {

                                Label LabelDemo = new Label();
                                LabelDemo.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                //          Console.WriteLine("LabelDemo.Name" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                LabelDemo.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                g = g + 2;

                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (LabelDemo.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                //        Console.WriteLine("LabelDemo.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                LabelDemo.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                LabelDemo.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                LabelDemo.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                //LabelDemo.Name = labelname;
                                //  LabelDemo.Text = labeldata;
                                pictureBox1.Controls.Add(LabelDemo);
                                LabelDemo.AutoSize = true;
                                LabelDemo.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                LabelDemo.TextAlign = ContentAlignment.MiddleCenter;
                                g = g + 2;
                                LabelDemo.Font = new Font(dataGridView.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Regular"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Regular);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Bold"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Bold);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Italic"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Italic);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Underline"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Underline);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Strikeout"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Strikeout);
                                // if (Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g + 2].Value) == 3)
                                // LabelDemo.Font = new Font(dataGridView2.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g+1].Value),(FontStyle)dataGridView2.Rows[currentRow].Cells[g + 2].Value);
                                g = g + 3;
                                LabelDemo.ForeColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                g = g + 4;
                                LabelDemo.BackColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                LabelDemo.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                LabelDemo.MouseMove += new System.Windows.Forms.MouseEventHandler(Label_MouseMove);
                                LabelDemo.Click += new System.EventHandler(this.LabelDemo_Click);

                            }
                            if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "T")
                            {

                                TextBox LabelDemo = new TextBox();
                                LabelDemo.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                //    Console.WriteLine("LabelDemo.Name" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                LabelDemo.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                g = g + 2;

                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (LabelDemo.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                //    Console.WriteLine("LabelDemo.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                LabelDemo.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                LabelDemo.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                LabelDemo.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                //LabelDemo.Name = labelname;
                                //  LabelDemo.Text = labeldata;
                                pictureBox1.Controls.Add(LabelDemo);
                                LabelDemo.AutoSize = true;
                                LabelDemo.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                LabelDemo.TextAlign = HorizontalAlignment.Center;
                                g = g + 2;
                                LabelDemo.Font = new Font(dataGridView.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Regular"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Regular);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Bold"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Bold);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Italic"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Italic);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Underline"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Underline);
                                if (dataGridView.Rows[currentRow].Cells[g + 2].Value.ToString().Contains("Strikeout"))
                                    LabelDemo.Font = new Font(LabelDemo.Font, LabelDemo.Font.Style | FontStyle.Strikeout);
                                // if (Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g + 2].Value) == 3)
                                // LabelDemo.Font = new Font(dataGridView2.Rows[currentRow].Cells[g].Value.ToString(), Convert.ToInt32(dataGridView2.Rows[currentRow].Cells[g+1].Value),(FontStyle)dataGridView2.Rows[currentRow].Cells[g + 2].Value);
                                g = g + 3;
                                LabelDemo.ForeColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                g = g + 4;
                                LabelDemo.BackColor = Color.FromArgb(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 2].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 3].Value));
                                LabelDemo.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                LabelDemo.MouseMove += new System.Windows.Forms.MouseEventHandler(TextBox_MouseMove);
                                LabelDemo.Click += new System.EventHandler(this.TextBoxDemo_Click);

                            }
                            else if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "B")
                            {
                                //pictureBox2.AutoSize = true;
                                Regex NumandEG = new Regex("[^A-Za-z0-9]");
                                bool type;
                                string barcodevalue;
                                pictureboxBarcode pictureBox2 = new pictureboxBarcode();
                                pictureBox2.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                //  Console.WriteLine(" pictureBox2.Name" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                pictureBox2.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                g = g + 2;

                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (pictureBox2.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                //   Console.WriteLine(" pictureBox2.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                pictureBox2.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                type = NumandEG.IsMatch(dataGridView.Rows[currentRow].Cells[g].Value.ToString());
                                barcodevalue = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                pictureBox1.Controls.Add(pictureBox2);
                                //   Console.WriteLine("barcode_w.Options.Width" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                // Console.WriteLine("barcode_w.Options.Height" + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                pictureBox2.Size = new System.Drawing.Size(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                pictureBox2.SizeMode = PictureBoxSizeMode.CenterImage;
                                Bitmap bar = new Bitmap(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                BarcodeWriter barcode_w = new BarcodeWriter();       // 建立條碼物件
                                                                                     //   if (type)
                                                                                     //  {
                                barcode_w.Format = BarcodeFormat.CODE_93;
                                // }
                                //      else {
                                //         barcode_w.Format = BarcodeFormat.EAN_13;            // 條碼類別.
                                //    }

                                barcode_w.Options.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                barcode_w.Options.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                barcode_w.Options.PureBarcode = true;               // 顯示條碼字串
                                                                                    //   Console.WriteLine("pictureBox2.Location " + dataGridView2.Rows[currentRow].Cells[g].Value.ToString() + dataGridView2.Rows[currentRow].Cells[g + 1].Value.ToString());
                                pictureBox2.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                g = g + 12;
                                bar = barcode_w.Write("4253786521345");
                                pictureBox2.barcodedata = "4253786521345";
                                pictureBox2.Image = bar;
                                pictureBox2.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                pictureBox2.MouseMove += new System.Windows.Forms.MouseEventHandler(picture_MouseMove);
                                pictureBox2.Click += new System.EventHandler(this.PictureBoxDemo_Click);
                                pictureBox2.SizeChanged += new System.EventHandler(this.pictureBox2B_SizeChanged);
                            }
                            else if (dataGridView.Rows[currentRow].Cells[g].Value.ToString() == "Q")
                            {
                                pictureboxBarcode pictureBox2 = new pictureboxBarcode();
                                pictureBox2.Tag = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                pictureBox2.Name = dataGridView.Rows[currentRow].Cells[g + 1].Value.ToString();
                                g = g + 2;

                                foreach (DataGridViewRow dr8 in dataGridView8.Rows)
                                {
                                    if (pictureBox2.Name == dr8.Cells[1].Value.ToString())
                                    {
                                        dr8.Cells[0].Value = true;
                                    }
                                }
                                //    Console.WriteLine(" pictureBox2.Text" + dataGridView2.Rows[currentRow].Cells[g].Value.ToString());
                                pictureBox2.Text = dataGridView.Rows[currentRow].Cells[g].Value.ToString();
                                g++;
                                pictureBox1.Controls.Add(pictureBox2);
                                pictureBox2.Size = new System.Drawing.Size(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                pictureBox2.SizeMode = PictureBoxSizeMode.CenterImage;
                                Bitmap bqr = new Bitmap(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                BarcodeWriter qr = new BarcodeWriter();       // 建立條碼物件
                                qr.Format = BarcodeFormat.QR_CODE;
                                qr.Options.Width = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                qr.Options.Height = Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value);
                                g++;
                                qr.Options.Margin = 0;
                                pictureBox2.Location = new Point(Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g].Value), Convert.ToInt32(dataGridView.Rows[currentRow].Cells[g + 1].Value));
                                g = g + 12;
                                bqr = qr.Write("http://www.smartchip.com.tw");
                                pictureBox2.barcodedata = "http://www.smartchip.com.tw";
                                pictureBox2.Image = bqr;
                                pictureBox2.MouseDown += new System.Windows.Forms.MouseEventHandler(Label_MouseDown);
                                pictureBox2.MouseMove += new System.Windows.Forms.MouseEventHandler(picture_MouseMove);
                                pictureBox2.Click += new System.EventHandler(this.PictureBoxDemo_Click);
                                pictureBox2.SizeChanged += new System.EventHandler(this.pictureBox2Q_SizeChanged);
                            }
                            picturelabel++;
                        }
                    }
                    // Console.WriteLine("currentRow" + currentRow);
                }
            }
         }

        private void dataGridView8_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
            if (e.ColumnIndex == 0)
            {
                if (dataGridView8.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null && (bool)dataGridView8.Rows[e.RowIndex].Cells[e.ColumnIndex].Value== false)
                {
                    if (dataGridView8.Rows[e.RowIndex].Cells[1].Value != null) {
                        Console.WriteLine("AEAELOK"+ dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString());
                        if (dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "品名(最多10字)"||dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "售價" || dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "促銷價" || dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "ESL ID" || dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "說明文字" || dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "Qrcode 網址")
                        {
                            if (dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "說明文字")
                            {
                                Labelcreate("TEXT", "TEXT",true);
                            }

                            if (dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "售價")
                            {
                                Labelcreate("售價",dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString(),true);
                                Labelcreate("0", dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString(),false);

                                bool aas = false;
                                foreach (Control x in pictureBox1.Controls)
                                {
                                    
                                    if (x.Name == "促銷價")
                                    {
                                        aas = true;
                                    }
                                }
                                foreach (Control x in pictureBox1.Controls)
                                {

                                    if (aas == true&& x.Name == "售價")
                                    {
                                        if (x.Tag.ToString() == "Header")
                                        {
                                            x.Text = "售價:";
                                            x.Location = new Point(8, 58);
                                        }
                                        else
                                        {
                                            x.ForeColor = Color.Black;
                                            x.Font = new Font("Calibri", 12, FontStyle.Regular);
                                            x.Location = new Point(39, 54);
                                        }
                                    }
                                }
                            }

                            if (dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "促銷價")
                            {
                                
                                Labelcreate("促銷價", dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString(), true);
                                Labelcreate("0", dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString(), false);
                                foreach (Control x in pictureBox1.Controls)
                                {
                                    if (x.Name == "售價")
                                    {
                                        if (x.Tag.ToString() == "Header")
                                        {
                                            x.Text = "售價:";
                                            x.Location = new Point(8, 58);
                                        }
                                        else
                                        {
                                            x.ForeColor = Color.Black;
                                            x.Font = new Font("Calibri", 12, FontStyle.Regular);
                                            x.Location = new Point(39, 54);
                                        }
                                    }
                                }

                            }

                            if (dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "ESL ID")
                            {
                                codecreate(dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString(), dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString(),true);
                            }

                            if (dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "Qrcode 網址")
                            {
                                codecreate(dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString(), dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString(), false);
                            }

                            if (dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() == "品名(最多10字)")
                            {
                                TextBoxcreate(dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString(), "品名");
                            }
                            
                            dataGridView8.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = true;
                        }
                        else
                        {
                            Labelcreate(dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString() + ":", dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString(),false);
                            dataGridView8.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = true;
                        }
                        
                    }
                }
                else
                {

                    bool bb = false;
                    for (int i=0;i< pictureBox1.Controls.Count;i++)
                    {
                    foreach (Control x in pictureBox1.Controls)
                    {
                        Console.WriteLine("x.Name" + x.Name+ "x.Text" + x.Text);
                        if(x.Name== dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString()) {
                                if (x.Name == "促銷價")
                                {
                                    bb = true;
                                }
                            Console.WriteLine("froze");
                            x.Dispose();
                                dataGridView8.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = false;
                        }
                    }
                    }
                    foreach (Control x in pictureBox1.Controls)
                    {
                            if (x.Name == "售價"&& bb)
                            {
                                bb = true;
                            if (x.Tag.ToString() == "Header")
                            {
                                x.Text = "售價";
                                x.Location = new Point(153, 24);
                            }
                            else
                            {
                                x.Font = new Font("Calibri", 26,FontStyle.Bold);
                                x.ForeColor = Color.Red;
                                x.Location = new Point(153, 35);
                            }
                        }
                    }

                }
               
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string ESLStyle="";
            string ESLSaleStyle="";
            string eslAPNoSetMsg = null;
            string eslVState = null;
            string eslNotMateAP = null;
            PageList.Clear();
            if (!testest) {
                nullMsg = null;
                foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
            {
                if (dr2.Cells[2].Value.ToString()=="V")
                {
                    ESLStyle = dr2.Cells[1].Value.ToString();
                }
            }

            foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
            {
                if (dr7.Cells[2].Value.ToString() == "V")
                {
                    ESLSaleStyle = dr7.Cells[1].Value.ToString();
                }
            }

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先載入資料表");
                dataGridView1.Enabled = true;
                return;
            }
            if (!APStart)
            {
                MessageBox.Show("請先連接AP");
                dataGridView1.Enabled = true;
                return;
            }
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[13].Value!=null&&dr.Cells[13].Value.ToString() != "")
                {
                    if (!ESLStyleDataChange && !ESLSaleStyleDataChange)
                    {

                            Console.WriteLine("ESLSAVE1");
                    }
                    else if (ESLStyleDataChange && !ESLSaleStyleDataChange)
                    {
                            Console.WriteLine("ESLSAVE2");
                            if (dr.Cells[15].Value.ToString() == "X")
                        {
                            Page1 mPage = new Page1();
                            dr.Selected = false;
                            int aaaa = dr.Cells[1].Value.ToString().Length;
                            if (aaaa > 14)
                            {


                                string[] drrow = dr.Cells[1].Value.ToString().Split(',');
                                //dr.Cells[16].Value = "V";
                                for (int bb = 0; bb < drrow.Length; bb++)
                                {
                                    Page1 mPageC = new Page1();

                                    Console.WriteLine("MT WRITE QQW" + drrow[bb]);
                                    UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                    mPageC.no = (dr.Index + 1).ToString();
                                    mPageC.BleAddress = drrow[bb];
                                    mPageC.barcode = dr.Cells[5].Value.ToString();
                                    mPageC.product_name = dr.Cells[6].Value.ToString();
                                    mPageC.Brand = dr.Cells[7].Value.ToString();
                                    mPageC.specification = dr.Cells[8].Value.ToString();
                                    mPageC.price = dr.Cells[9].Value.ToString();
                                    mPageC.Special_offer = dr.Cells[10].Value.ToString();
                                    mPageC.Web = dr.Cells[11].Value.ToString();
                                    mPageC.ProductStyle = styleName;
                                    mPageC.HeadertextALL = headertextall;
                                    mPageC.usingAddress = drrow[bb];
                                    mPageC.onsale = dr.Cells[15].Value.ToString();
                                    mPageC.TimerConnect = new System.Windows.Forms.Timer();
                                    mPageC.TimerConnect.Interval = (30 * 1000);
                                    mPageC.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                    mPageC.TimerSeconds = new Stopwatch();
                                    mPageC.actionName = "EslStyleChangeUpdate";
                                    foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                    {
                                        if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == drrow[bb])
                                        {
                                                if (drAP.Cells[8].Value.ToString() == "")
                                                {
                                                    if (eslNotMateAP == null)
                                                        eslNotMateAP = drrow[bb];
                                                    else
                                                        eslNotMateAP = eslNotMateAP + "," + drrow[bb];
                                                    // MessageBox.Show("請先配對ESL IP");
                                                    //break;
                                                }
                                                if (drAP.Cells[5].Value.ToString() != "" && Convert.ToDouble(drAP.Cells[5].Value) < 2.85)
                                                {
                                                    if (eslVState == null)
                                                        eslVState = drrow[bb];
                                                    else
                                                        eslVState = eslVState + "," + drrow[bb];
                                                    // MessageBox.Show("請先配對ESL IP");
                                                    //break;
                                                }
                                                mPageC.APLink = drAP.Cells[8].Value.ToString();
                                            break;
                                        }
                                    }
                                        foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                        {
                                            if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPageC.APLink)
                                            {
                                                if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                                {
                                                    PageList.Add(mPageC);
                                                }
                                                else
                                                {
                                                    if (eslAPNoSetMsg == null)
                                                        eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                                    if (eslAPNoSetMsg != null)
                                                        eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                                }
                                            }
                                        }
                                    }
                            }
                            else
                            {
                                //  MessageBox.Show(" " + ((DataRowView)(dr.DataBoundItem)).Row.ItemArray[0] + " 被選取了！");
                                //字體   品名  品牌  規格  價格  特價  條碼  Qr
                                Console.WriteLine("AAAAAAAAAA");

                                if (dr.Cells[1].Value.ToString() == "")
                                {
                                    mPage.no = (rownoinsertcount + rownonullinsert).ToString();
                                    rownonullinsert++;
                                }
                                else
                                {
                                    mPage.no = (dr.Index + 1).ToString();

                                }
                                UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                mPage.BleAddress = dr.Cells[1].Value.ToString();
                                mPage.barcode = dr.Cells[5].Value.ToString();
                                mPage.product_name = dr.Cells[6].Value.ToString();
                                mPage.Brand = dr.Cells[7].Value.ToString();
                                mPage.specification = dr.Cells[8].Value.ToString();
                                mPage.price = dr.Cells[9].Value.ToString();
                                mPage.Special_offer = dr.Cells[10].Value.ToString();
                                mPage.Web = dr.Cells[11].Value.ToString();
                                mPage.ProductStyle = styleName;
                                mPage.HeadertextALL = headertextall;
                                mPage.usingAddress = dr.Cells[1].Value.ToString();
                                mPage.onsale = dr.Cells[15].Value.ToString();
                                mPage.TimerConnect = new System.Windows.Forms.Timer();
                                mPage.TimerConnect.Interval = (30 * 1000);
                                mPage.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                mPage.TimerSeconds = new Stopwatch();
                                mPage.actionName = "EslStyleChangeUpdate";
                                foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                {
                                    if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == dr.Cells[1].Value.ToString())
                                    {
                                            if (drAP.Cells[8].Value.ToString() == "")
                                            {
                                                if (eslNotMateAP == null)
                                                    eslNotMateAP = dr.Cells[1].Value.ToString();
                                                else
                                                    eslNotMateAP = eslNotMateAP + "," + dr.Cells[1].Value.ToString();
                                                // MessageBox.Show("請先配對ESL IP");
                                                //break;
                                            }
                                            if (drAP.Cells[5].Value.ToString() != "" && Convert.ToDouble(drAP.Cells[5].Value) < 2.85)
                                            {
                                                if (eslVState == null)
                                                    eslVState = dr.Cells[1].Value.ToString();
                                                else
                                                    eslVState = eslVState + "," + dr.Cells[1].Value.ToString();
                                                // MessageBox.Show("請先配對ESL IP");
                                                //break;
                                            }
                                            mPage.APLink = drAP.Cells[8].Value.ToString();
                                        break;
                                    }
                                }
                                    foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                    {
                                        if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPage.APLink)
                                        {
                                            if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                            {
                                                PageList.Add(mPage);
                                            }
                                            else
                                            {
                                                if (eslAPNoSetMsg == null)
                                                    eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                                if (eslAPNoSetMsg != null)
                                                    eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                            }
                                        }
                                    }
                                    Console.WriteLine("BBBBBBBBBB");
                            }
                        }
                      /*  else
                        {
                            if (nullMsg == null)
                            {
                                nullMsg = dr.Cells[6].Value.ToString();
                            }
                            else
                            {
                                nullMsg = nullMsg + "," + dr.Cells[6].Value.ToString();
                            }
                        }*/
                    }
                    else if (!ESLStyleDataChange && ESLSaleStyleDataChange)
                    {
                            Console.WriteLine("ESLSAVE3");
                            if (dr.Cells[15].Value.ToString() == "V")
                        {
                            Page1 mPage = new Page1();
                            dr.Selected = false;
                            int aaaa = dr.Cells[1].Value.ToString().Length;
                            if (aaaa > 14)
                            {


                                string[] drrow = dr.Cells[1].Value.ToString().Split(',');
                                //dr.Cells[16].Value = "V";
                                for (int bb = 0; bb < drrow.Length; bb++)
                                {
                                    Page1 mPageC = new Page1();

                                    Console.WriteLine("MT WRITE QQW" + drrow[bb]);
                                    UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                    mPageC.no = (dr.Index + 1).ToString();
                                    mPageC.BleAddress = drrow[bb];
                                    mPageC.barcode = dr.Cells[5].Value.ToString();
                                    mPageC.product_name = dr.Cells[6].Value.ToString();
                                    mPageC.Brand = dr.Cells[7].Value.ToString();
                                    mPageC.specification = dr.Cells[8].Value.ToString();
                                    mPageC.price = dr.Cells[9].Value.ToString();
                                    mPageC.Special_offer = dr.Cells[10].Value.ToString();
                                    mPageC.Web = dr.Cells[11].Value.ToString();
                                    mPageC.ProductStyle = styleSaleName;
                                    mPageC.HeadertextALL = headertextall;
                                    mPageC.usingAddress = drrow[bb];
                                    mPageC.onsale = dr.Cells[15].Value.ToString();
                                    mPageC.TimerConnect = new System.Windows.Forms.Timer();
                                    mPageC.TimerConnect.Interval = (30 * 1000);
                                    mPageC.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                    mPageC.TimerSeconds = new Stopwatch();
                                    mPageC.actionName = "EslStyleChangeUpdate";
                                    foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                    {
                                        if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == drrow[bb])
                                        {
                                                if (drAP.Cells[8].Value.ToString() == "")
                                                {
                                                    if (eslNotMateAP == null)
                                                        eslNotMateAP = dr.Cells[1].Value.ToString();
                                                    else
                                                        eslNotMateAP = eslNotMateAP + "," + dr.Cells[1].Value.ToString();
                                                    // MessageBox.Show("請先配對ESL IP");
                                                    //break;
                                                }
                                                if (drAP.Cells[5].Value.ToString() != "" && Convert.ToDouble(drAP.Cells[5].Value) < 2.85)
                                                {
                                                    if (eslVState == null)
                                                        eslVState = dr.Cells[1].Value.ToString();
                                                    else
                                                        eslVState = eslVState + "," + dr.Cells[1].Value.ToString();
                                                    // MessageBox.Show("請先配對ESL IP");
                                                    //break;
                                                }
                                                mPageC.APLink = drAP.Cells[8].Value.ToString();
                                            break;
                                        }
                                    }
                                        foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                        {
                                            if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPageC.APLink)
                                            {
                                                if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                                {
                                                    PageList.Add(mPageC);
                                                }
                                                else
                                                {
                                                    if (eslAPNoSetMsg == null)
                                                        eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                                    if (eslAPNoSetMsg != null)
                                                        eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                                }
                                            }
                                        }
                                    }
                            }
                            else
                            {
                                //  MessageBox.Show(" " + ((DataRowView)(dr.DataBoundItem)).Row.ItemArray[0] + " 被選取了！");
                                //字體   品名  品牌  規格  價格  特價  條碼  Qr
                                Console.WriteLine("AAAAAAAAAA");

                                if (dr.Cells[1].Value.ToString() == "")
                                {
                                    mPage.no = (rownoinsertcount + rownonullinsert).ToString();
                                    rownonullinsert++;
                                }
                                else
                                {
                                    mPage.no = (dr.Index + 1).ToString();

                                }
                                UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                mPage.BleAddress = dr.Cells[1].Value.ToString();
                                mPage.barcode = dr.Cells[5].Value.ToString();
                                mPage.product_name = dr.Cells[6].Value.ToString();
                                mPage.Brand = dr.Cells[7].Value.ToString();
                                mPage.specification = dr.Cells[8].Value.ToString();
                                mPage.price = dr.Cells[9].Value.ToString();
                                mPage.Special_offer = dr.Cells[10].Value.ToString();
                                mPage.Web = dr.Cells[11].Value.ToString();
                                mPage.ProductStyle = styleSaleName;
                                mPage.HeadertextALL = headertextall;
                                mPage.usingAddress = dr.Cells[1].Value.ToString();
                                mPage.onsale = dr.Cells[15].Value.ToString();
                                mPage.TimerConnect = new System.Windows.Forms.Timer();
                                mPage.TimerConnect.Interval = (30 * 1000);
                                mPage.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                mPage.TimerSeconds = new Stopwatch();
                                mPage.actionName = "EslStyleChangeUpdate";
                                foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                {
                                    if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == dr.Cells[1].Value.ToString())
                                    {
                                            if (drAP.Cells[8].Value.ToString() == "")
                                            {
                                                if (eslNotMateAP == null)
                                                    eslNotMateAP = dr.Cells[1].Value.ToString();
                                                else
                                                    eslNotMateAP = eslNotMateAP + "," + dr.Cells[1].Value.ToString();
                                                // MessageBox.Show("請先配對ESL IP");
                                                //break;
                                            }
                                            if (drAP.Cells[5].Value.ToString() != "" && Convert.ToDouble(drAP.Cells[5].Value) < 2.85)
                                            {
                                                if (eslVState == null)
                                                    eslVState = dr.Cells[1].Value.ToString();
                                                else
                                                    eslVState = eslVState + "," + dr.Cells[1].Value.ToString();
                                                // MessageBox.Show("請先配對ESL IP");
                                                //break;
                                            }
                                            mPage.APLink = drAP.Cells[8].Value.ToString();
                                        break;
                                    }
                                }
                                    foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                    {
                                        if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPage.APLink)
                                        {
                                            if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                            {
                                                PageList.Add(mPage);
                                            }
                                            else
                                            {
                                                if (eslAPNoSetMsg == null)
                                                    eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                                if (eslAPNoSetMsg != null)
                                                    eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                            }
                                        }
                                    }
                                    Console.WriteLine("BBBBBBBBBB");
                            }
                        }
                    /*    else
                        {
                            if (nullMsg == null)
                            {
                                nullMsg = dr.Cells[6].Value.ToString();
                            }
                            else
                            {
                                nullMsg = nullMsg + "," + dr.Cells[6].Value.ToString();
                            }
                        }*/
                    }
                    else
                    {
                            Console.WriteLine("ESLSAVE4");
                            if (dr.Cells[15].Value.ToString() == "V"|| dr.Cells[15].Value.ToString() == "X")
                        {
                            Page1 mPage = new Page1();
                            dr.Selected = false;
                            int aaaa = dr.Cells[1].Value.ToString().Length;
                            if (aaaa > 14)
                            {


                                string[] drrow = dr.Cells[1].Value.ToString().Split(',');
                                //dr.Cells[16].Value = "V";
                                for (int bb = 0; bb < drrow.Length; bb++)
                                {
                                    Page1 mPageC = new Page1();

                                    Console.WriteLine("MT WRITE QQW" + drrow[bb]);
                                    UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                    mPageC.no = (dr.Index + 1).ToString();
                                    mPageC.BleAddress = drrow[bb];
                                    mPageC.barcode = dr.Cells[5].Value.ToString();
                                    mPageC.product_name = dr.Cells[6].Value.ToString();
                                    mPageC.Brand = dr.Cells[7].Value.ToString();
                                    mPageC.specification = dr.Cells[8].Value.ToString();
                                    mPageC.price = dr.Cells[9].Value.ToString();
                                    mPageC.Special_offer = dr.Cells[10].Value.ToString();
                                    mPageC.Web = dr.Cells[11].Value.ToString();

                                    mPageC.TimerConnect = new System.Windows.Forms.Timer();
                                    mPageC.TimerConnect.Interval = (30 * 1000);
                                    mPageC.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                    mPageC.TimerSeconds = new Stopwatch();
                                    if (dr.Cells[15].Value.ToString() == "V")
                                    {
                                        mPageC.ProductStyle = styleSaleName;
                                    }
                                    else
                                    {
                                        mPageC.ProductStyle = styleName;
                                    }
                                    mPageC.ProductStyle = dr.Cells[13].Value.ToString();
                                    mPageC.HeadertextALL = headertextall;
                                    mPageC.usingAddress = drrow[bb];
                                    mPageC.onsale = dr.Cells[15].Value.ToString();
                                    mPageC.actionName = "EslStyleChangeUpdate";

                                    foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                    {
                                        if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == drrow[bb])
                                        {
                                                if (drAP.Cells[8].Value.ToString() == "")
                                                {
                                                    if (eslNotMateAP == null)
                                                        eslNotMateAP = dr.Cells[1].Value.ToString();
                                                    else
                                                        eslNotMateAP = eslNotMateAP + "," + dr.Cells[1].Value.ToString();
                                                    // MessageBox.Show("請先配對ESL IP");
                                                    //break;
                                                }
                                                if (drAP.Cells[5].Value.ToString() != "" && Convert.ToDouble(drAP.Cells[5].Value) < 2.85)
                                                {
                                                    if (eslVState == null)
                                                        eslVState = dr.Cells[1].Value.ToString();
                                                    else
                                                        eslVState = eslVState + "," + dr.Cells[1].Value.ToString();
                                                    // MessageBox.Show("請先配對ESL IP");
                                                    //break;
                                                }
                                                mPageC.APLink = drAP.Cells[8].Value.ToString();
                                            break;
                                        }
                                    }
                                        foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                        {
                                            if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPageC.APLink)
                                            {
                                                if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                                {
                                                    PageList.Add(mPageC);
                                                }
                                                else
                                                {
                                                    if (eslAPNoSetMsg == null)
                                                        eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                                    if (eslAPNoSetMsg != null)
                                                        eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                                }
                                            }
                                        }
                                    }
                            }
                            else
                            {
                                //  MessageBox.Show(" " + ((DataRowView)(dr.DataBoundItem)).Row.ItemArray[0] + " 被選取了！");
                                //字體   品名  品牌  規格  價格  特價  條碼  Qr
                                Console.WriteLine("AAAAAAAAAA");

                                if (dr.Cells[1].Value.ToString() == "")
                                {
                                    mPage.no = (rownoinsertcount + rownonullinsert).ToString();
                                    rownonullinsert++;
                                }
                                else
                                {
                                    mPage.no = (dr.Index + 1).ToString();

                                }
                                UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                mPage.BleAddress = dr.Cells[1].Value.ToString();
                                mPage.barcode = dr.Cells[5].Value.ToString();
                                mPage.product_name = dr.Cells[6].Value.ToString();
                                mPage.Brand = dr.Cells[7].Value.ToString();
                                mPage.specification = dr.Cells[8].Value.ToString();
                                mPage.price = dr.Cells[9].Value.ToString();
                                mPage.Special_offer = dr.Cells[10].Value.ToString();
                                mPage.Web = dr.Cells[11].Value.ToString();
                                if (dr.Cells[15].Value.ToString() == "V")
                                {
                                    mPage.ProductStyle = styleSaleName;
                                }
                                else
                                {
                                    mPage.ProductStyle = styleName;
                                }
                                mPage.HeadertextALL = headertextall;
                                mPage.usingAddress = dr.Cells[1].Value.ToString();
                                mPage.onsale = dr.Cells[15].Value.ToString();
                                mPage.TimerConnect = new System.Windows.Forms.Timer();
                                mPage.TimerConnect.Interval = (30 * 1000);
                                mPage.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                mPage.TimerSeconds = new Stopwatch();
                                mPage.actionName = "EslStyleChangeUpdate";
                                foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                                {
                                    if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == dr.Cells[1].Value.ToString())
                                    {
                                            if (drAP.Cells[8].Value.ToString() == "")
                                            {
                                                if (eslNotMateAP == null)
                                                    eslNotMateAP = dr.Cells[1].Value.ToString();
                                                else
                                                    eslNotMateAP = eslNotMateAP + "," + dr.Cells[1].Value.ToString();
                                                // MessageBox.Show("請先配對ESL IP");
                                                //break;
                                            }
                                            if (drAP.Cells[5].Value.ToString() != "" && Convert.ToDouble(drAP.Cells[5].Value) < 2.85)
                                            {
                                                if (eslVState == null)
                                                    eslVState = dr.Cells[1].Value.ToString();
                                                else
                                                    eslVState = eslVState + "," + dr.Cells[1].Value.ToString();
                                                // MessageBox.Show("請先配對ESL IP");
                                                //break;
                                            }
                                            mPage.APLink = drAP.Cells[8].Value.ToString();
                                        break;
                                    }
                                }
                                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                                    {
                                        if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == mPage.APLink)
                                        {
                                            if (dr5.Cells[4].Value != null && dr5.Cells[4].Value.ToString() == "已啟用")
                                            {
                                                PageList.Add(mPage);
                                            }
                                            else
                                            {
                                                if (eslAPNoSetMsg == null)
                                                    eslAPNoSetMsg = dr.Cells[6].Value.ToString();
                                                if (eslAPNoSetMsg != null)
                                                    eslAPNoSetMsg = eslAPNoSetMsg + "," + dr.Cells[6].Value.ToString();
                                            }
                                        }
                                    }
                                Console.WriteLine("BBBBBBBBBB");
                            }
                        }
                    /*    else
                        {
                            if (nullMsg == null)
                            {
                                nullMsg = dr.Cells[6].Value.ToString();
                            }
                            else
                            {
                                nullMsg = nullMsg + "," + dr.Cells[6].Value.ToString();
                            }
                        }*/
                    }
                           
                        }
                    }

                    if (nullMsg != null)
                    {
                      //  MessageBox.Show("勾選" + nullMsg + "未綁定電子標籤");
                        PageList.Clear();
                        dataGridView1.Enabled = true;
                        nullMsg = null;
                        datagridview1curr = 2;
                        aaa(1, false, 0);
                    DialogResult dialogResult = MessageBox.Show("勾選" + nullMsg + "未綁定電子標籤" + "\r\n" + "是否繼續綁定", "未綁定", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }


                if (eslNotMateAP != null)
                {
                    //MessageBox.Show("勾選"+ nullMsg + "未綁定電子標籤");
                    PageList.Clear();
                    dataGridView1.Enabled = true;
                    nullMsg = null;
                    datagridview1curr = 2;
                    aaa(1, false, 0);
                    DialogResult dialogResult = MessageBox.Show(eslNotMateAP + "未配對AP請自動配對" + "\r\n" + "是否繼續執行", "未配對ESL", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }

                if (eslVState != null)
                {
                    //MessageBox.Show("勾選"+ nullMsg + "未綁定電子標籤");
                    PageList.Clear();
                    dataGridView1.Enabled = true;
                    nullMsg = null;
                    datagridview1curr = 2;
                    aaa(1, false, 0);
                    DialogResult dialogResult = MessageBox.Show(eslVState + "電壓未達2.85V" + "\r\n" + "是否繼續執行", "電壓", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                }

                if (eslAPNoSetMsg != null)
                {

                    //MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用");

                    dataGridView1.Enabled = true;
                    DialogResult dialogResult = MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用" + "\r\n" + "是否繼續執行", "AP未啟用", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        //do something
                    }
                    else if (dialogResult == DialogResult.No)
                    {
                        return;
                    }
                    else if (PageList.Count ==1)
                    {
                        return;
                    }
                    //return;
                }

                if (PageList.Count > 0)
                        {

                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                    {

                        kvp.Value.mSmcEsl.stopScanBleDevice();
                    }
                    EslStyleChangeUpdate = true;
                        int tt = 0;
                        testest = true;
                        onlockedbutton(testest);
                        List<string> RunAPList = new List<string>();
                        ProgressBarVisible(PageList.Count);
                        List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).ToList();
                        foreach (Page1 p in list)
                        {
                            RunAPList.Add(p.APLink);

                        }

                        stopwatch.Reset();
                        stopwatch.Start();


                for (int a = 0; a < RunAPList.Count; a++)
                {
                    for (int i = 0; i < PageList.Count; i++)
                    {
                        if (PageList[i].APLink == RunAPList[a])
                        {
                            Page1 mPage1 = PageList[i];
                            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                            {

                                if (kvp.Key.Contains(mPage1.APLink))
                                {

                                    Bitmap bmp;
                                    int Blcount = mPage1.BleAddress.Length;
                                    if (mPage1.onsale == "X")
                                    {
                                        bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                            mPage1.specification, mPage1.price, mPage1.Special_offer,
                                             mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                    }
                                    else if (mPage1.onsale == "V")
                                    {
                                        bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                             mPage1.specification, mPage1.price, mPage1.Special_offer,
                                              mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                    }
                                    else {
                                        bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                                    }
                                    int numVal = Convert.ToInt32(mPage1.no) - 1;
                                    Console.WriteLine("mPage1.no" + mPage1.no);
                                    dataGridView1.Rows[numVal].Cells[0].Selected = true;
                                    aaa(datagridview1curr, true, numVal);
                                    dataGridView1.Rows[numVal].Cells[17].Value = DateTime.Now.ToString();
                                    pictureBoxPage1.Image = bmp;

                                    Console.WriteLine("ININ");
                                        deviceIPData = mPage1.APLink;
                                       //  ConnectBleTimeOut.Start();
                                       
                                        // kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                        //   kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.BleAddress,0);
                                        //System.Threading.Thread.Sleep(1000);
                                        //      EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                        //    mSmcEsl.UpdataESLDataFromBuffer(mPage1.BleAddress, 0, 3,0);
                                     /*   mPage1.TimerConnect = new System.Windows.Forms.Timer();
                                        mPage1.TimerConnect.Interval = (30 * 1000);
                                        mPage1.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                        mPage1.TimerSeconds = new Stopwatch();*/
                                        mPage1.TimerSeconds.Start();
                                        mPage1.TimerConnect.Start();
                                        kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                                        richTextBox1.Text = richTextBox1.Text + PageList[i].usingAddress + "  嘗試連線中請稍候... \r\n";
                                }
                            }
                            break;
                        }
                    }
                }
                //  mSmcEsl.TransformImageToData(bmp);
                //  mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                macaddress = PageList[listcount].BleAddress;
                // richTextBox1.Text = mPage1.BleAddress + "  嘗試連線中請稍候... \r\n";

                }
            }
            else
            {
                MessageBox.Show("ESL更新中請稍後", "更新中");
            }
        }
        string textBeforeEdit;
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            Console.WriteLine("Big");
            if (e.ColumnIndex == 0)
            {
                if (!APStart)
                {
                    MessageBox.Show("請先連接AP");
                    dataGridView1.Enabled = true;
                    return;
                }

            }
            if (e.ColumnIndex != 3) { 
            if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == null)
                editdatagirdcell = "";
            else
            editdatagirdcell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            }
            if (e.ColumnIndex == 1)
            {
                
                OldEslList.Clear();
               
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {

                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                    {
                        if (dr.Cells[1].Value.ToString().Length > 12)
                        {
                            string[] esl = dr.Cells[1].Value.ToString().Split(',');
                            for (int i = 0; i < esl.Length; i++)
                            {
                                OldEslPage eslList = new OldEslPage();
                                eslList.ESLID = esl[i];
                                eslList.dataGridRowIndex = dr.Index;
                                OldEslList.Add(eslList);
                            }

                        }
                        else
                        {
                            OldEslPage eslList = new OldEslPage();
                            eslList.ESLID = dr.Cells[1].Value.ToString();
                            eslList.dataGridRowIndex = dr.Index;
                            OldEslList.Add(eslList);
                        }
                        
                    }
                }
            }
            if (e.ColumnIndex == 0)
            {

                textBeforeEdit = (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value ?? "").ToString();
                return;
            }

        }

        //商品狀態圖示
        private void productState(DataGridViewRow dr) {
            bool sale = false;
            bool beacon = false;

            if (dr.Cells[21].Value!=null && dr.Cells[21].Value.ToString() != "")
                beacon = true;

            if (dr.Cells[15].Value!=null && dr.Cells[15].Value.ToString() == "V")
                sale = true;


            string exeDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            if (beacon && sale)
            {
                Bitmap Image2 = new Bitmap(7, 7); //replace with your first image
                Image2 = (Bitmap)Image.FromFile(exeDir + @"\" + "beacon.png");
                Image2.MakeTransparent(Color.FromArgb(192, 255, 255));
                
                Image2.MakeTransparent(Color.White);
                Bitmap Image1 = new Bitmap(7, 7); //replace with your second image
                Image1 = (Bitmap)Image.FromFile(exeDir + @"\" + "sale.png");
                Image1.MakeTransparent(Color.White);
                Image1.MakeTransparent(Color.FromArgb(255, 128, 128));
                Console.WriteLine(Image1.Width + "," + Image2.Width);
                Console.WriteLine(Image1.Height + "," + Image2.Height);
                Bitmap ImageToDisplayInColumn = new Bitmap(Image1.Width + Image2.Width, Image2.Height);
                Console.WriteLine("PASS");
                using (Graphics graphicsObject = Graphics.FromImage(ImageToDisplayInColumn))
                {
                    graphicsObject.DrawImage(Image1, new Point(0, 0));
                    graphicsObject.DrawImage(Image2, new Point(Image1.Width, 0));
                }
                DataGridViewImageCell cell = dr.Cells[2] as DataGridViewImageCell;
                cell.Value = ImageToDisplayInColumn;
            }
            else if (beacon && !sale)
            {
                Console.WriteLine("beacon 小圖");
                Bitmap Image2 = new Bitmap(7, 7); //replace with your first image
                Image2 = (Bitmap)Image.FromFile(exeDir + @"\" + "beacon.png");
                Image2.MakeTransparent(Color.FromArgb(192, 255, 255));
                DataGridViewImageCell cell = dr.Cells[2] as DataGridViewImageCell;
                cell.Value = Image2;
            }
            else if (!beacon && sale)
            {
                Bitmap Image2 = new Bitmap(7, 7); //replace with your first image
                Image2 = (Bitmap)Image.FromFile(exeDir + @"\" + "sale.png");
                Image2.MakeTransparent(Color.White);
                Image2.MakeTransparent(Color.FromArgb(255, 128, 128));
                DataGridViewImageCell cell = dr.Cells[2] as DataGridViewImageCell;
                cell.Value = Image2;
            }
            else
            {
                DataGridViewImageCell cell = dr.Cells[2] as DataGridViewImageCell;
                cell.Value = DBNull.Value;
            }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void button20_Click(object sender, EventArgs e)
        {

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }
            if (label1.Text == "預設版型" || label1.Text == "特價版型")
            {
                MessageBox.Show("預設格式無法修改");
                return;
            }
            else
            {
                //  2/1
                bool ESLStyleCover = false;
                foreach (DataGridViewRow dr2 in this.dataGridView2.Rows)
                {
                    if (label1.Text == dr2.Cells[1].Value.ToString())
                    {
                        ESLStyleCover = true;
                        Console.WriteLine("ESLStyleCover");
                        if (dr2.Cells[2].Value.ToString() == "V")
                        {
                            ESLStyleDataChange = true;
                            button19.Enabled = true;
                            button19.BackColor = Color.FromArgb(255, 255, 192);
                        }
                    }
                }
                foreach (DataGridViewRow dr7 in this.dataGridView7.Rows)
                {
                    if (label1.Text == dr7.Cells[1].Value.ToString())
                    {
                        ESLStyleCover = true;
                        if (dr7.Cells[2].Value.ToString() == "V")
                        {
                            ESLSaleStyleDataChange = true;
                            button19.Enabled = true;
                            button19.BackColor = Color.FromArgb(255, 255, 192);
                        }
                    }
                }


                testest = true;
                // string add_PicBox=""+","+"";
                //   DataTable dt = dataGridView5.DataSource as DataTable;

                // dt.Rows.Add(new object[] { mAP_Information.AP_Name, mAP_Information.AP_IP, "8899" });
                //  foreach (Control x in pictureBox1.Controls)
                // {
                // mySheet.Rows[lastUsedRow].Add(x.Name, x.Width, x.Height, x.Location.X, x.Location.Y, x.Font);
                //   Console.WriteLine("lastUsedRow" + lastUsedRow);
                /* switch (x.Name)
                 {
                     case "ProName":
                         col = 1;
                         break;
                     case "ProBrand":
                         col = 7;
                         break;
                     case "ProFormat":
                         col = 13;
                         break;
                     case "ProPrice":
                         col = 19;
                         break;
                     case "ProPromotion":
                         col = 25;
                         break;
                     case "ProBarcode":
                         col = 31;
                         break;
                     case "ProESLID":
                         col = 37;
                         break;
                 //}*/
                // Console.WriteLine("col" + col + "lastUsedRow" + lastUsedRow + "x.Tag.ToString()" + x.Tag.ToString());
                /*   aaa.Add(x.Tag.ToString());
              add_PicBox = add_PicBox+"," +x.Tag.ToString();
                   aaa.Add(x.Name);
                   add_PicBox = add_PicBox + "," + x.Name;
                   aaa.Add(x.Text);
                   add_PicBox = add_PicBox + "," + x.Text;
                   aaa.Add(x.Width);
                   add_PicBox = add_PicBox + "," + x.Width;
                   aaa.Add(x.Height);
                   add_PicBox = add_PicBox + "," + x.Height;
                 add_PicBox = add_PicBox + "," + x.Location.X;
                 add_PicBox = add_PicBox + "," + x.Location.Y;
                 add_PicBox = add_PicBox + "," + x.Font.Name;
                 //    Console.WriteLine("Name" + x.Name + "width" + x.Width + x.Height + "textBox1.Location" + x.Location + "x.font" + x.Font + " x.ForeColor" + x.ForeColor.A + "," + x.ForeColor.R + "," + x.ForeColor.G + "," + x.ForeColor.B + "x.Font.Style" + x.Font.Style + "x.BackColor" + x.BackColor.A + "," + x.BackColor.R + "," + x.BackColor.G + "," + x.BackColor.B);
                 add_PicBox = add_PicBox + "," + x.Font.Size;
                 add_PicBox = add_PicBox + "," + x.Font.Style;
                 add_PicBox = add_PicBox + "," + x.ForeColor.A;
                 add_PicBox = add_PicBox + "," + x.ForeColor.R;
                 add_PicBox = add_PicBox + "," + x.ForeColor.G;
                 add_PicBox = add_PicBox + "," + x.ForeColor.B;
                 add_PicBox = add_PicBox + "," + x.BackColor.A;
                 add_PicBox = add_PicBox + "," + x.BackColor.R;
                 add_PicBox = add_PicBox + "," + x.BackColor.G;
                 add_PicBox = add_PicBox + "," + x.BackColor.B;

             }*/
                string filename = openExcelAddress;
                if (ESLStyleCover)
                {
                    Console.WriteLine("ESLStyleCover");
                    int size = 0;
                    if (pictureBox1.Height == 104)
                        size = 0;
                    else if (pictureBox1.Height == 128)
                        size = 1;
                    else if (pictureBox1.Height == 300)
                        size = 2;
                    mExcelData.ESLStyleCover(label1.Text, pictureBox1, excel, excelwb, mySheet,size);
                }
                else
                {
                    if (ESLStyleSave)
                    {
                        Console.WriteLine("ESLStyleSave1");
                        int size = 0;
                        if (pictureBox1.Height == 104)
                            size = 0;
                        else if (pictureBox1.Height == 128)
                            size = 1;
                        else if (pictureBox1.Height == 300)
                            size = 2;


                        mExcelData.dataGridView2Update(dataGridView2, label1.Text, filename, pictureBox1, excel, excelwb, mySheet, 0,size);
                    }
                    if (ESLSaleStyleSave)
                    {
                        Console.WriteLine("ESLStyleSave2");
                        int size = 0;
                        if (pictureBox1.Height == 104)
                            size = 0;
                        else if (pictureBox1.Height == 128)
                            size = 1;
                        else if (pictureBox1.Height == 300)
                            size = 2;

                        mExcelData.dataGridView2Update(dataGridView7, label1.Text, filename, pictureBox1, excel, excelwb, mySheet, 1,size);
                    }
                }

                dataGridView2.Columns.Clear();
                dataGridView2.DataSource = d;
                DataGridViewColumn dgvc2 = new DataGridViewCheckBoxColumn();
                dgvc2.Width = 60;
                dgvc2.Name = "選取";
                dgvc2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                this.dataGridView2.Columns.Insert(0, dgvc2);
                string tableName2 = "[工作表2$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql2 = "select * from " + tableName2 + "WHERE 版型類型=0";//SQL查詢
                DataTable kk2 = mExcelData.GetExcelDataTable(filename, sql2);
                dataGridView2.DataSource = kk2;

                dataGridView7.Columns.Clear();
                dataGridView7.DataSource = d;
                DataGridViewColumn dgvc7 = new DataGridViewCheckBoxColumn();
                dgvc7.Width = 60;
                dgvc7.Name = "選取";
                dgvc7.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                this.dataGridView7.Columns.Insert(0, dgvc7);
                string tableName7 = "[工作表2$]";//在頁簽名稱後加$，再用中括號[]包起來
                string sql7 = "select * from " + tableName7 + "WHERE 版型類型=1";//SQL查詢
                DataTable kk7 = mExcelData.GetExcelDataTable(filename, sql7);
                dataGridView7.DataSource = kk7;

           
                for (int ee = 0; ee < this.dataGridView7.ColumnCount; ee++)
                {
                    if (ee != 0 && ee != 1 && ee != 2)
                        this.dataGridView7.Columns[ee].Visible = false;
                }

                for (int ee = 0; ee < this.dataGridView2.ColumnCount; ee++)
                {
                    if (ee != 0 && ee != 1 && ee != 2)
                        this.dataGridView2.Columns[ee].Visible = false;
                }


                this.dataGridView2.Columns[1].ReadOnly = true;
                this.dataGridView2.Columns[2].ReadOnly = true;
                this.dataGridView2.Columns[0].Width = 20;
                this.dataGridView2.Columns[1].Width = 79;
                this.dataGridView2.Columns[2].Width = 20;
                this.dataGridView7.Columns[1].ReadOnly = true;
                this.dataGridView7.Columns[2].ReadOnly = true;
                this.dataGridView7.Columns[0].Width = 20;
                this.dataGridView7.Columns[1].Width = 79;
                this.dataGridView7.Columns[2].Width = 20;

                excel = new Excel.Application();
                excelwb = excel.Workbooks.Open(openExcelAddress);
                // excel.Application.Workbooks.Add(true);
                mySheet = new Excel.Worksheet();
                //excel.Visible = false;
                //excel.Quit();//離開聯結 
                if (ESLSaleStyleDataChange)
                {
                    ESLSaleFormat.Clear();
                    foreach (DataGridViewRow dr77 in this.dataGridView7.Rows)
                    {
                        if (dr77.Cells[2].Value != null && dr77.Cells[2].Value.ToString() == "V")
                        {

                            for (int i = 1; i < dr77.Cells.Count; i++)
                            {
                               
                                    //           Console.WriteLine("HGEEGE");
                                    if (i == 1)
                                    {
                                        styleSaleName = dr77.Cells[1].Value.ToString();
                                    }
                                    if (i != 0 && i != 1 && i != 2)
                                    {
                                    if (dr77.Cells[i].Value != null && dr77.Cells[i].Value.ToString() != "")
                                    {
                                        ESLSaleFormat.Add(dr77.Cells[i].Value.ToString());
                                    }
                                    else
                                        {
                                            if (i < dataGridView7.ColumnCount)
                                                if (dr77.Cells[i - 1].Value.ToString() != "")
                                                    ESLSaleFormat.Add(dr77.Cells[i].Value.ToString());
                                        }
                                        //            Console.WriteLine(dr.Cells[i].Value.ToString());
                                    }

                            }

                        }
                    }
                }
                if (ESLStyleDataChange)
                {
                    ESLFormat.Clear();
                    foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                    {


                        if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V")
                        {
                            for (int i = 1; i < dr.Cells.Count; i++)
                            {
                               
                                    if (i == 1)
                                    {
                                        styleName = dr.Cells[1].Value.ToString();
                                    }
                                    if (i != 0 && i != 1 && i != 2)
                                    {
                                    if (dr.Cells[i].Value != null && dr.Cells[i].Value.ToString() != "")
                                    {
                                        ESLFormat.Add(dr.Cells[i].Value.ToString());
                                    }
                                    else
                                        {
                                            if (i < dataGridView2.ColumnCount)
                                                if (dr.Cells[i - 1].Value.ToString() != "")
                                                ESLFormat.Add(dr.Cells[i].Value.ToString());
                                        }
                                    }

                            }

                        }
                    }
                }

                testest = false;
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }




            List<string> ESLTYPE = new List<string>();
            ESLTYPE.Add("0");
              foreach (Control x in pictureBox1.Controls)
              {

                  ESLTYPE.Add(x.Tag.ToString());
                  ESLTYPE.Add(x.Name.ToString());
                  ESLTYPE.Add(x.Text.ToString());
                  ESLTYPE.Add(x.Width.ToString());
                  ESLTYPE.Add(x.Height.ToString());
                  ESLTYPE.Add(x.Location.X.ToString());
                  ESLTYPE.Add(x.Location.Y.ToString());
                  ESLTYPE.Add(x.Font.Name.ToString());
                  ESLTYPE.Add(x.Font.Size.ToString());

                Console.WriteLine(x.Name.ToString() + ":" + x.Font.Style.ToString());
                string FontStyle=null;
                /*if (x.Font.Style.ToString().Contains("Regular"))
                {
                    if (FontStyle == null)
                        FontStyle = "Regular";
                    else
                        FontStyle = FontStyle + ",Regular";
                }
                else if (x.Font.Style.ToString() == "Bold")
                    ESLTYPE.Add("1");
                else if (x.Font.Style.ToString() == "Italic")
                    ESLTYPE.Add("2");
                else if (x.Font.Style.ToString() == "Underline")
                    ESLTYPE.Add("4");
                else if (x.Font.Style.ToString() == "Strikeout")
                    ESLTYPE.Add("8");
                else if (x.Font.Style.ToString() == "Bold")
                    ESLTYPE.Add("3");*/
                  ESLTYPE.Add(x.Font.Style.ToString());

                  ESLTYPE.Add(x.ForeColor.A.ToString());
                  ESLTYPE.Add(x.ForeColor.R.ToString());
                  ESLTYPE.Add(x.ForeColor.G.ToString());
                  ESLTYPE.Add(x.ForeColor.B.ToString());
                  ESLTYPE.Add(x.BackColor.A.ToString());
                  ESLTYPE.Add(x.BackColor.R.ToString());
                  ESLTYPE.Add(x.BackColor.G.ToString());
                  ESLTYPE.Add(x.BackColor.B.ToString());


              }


          /*  foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                Console.WriteLine("label1.Text"+ label1.Text);
                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == label1.Text)
                {

                    Console.WriteLine("HGEEGE" + dr.Cells[1].Value.ToString());
                    for (int i = 0; i < dr.Cells.Count; i++)
                    {
                        if (i != 2&& i != 0)
                        {
                            if (dr.Cells[i].Value != null && dr.Cells[i].Value.ToString() != "")
                            {
                                Console.WriteLine("HGEEGE");
                                if (i == 1)
                                {
                                    styleName = dr.Cells[1].Value.ToString();
                                }
                                if (i != 0 && i != 1 && i != 2)
                                {

                                    ESLTYPE.Add(dr.Cells[i].Value.ToString());

                                    Console.WriteLine(dr.Cells[i].Value.ToString());
                                }
                            }
                            else
                            {
                                break;
                            }
                        }

                    }

                }
            }

            foreach (DataGridViewRow dr in this.dataGridView7.Rows)
            {
                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == label1.Text)
                {

                    for (int i = 0; i < dr.Cells.Count; i++)
                    {
                        if (i != 2 && i != 0)
                        {
                            if (dr.Cells[i].Value != null && dr.Cells[i].Value.ToString() != "")
                        {
                            //     Console.WriteLine("HGEEGE");

                            if (i == 1)
                            {
                                styleName = dr.Cells[1].Value.ToString();
                            }
                            if (i != 0 && i != 1 && i != 2)
                            {

                                ESLTYPE.Add(dr.Cells[i].Value.ToString());

                                //       Console.WriteLine(dr.Cells[i].Value.ToString());
                            }
                        }
                        else
                        {
                            break;
                        }
                        }
                    }

                }
            }*/
            
            if (ESLTYPE.Count == 0)
            {
                MessageBox.Show("該版型未新增");
                return;
            }

            Bitmap bmp = mElectronicPriceData.setPage1("Calibri", "綜合B群加強錠", "統一股份有限公司",
                                             "90錠入","450", "400",
                                              "4444444444444", "http://www.smartchip.com.tw", "9ADF12356891","隨便", ESLTYPE);
            pictureBoxPage1.Image = bmp;

        }

        private void dataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            panel2.Visible=true;
            panel3.Visible=false;
            button10.BackColor = Color.FromArgb(255, 255, 192);
            button22.BackColor = Color.Gray;
           
        }

        private void button22_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
            panel3.Visible = true;
            button10.BackColor = Color.Gray;
            button22.BackColor = Color.FromArgb(255, 255, 192);

        }

        private void button23_Click(object sender, EventArgs e)
        {
            string ESL=null;
            string Oldcode=null;
            for (int i = backESLList.Count-1; i >= 0; i--)
            {
                if (!backESLList[i].isBack)
                {
                    ESL = backESLList[i].NewMateESL;
                    Oldcode = backESLList[i].OldMateProduct;
                    backESLList[i].isBack = true;
                    break;
                }
            }
            if (ESL == null)
                return;

            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {

                if (dr.Cells[1].Value.ToString().Contains(',' + ESL))
                {
                    int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + ESL);
                    dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                    break;
                }
                if (dr.Cells[1].Value.ToString().Contains(ESL + ','))
                {

                    int changeaddr = dr.Cells[1].Value.ToString().IndexOf(ESL + ',');
                    dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                    break;
                }

                if (dr.Cells[1].Value.ToString().Contains(ESL))
                {

                    int changeaddr = dr.Cells[1].Value.ToString().IndexOf(ESL);
                    dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                    break;
                }

            }

            if (Oldcode != null)
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    if (dr.Cells[5].Value != null && Oldcode == dr.Cells[5].Value.ToString())
                    {
                        if (dr.Cells[1].Value.ToString().Length > 1)
                        {
                            dr.Cells[1].Value = dataGridView1[1, rowIndex].Value + "," + ESL;
                            if (dr.Cells[12].Value != null && dr.Cells[1].Value.ToString() == dr.Cells[12].Value.ToString())
                            {
                                dr.Cells[1].Style.ForeColor = Color.Black;
                            }
                            else
                            {
                                dr.Cells[1].Style.ForeColor = Color.Orange;
                            }
                            
                            //dataGridView1[2, rowIndex].Value = "已綁定";
                        }
                        else
                        {
                            dr.Cells[1].Value = ESL;

                            if (dr.Cells[12].Value != null && dr.Cells[1].Value.ToString() == dr.Cells[12].Value.ToString())
                            {
                                dr.Cells[1].Style.ForeColor = Color.Black;
                            }
                            else
                            {
                                dr.Cells[1].Style.ForeColor = Color.Orange;
                            }
                            //dataGridView1[2, rowIndex].Value = "已綁定";
                        }
                        break;
                    }
                }
            }



        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {
                    }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            if (dataGridView1.RowCount == 0)
            {
                MessageBox.Show("請先匯入EXCEL");
                dataGridView1.Enabled = true;
                return;
            }




            if (!testest)
            {
                int relrowno = 0;
                string pp = "";
                List<DataGridViewRow> toDelete = new List<DataGridViewRow>();
                List<int> deldataview4no = new List<int>();

                foreach (DataGridViewRow dr in this.dataGridView5.Rows)
                {
                    if (dr.Cells[0].Value != null && dr.Cells[0].Value.ToString() == "True")
                    {

                        deldataview4no.Add(dr.Cells[0].RowIndex + 2 - relrowno);
                        relrowno++;
                        toDelete.Add(dr);
                    }
                }

                if (deldataview4no != null)
                {
                    DialogResult result = MessageBox.Show("該欄位是否刪除?", "刪除", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {



                        foreach (DataGridViewRow row in toDelete)
                        {

                            /* foreach (DataGridViewRow dr1 in this.dataGridView1.Rows)
                             {
                                 if (row.Cells[1].Value == dr1.Cells[2].Value) {

                                 }
                             }*/

                            dataGridView5.Rows.Remove(row);
                            CountESLAll.Text = (Convert.ToInt32(CountESLAll.Text) - 1).ToString();
                        }
                        mExcelData.dataviewdel(dataGridView5, deldataview4no, "工作表4", openExcelAddress, excel, excelwb, mySheet);

                    }
                }
                else
                {
                    MessageBox.Show("未選取選項", "刪除");
                    
                }
            }
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }


        private void onlockedbutton(bool locked)
        {
            if (locked)
            {
                SendData.Enabled = false;
                button2.Enabled = false;
                button15.Enabled = false;
                button19.Enabled = false;
                dataGridView1.Enabled = false;
            }
            else
            {
                SendData.Enabled = true;
                button2.Enabled = true;
                button15.Enabled = true;
                button19.Enabled = true;
                dataGridView1.Enabled = true;
            }
        }




        private void immediateESLUpdate(string eslID)
        {
            // DataGridViewRow drR = dataGridView1.Rows[e.RowIndex];

            Page1 mPage = new Page1();
            // drR.Selected = false;
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                //  MessageBox.Show(" " + ((DataRowView)(dr.DataBoundItem)).Row.ItemArray[0] + " 被選取了！");
                //字體   品名  品牌  規格  價格  特價  條碼  Qr
                Console.WriteLine("AAAAAAAAAA");
                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString().Contains(eslID))
                {
                    mPage.no = (dr.Index + 1).ToString();
                    UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                    mPage.BleAddress = eslID;
                    mPage.barcode = dr.Cells[5].Value.ToString();
                    mPage.product_name = dr.Cells[6].Value.ToString();
                    mPage.Brand = dr.Cells[7].Value.ToString();
                    mPage.specification = dr.Cells[8].Value.ToString();
                    mPage.price = dr.Cells[9].Value.ToString();
                    mPage.Special_offer = dr.Cells[10].Value.ToString();
                    mPage.Web = dr.Cells[11].Value.ToString();
                    mPage.HeadertextALL = headertextall;
                    mPage.usingAddress = eslID;
                    mPage.onsale = dr.Cells[15].Value.ToString();
                    mPage.actionName = "immediateUpdate";
                    if (mPage.onsale == "V")
                        mPage.ProductStyle =styleSaleName;
                    else
                        mPage.ProductStyle = styleName;
                    foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                        {
                            if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == eslID)
                            {
                                mPage.APLink = drAP.Cells[8].Value.ToString();
                                break;
                            }
                        }
                    dr.Cells[1].Value.ToString();
                   // dr.Cells[16].Value = "X";
                    PageList.Add(mPage);
                    Console.WriteLine("BBBBBBBBBB");
                    //     Console.WriteLine("checkClick"+ checkClick);
                    if (!checkClick&& mPage.APLink!=null)
                    {
                        listcount = 0;
                        checkClick = true;
                        immediateUpdate = true;
                        testest = true;
                        stopwatch.Reset();
                        stopwatch.Start();
                    }
                    List<string> RunAPList = new List<string>();
                    List<string> newAPList = new List<string>();
                    List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).Where(p => p.UpdateState == null).ToList();
                    foreach (Page1 p in list)
                    {
                        RunAPList.Add(p.APLink);
                        Console.WriteLine("RunAPList" + p.APLink + p.BleAddress);
                    }

                    for (int a = 0; a < OldRunAPList.Count; a++)
                    {
                        Console.WriteLine("OldRunAPList" + OldRunAPList[a]);

                    }

                    Console.WriteLine("RunAPList.Except(OldRunAPList).Count()" + RunAPList.Except(OldRunAPList).Count());
                    if (RunAPList.Count == OldRunAPList.Count && RunAPList.Except(OldRunAPList).Count() == 0)
                    {
                        Console.WriteLine("一樣");
                    }
                    else
                    {
                        newAPList = RunAPList.Except(OldRunAPList).ToList();
                        Console.WriteLine("我們不一樣" + newAPList);


                        for (int a = 0; a < newAPList.Count; a++)
                        {
                            Console.WriteLine("newAPList" + newAPList[a]);
                            for (int i = 0; i < PageList.Count; i++)
                            {
                                if (PageList[i].APLink == newAPList[a])
                                {
                                    Page1 mPage1 = PageList[i];
                                    
                                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                                    {

                                        if (mPage1.APLink!=null&&kvp.Key.Contains(mPage1.APLink))
                                        {
                                            Bitmap bmp;
                                            if (mPage1.onsale == "V")
                                            {
                                                 bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                     mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                            }
                                            
                                            else
                                            {
                                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                     mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                            }
                                            //  Bitmap bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                                            

                                            /*  Bitmap bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                                        mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                         mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);*/
                                            int numVal = Convert.ToInt32(mPage1.no) - 1;
                                            //  Console.WriteLine("mPage1.no" + mPage1.no);
                                            dataGridView1.Rows[numVal].Cells[0].Selected = true;
                                            aaa(datagridview1curr, true, numVal);
                                            dataGridView1.Rows[numVal].Cells[17].Value = DateTime.Now.ToString();
                                            pictureBoxPage1.Image = bmp;

                                            //  Console.WriteLine("ININ");
                                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                            kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.BleAddress,0);
                                            //System.Threading.Thread.Sleep(1000);
                                            EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                            mSmcEsl.UpdataESLDataFromBuffer(mPage1.BleAddress, 0, 3,0);
                                            pictureBoxPage1.Image = bmp;
                                            richTextBox1.Text = richTextBox1.Text + PageList[i].usingAddress + "  嘗試連線中請稍候... \r\n";
                                            System.Threading.Thread.Sleep(1000);
                                        }
                                    }
                                    break;
                                }
                            }

                        }

                    }

                    /*     if (eslAPNoSetMsg != null)
                         {
                             // tt = 1;
                             // MessageBox.Show(eslAPNoSetMsg + "該ESL綁定AP未啟用");

                             dataGridView1.Enabled = true;
                             DialogResult dialogResult = MessageBox.Show("AP未启用", eslAPNoSetMsg + "該ESL綁定AP未啟用" + "\r\n" + "是否繼續執行", MessageBoxButtons.YesNo);
                             if (dialogResult == DialogResult.Yes)
                             {
                                 //do something
                             }
                             else if (dialogResult == DialogResult.No)
                             {
                                 return;
                             }
                             //return;
                         }*/

                    //   Console.WriteLine("OldRunAPList:" + OldRunAPList.Count + "RunAPList:" + RunAPList.Count);
                    OldRunAPList = RunAPList;
                }
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                Console.WriteLine("textBox2.Text" + textBox2.Text);

                string text=textBox2.Text.Trim();
                if (text != "")
                {
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {



                        if (dr.Cells[6].RowIndex == dataGridView1.Rows.Count - 1)
                        {
                            break;
                        }
                        if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString().Contains(text))
                        {
                            dr.Visible = true;
                        }
                        else
                        {
                            dataGridView1.CurrentCell = null;
                            dr.Visible = false;
                        }

                        //  textBox2.Text = "";
                    }

                }
                else
                {
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {
                            dr.Visible = true;
                    }
                }
                textBox2.Text =null;
                return;
            }
           
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
             

                string text = textBox3.Text.Trim();
                if (text != "")
                {
                    foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                    {



                        if (dr.Cells[1].RowIndex == dataGridView4.Rows.Count - 1)
                        {
                            break;
                        }
                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString().Contains(text))
                        {
                            dr.Visible = true;
                        }
                        else
                        {
                            dataGridView4.CurrentCell = null;
                            dr.Visible = false;
                        }

                        //  textBox2.Text = "";
                    }

                }
                else
                {
                    foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                    {
                        dr.Visible = true;
                    }
                }
                textBox3.Text = null;
                return;
            }
        }

        private void propertyGrid1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void tabControl1_TabIndexChanged(object sender, EventArgs e)
        {
            pictureBoxPage1.Image = null;
            richTextBox1.Text = "";
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            pictureBoxPage1.Image = null;
            richTextBox1.Text = "";
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
            {

                kvp.Value.mSmcEsl.DisConnectBleDevice();

            }

            if (down) {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    //  Console.WriteLine("----------------------------------------");
                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                    {
                        dr.Cells[0].Value = true;
                        dr.DefaultCellStyle.ForeColor = Color.Black;
                        dr.ReadOnly = false;
                    }
                }
            }

            if (sale)
            {
                if (testest)
                {
                    button2.Enabled = true;
                }
                else
                {
                    button2.Enabled = false;
                }
            }

           
            testest = false;
            onlockedbutton(testest);
            checkClick = false;
            down = false;
            OldRunAPList.Clear();
            backESLList.Clear();
            Console.WriteLine("87877878787");
            progressBar1.Visible = false;
            down = false;
            sale = false;
            reset = false;
            saletime = false;
            immediateUpdate = false;
            listcount = 0;
            EslStyleChangeUpdate = false;
            
            pictureBoxPage1.Image = null;
            richTextBox1.Text = "";
            stopwatch.Stop();//碼錶停止
            TimeSpan ts = stopwatch.Elapsed;
            string elapsedTime = String.Format("{0:00} 分 {1:00} 秒 {2:000} ms",
            ts.Minutes, ts.Seconds,
            ts.Milliseconds);
            dataGridView1.Enabled = true;
            removeESLingstate = false;
            onsaleESLingstate = false;
            updateESLingstate = false;
            ESLStyleDataChange = false;
            ESLSaleStyleDataChange = false;
            PageList.Clear();
            testest = false;

            //pictureBox4.Visible = false;
            //   checkClick = false;
            //  OldRunAPList.RemoveAll(it => true);
            OldRunAPList.Clear();
            // mSmcEsl.DisConnectBleDevice();
            //ConnectTimer.Stop();
            CheckBeaconTimer.Start();
        }

        private void eslonsalecheck() {
            //SALE ALL CHECK===========================================
            // PageList.Clear();
            //  listcount = 0;
            string saletimemsg="";
            SalePageListUpdate.Clear();
            Console.WriteLine("BeaconPP" + PageList.Count);
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                if (dr.Cells[19].Value != null && dr.Cells[20].Value != null && dr.Cells[19].Value.ToString() != "" && dr.Cells[20].Value.ToString() != "")
                {
                    Console.WriteLine("199119" + dr.Cells[6].Value.ToString());
                    Console.WriteLine("199119" + dr.Cells[19].Value.ToString());
                    string format = "yyyy/MM/dd HH:mm:ss";
                    string start = Convert.ToDateTime(dr.Cells[19].Value).ToString("yyyy/MM/dd HH:mm:ss");
                    string end = Convert.ToDateTime(dr.Cells[20].Value).ToString("yyyy/MM/dd HH:mm:ss");
                    DateTime strDate = DateTime.ParseExact(start, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                    DateTime endDate = DateTime.ParseExact(end, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                    if (dr.Cells[1].Value.ToString().Length > 13)
                    {

                        string[] usingAddressSplit = dr.Cells[1].Value.ToString().Split(',');
                        string Special_offer;
                        string onSale = "";
                        string updateStyle = "";
                        if (DateTime.Compare(strDate, DateTime.Now) < 0 && DateTime.Compare(endDate, DateTime.Now) > 0)
                        {

                            dr.Cells[15].Value = onSale = "V";
                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                            Special_offer = dr.Cells[10].Value.ToString();
                            updateStyle = styleSaleName;
                        }
                        else
                        {
                            dr.Cells[15].Value = onSale = "X";
                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                            Special_offer = dr.Cells[10].Value.ToString();
                            updateStyle = styleName;
                        }
                        for (int i = 0; i < usingAddressSplit.Length; i++)
                        {
                            Page1 mPageA = new Page1();
                            mPageA.no = (dr.Index + 1).ToString();
                            mPageA.BleAddress = usingAddressSplit[i];
                            mPageA.barcode = dr.Cells[5].Value.ToString();
                            mPageA.product_name = dr.Cells[6].Value.ToString();
                            mPageA.Brand = dr.Cells[7].Value.ToString();
                            mPageA.specification = dr.Cells[8].Value.ToString();
                            mPageA.price = dr.Cells[9].Value.ToString();
                            mPageA.Web = dr.Cells[11].Value.ToString();
                            mPageA.usingAddress = usingAddressSplit[i];
                            mPageA.HeadertextALL = headertextall;
                            mPageA.Special_offer = dr.Cells[10].Value.ToString();
                            mPageA.onsale = onSale;
                            mPageA.onSaleTimeS = dr.Cells[19].Value.ToString();
                            mPageA.ProductStyle = updateStyle;
                            mPageA.onSaleTimeE = dr.Cells[20].Value.ToString();
                            mPageA.actionName = "saletime";
                            Console.WriteLine("usingAddressSplit[i] mPageA.ProductStyle " + mPageA.ProductStyle);
                            foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                            {
                                if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == usingAddressSplit[i])
                                {
                                    mPageA.APLink = drAP.Cells[8].Value.ToString();
                                    break;
                                }
                            }

                            //  PageList.Add(mPageA);
                            SalePageListUpdate.Add(mPageA);

                        }
                    }
                    else
                    {
                        Console.WriteLine("dr.Cells[6].Value.ToString()" + dr.Cells[6].Value.ToString());
                        Page1 mPageC = new Page1();
                        mPageC.no = (dr.Index + 1).ToString();
                        mPageC.BleAddress = dr.Cells[1].Value.ToString();
                        mPageC.barcode = dr.Cells[5].Value.ToString();
                        mPageC.product_name = dr.Cells[6].Value.ToString();
                        mPageC.Brand = dr.Cells[7].Value.ToString();
                        mPageC.specification = dr.Cells[8].Value.ToString();
                        mPageC.price = dr.Cells[9].Value.ToString();

                        mPageC.Web = dr.Cells[11].Value.ToString();
                        mPageC.usingAddress = dr.Cells[1].Value.ToString();
                        mPageC.HeadertextALL = headertextall;
                        //mPageC.Special_offer = dr.Cells[10].Value.ToString();
                        string updateStyle;
                        if (DateTime.Compare(strDate, DateTime.Now) < 0 && DateTime.Compare(endDate, DateTime.Now) > 0)
                        {
                            dr.Cells[15].Value = mPageC.onsale = "V";
                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                            mPageC.Special_offer = dr.Cells[10].Value.ToString();
                            updateStyle = styleSaleName;
                        }
                        else
                        {
                            dr.Cells[15].Value = mPageC.onsale = "X";
                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                            mPageC.Special_offer = dr.Cells[10].Value.ToString();
                            updateStyle = styleName;
                        }
                        mPageC.ProductStyle = updateStyle;
                        mPageC.onsale = dr.Cells[15].Value.ToString();
                        mPageC.onSaleTimeS = dr.Cells[19].Value.ToString();
                        mPageC.onSaleTimeE = dr.Cells[20].Value.ToString();
                        mPageC.actionName = "saletime";
                        Console.WriteLine("dr.Cells[1].Value.ToString() mPageC.ProductStyle " + mPageC.ProductStyle);
                        foreach (DataGridViewRow drAP in this.dataGridView4.Rows)
                        {
                            if (drAP.Cells[1].Value != null && drAP.Cells[1].Value.ToString() == dr.Cells[1].Value.ToString())
                            {
                                mPageC.APLink = drAP.Cells[8].Value.ToString();
                                break;
                            }
                        }

                        SalePageListUpdate.Add(mPageC);
                    }
                }
            }



            // ------------------明天 初始修改

            if (SalePageListUpdate.Count != 0)
            {

                List<string> RunAPList = new List<string>();
                List<Page1> list = SalePageListUpdate.GroupBy(a => a.APLink).Select(g => g.First()).ToList();
                foreach (Page1 p in list)
                {
                    RunAPList.Add(p.APLink);

                }

                foreach (DataGridViewRow dr5 in this.dataGridView5.Rows)
                {
                    if (dr5.Cells[2].Value != null && dr5.Cells[2].Value.ToString() == SalePageListUpdate[0].APLink)
                    {
                        if (dr5.Cells[4].Value.ToString() == "")
                        {
                            if (down || sale)
                            {

                                MessageBox.Show(SalePageListUpdate[0].usingAddress + "該ESL綁定AP未啟用");
                            }
                            else
                            {

                                MessageBox.Show(SalePageListUpdate[0].BleAddress + "該ESL綁定AP未啟用");
                            }
                            return;
                        }
                    }

                }


                //   sale = true;

                Boolean assalepage = true;
                //mSmcEsl.TransformImageToData(bmp);
                Console.WriteLine("Count" + SalePageListUpdate.Count + "k" + SalePageList.Count + "p" + PageList.Count);

                if (SalePageListUpdate.Count != SalePageList.Count)
                {
                    assalepage = false;
                }
                else
                {
                    if (SalePageList.Count != 0)
                    {
                        for (int i = 0; i < SalePageListUpdate.Count; i++)
                        {

                            if (SalePageListUpdate[i].onsale != SalePageList[i].onsale || SalePageListUpdate[i].onSaleTimeS != SalePageList[i].onSaleTimeS || SalePageListUpdate[i].onSaleTimeE != SalePageList[i].onSaleTimeE || SalePageListUpdate[i].product_name != SalePageList[i].product_name || SalePageListUpdate[i].barcode != SalePageList[i].barcode || SalePageListUpdate[i].price != SalePageList[i].price || SalePageListUpdate[i].Special_offer != SalePageList[i].Special_offer || SalePageListUpdate[i].specification != SalePageList[i].specification)
                            {
                                Console.WriteLine("LOOK" + SalePageListUpdate + "and" + SalePageList);
                                assalepage = false;
                            }
                        }
                    }
                }
                if (!assalepage)
                {
                    testest = true;
                    onlockedbutton(testest);
                    saletime = true;
                    /*for (int i = 0; i < SalePageListUpdate.Count; i++)
                    {
                        PageList.Add(SalePageListUpdate[i]);
                    }*/



                    //SalePageList.Clear();
                    /*   for (int i = 0; i < SalePageListUpdate.Count; i++)
                       {
                           Page1 mPageA = new Page1();
                           mPageA.no = SalePageListUpdate[i].no;
                           mPageA.BleAddress = SalePageListUpdate[i].BleAddress;
                           mPageA.barcode = SalePageListUpdate[i].barcode;
                           mPageA.product_name = SalePageListUpdate[i].product_name;
                           mPageA.Brand = SalePageListUpdate[i].Brand;
                           mPageA.specification = SalePageListUpdate[i].specification;
                           mPageA.price = SalePageListUpdate[i].price;
                           mPageA.Web = SalePageListUpdate[i].Web;
                           mPageA.usingAddress = SalePageListUpdate[i].usingAddress;
                           mPageA.HeadertextALL = SalePageListUpdate[i].HeadertextALL;
                           mPageA.Special_offer = SalePageListUpdate[i].Special_offer;
                           mPageA.onsale = SalePageListUpdate[i].onsale;
                           mPageA.onSaleTimeS = SalePageListUpdate[i].onSaleTimeS;
                           mPageA.ProductStyle = SalePageListUpdate[i].ProductStyle;
                           mPageA.onSaleTimeE = SalePageListUpdate[i].onSaleTimeE;
                           mPageA.actionName = SalePageListUpdate[i].actionName;
                           mPageA.APLink = SalePageListUpdate[i].APLink;
                           PageList.Add(mPageA);
                           SalePageList.Add(mPageA);
                       }*/

                    //PageList.AddRange(SalePageListUpdate);
                    PageList = SalePageListUpdate;
                    listcount = 0;
                    stopwatch.Reset();
                    stopwatch.Start();
                    Console.WriteLine("不依樣近來更新");
                    for (int a = 0; a < RunAPList.Count; a++)
                    {
                        for (int i = 0; i < PageList.Count; i++)
                        {
                            Console.WriteLine("(PageList[i]" + PageList[i].APLink + "RunAPList[a]" + RunAPList[a]);
                            if (PageList[i].APLink == RunAPList[a])
                            {
                                Console.WriteLine("PageList[i].APLink" + PageList[i].APLink);
                                Page1 mPage1 = PageList[i];
                                if (mPage1.usingAddress != "")
                                {
                                    // int Blcount = mPage1.BleAddress.Length;
                                    string format = "yyyy/MM/dd HH:mm:ss";
                                    string starta = Convert.ToDateTime(mPage1.onSaleTimeS).ToString("yyyy/MM/dd HH:mm:ss");
                                    string enda = Convert.ToDateTime(mPage1.onSaleTimeE).ToString("yyyy/MM/dd HH:mm:ss");
                                    DateTime strDatea = DateTime.ParseExact(starta, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                                    DateTime endDatea = DateTime.ParseExact(enda, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                                    Bitmap bmp;
                                    if (DateTime.Compare(strDatea, DateTime.Now) < 0 && DateTime.Compare(endDatea, DateTime.Now) > 0)
                                    {
                                        bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                        mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);

                                    }
                                    else
                                    {
                                        bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                        mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                        //CheckBeaconTimer.Stop();
                                        saletimemsg = saletimemsg + PageList[i].product_name + "特價已到期" + PageList[i].onSaleTimeS + "-" + PageList[i].onSaleTimeE + "\r\n";
                                        //if (result == DialogResult.OK)
                                        //{
                                        // Do something
                                        //1/31
                                        foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                                        {
                                            if (dr.Cells[5].Value != null && PageList[i].barcode == dr.Cells[5].Value.ToString())
                                            {
                                                dr.Cells[19].Value = DBNull.Value;
                                                dr.Cells[20].Value = DBNull.Value;
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            }
                                        }
                                        //CheckBeaconTimer.Start();
                                        //}
                                    }


                                    foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                                    {

                                        if (kvp.Key.Contains(mPage1.APLink))
                                        {



                                            int numVal = Convert.ToInt32(mPage1.no) - 1;
                                            Console.WriteLine("mPage1.no" + mPage1.no);
                                            Console.WriteLine("mPage1.APLink" + mPage1.APLink);
                                            //dataGridView1.ClearSelection();
                                            dataGridView1.Rows[numVal].Cells[0].Selected = true;
                                            aaa(datagridview1curr, true, numVal);
                                            dataGridView1.Rows[numVal].Cells[17].Value = DateTime.Now.ToString();
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, numVal, false, openExcelAddress, excel, excelwb, mySheet);
                                            pictureBoxPage1.Image = bmp;
                                            mPage1.TimerConnect = new System.Windows.Forms.Timer();
                                            mPage1.TimerConnect.Interval = (30 * 1000);
                                            mPage1.TimerConnect.Tick += new EventHandler(ConnectBle_TimeOut);
                                            mPage1.TimerSeconds = new Stopwatch();
                                            mPage1.TimerSeconds.Start();
                                            mPage1.TimerConnect.Start();

                                            kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.BleAddress);
                                            /*   Console.WriteLine("ININ");
                                               kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                               kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress,0);*/
                                            //  System.Threading.Thread.Sleep(100);
                                            //      SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                                            //    mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3);
                                            //  richTextBox1.Text = mPage1.usingAddress + "  嘗試連線中請稍候... \r\n";
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("該商品" + mPage1.product_name + "未裝置電子標籤");
                                    dataGridView1.Enabled = true;
                                }
                                break;
                            }
                        }
                    }
                    SalePageList.Clear();
                    /*  for (int i=0;i< SalePageListUpdate.Count;i++) {
                          SalePageList.Add(SalePageListUpdate[i]);
                      }*/

                    /*   for (int i = 0; i < SalePageListUpdate.Count; i++)
                       {
                           Page1 mPageA = new Page1();
                           mPageA.no = SalePageListUpdate[i].no;
                           mPageA.BleAddress = SalePageListUpdate[i].BleAddress;
                           mPageA.barcode = SalePageListUpdate[i].barcode;
                           mPageA.product_name = SalePageListUpdate[i].product_name;
                           mPageA.Brand = SalePageListUpdate[i].Brand;
                           mPageA.specification = SalePageListUpdate[i].specification;
                           mPageA.price = SalePageListUpdate[i].price;
                           mPageA.Web = SalePageListUpdate[i].Web;
                           mPageA.usingAddress = SalePageListUpdate[i].usingAddress;
                           mPageA.HeadertextALL = SalePageListUpdate[i].HeadertextALL;
                           mPageA.Special_offer = SalePageListUpdate[i].Special_offer;
                           mPageA.onsale = SalePageListUpdate[i].onsale;
                           mPageA.onSaleTimeS = SalePageListUpdate[i].onSaleTimeS;
                           mPageA.ProductStyle = SalePageListUpdate[i].ProductStyle;
                           mPageA.onSaleTimeE = SalePageListUpdate[i].onSaleTimeE;
                           mPageA.actionName = SalePageListUpdate[i].actionName;
                           mPageA.APLink = SalePageListUpdate[i].APLink;
                           SalePageList.Add(mPageA);
                       }*/
                    SalePageList.AddRange(SalePageListUpdate);


                }

                // dataGridView3.Rows[datagridview2no].Cells[4].Value = "連線中";

                //  mSmcEsl.TransformImageToData(bmp);

                // mSmcEsl.ConnectBleDevice(mPage1.usingAddress);

                // mSmcEsl.WriteESLData(mPage1.usingAddress);
                macaddress = PageList[listcount].usingAddress;
                string sub = Environment.CurrentDirectory;
                Console.WriteLine("sub" + sub);

                // dataGridView3.Rows[0].Cells[4].Value = "連線中";


            }
            if (saletimemsg != "")
            {
                //     CheckBeaconTimer.Stop();
                      MessageBox.Show(saletimemsg, "Beacon訊息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                /*     if (result == DialogResult.OK)
                     {
                         CheckBeaconTimer.Start();
                     }
                     */

            }
        }

        private void dataGridView4_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            Console.WriteLine("e.ColumnIndexerror" + e.ColumnIndex);
        }

        private void dataGridView1_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            e.Column.SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private void beacon_data_set(List<string> sss,string beaconStime, string beaconEtime,string product_onsale)
        {
            try {

                var request = (HttpWebRequest)WebRequest.Create("https://api.ihoin.com/esl_test/esl_beacon_set");

                /*string json = "{\"user\":\"test\"," +
                     "\"password\":\"bla\"}";*/
                // var postData = "productarr="+ sss;
                /* var postData = "productStime="+ beaconStime;
                 postData += "&productEtime="+ beaconEtime;
                 postData += "&product_onsale="+ product_onsale;*/

                ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(AcceptAllCertifications);

                string postData = new JavaScriptSerializer().Serialize(new
                {
                    productarr = sss,
                    productStime = beaconStime,
                    productEtime = beaconEtime,
                    product_onsale = product_onsale
                });
                Console.WriteLine("postData" + postData);
                var data = Encoding.ASCII.GetBytes(postData);
                Console.WriteLine("data" + data);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.ContentLength = data.Length;

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                var response = (HttpWebResponse)request.GetResponse();

                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
            } catch (Exception ex) {
                MessageBox.Show(ex.ToString(), "error", MessageBoxButtons.OK);
            }
 
        }

        public bool AcceptAllCertifications(object sender, System.Security.Cryptography.X509Certificates.X509Certificate certification, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }


        private void button7_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Enabled = true;
        }
        private void writeESLdataBySzie(string deviceIP,string EslSize,bool writeType) {

            Page1 mPage1 = new Page1();
            Console.WriteLine("writeType" + writeType);
            for (int i = 0; i < PageList.Count; i++)
            {
                if ((PageList[i].APLink + ":8899") == deviceIP && PageList[i].UpdateState == null)
                {
                    Console.WriteLine("aadddss" + PageList[i].APLink);
                    mPage1 = PageList[i];
                    break;
                }
                if (i == PageList.Count - 1)
                {
                    List<Page1> list = PageList.GroupBy(a => a.APLink).Select(g => g.First()).Where(p => p.UpdateState == null).ToList();
                    foreach (Page1 p in list)
                    {
                        OldRunAPList.Add(p.APLink);

                    }
                }

            }
            Bitmap bmp;
            if (mPage1.actionName == "down" || mPage1.actionName == "sale" || mPage1.actionName == "reset" || mPage1.actionName == "EslStyleChangeUpdate")
            {



                //dr.Cells[17].Value = DateTime.Now.ToString();
                macaddress = mPage1.usingAddress;
                Console.WriteLine("WWWWWWTTTFFFFFFFFFBBBCCC");
                if (mPage1.actionName == "reset")
                {
                    foreach (DataGridViewRow dr in this.dataGridView4.Rows)
                    {
                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                        {
                            dataGridView4.Rows[dr.Index].Cells[0].Selected = true;
                        }
                    }
                }
                else
                {
                    foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                    {
                        if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                        {
                            dataGridView1.Rows[dr.Index].Cells[0].Selected = true;
                        }
                    }
                }


                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {
                    Console.WriteLine("mPage1.APLink" + mPage1.APLink);
                    if (kvp.Key.Contains(mPage1.APLink))
                    {
                        Console.WriteLine("ININ");
                        if (mPage1.actionName == "down")
                        {
                            Console.WriteLine("d");

                            mPage1.ESLSize = EslSize;
                            if (EslSize == "01")
                                bmp = mElectronicPriceData.setESLimage_29(mPage1.usingAddress, "3.04");
                            else if (EslSize == "02")
                                bmp = mElectronicPriceData.setESLimage_42(mPage1.usingAddress, "3.04");
                            else
                                bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            pictureBoxPage1.Image = bmp;
                        }
                        /*     if (immediateUpdate)
                             {
                                 Console.WriteLine("d");
                                 if (mPage1.onsale == "V")
                                 {
                                     bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                              mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                 mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLSaleFormat);
                                 }
                                 else
                                 {
                                     bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                              mPage1.specification, mPage1.price, mPage1.Special_offer,
                                                 mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, ESLFormat);
                                 }

                                 kvp.Value.mSmcEsl.TransformImageToData(bmp);
                                 pictureBoxPage1.Image = bmp;
                             }*/

                        if (mPage1.actionName == "reset")
                        {
                            Console.WriteLine("r");
                            //bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                            mPage1.ESLSize = EslSize;
                            if (EslSize == "01")
                                bmp = mElectronicPriceData.setESLimage_29(mPage1.usingAddress, "3.04");
                            else if (EslSize == "02")
                                bmp = mElectronicPriceData.setESLimage_42(mPage1.usingAddress, "3.04");
                            else
                                bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            pictureBoxPage1.Image = bmp;
                        }

                        if (mPage1.actionName == "sale")
                        {

                            foreach (DataGridViewRow dr in this.dataGridView7.Rows)
                            {
                                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && ("0" + dr.Cells[4].Value.ToString()) == EslSize)
                                {
                                    PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                }
                            }

                            Console.WriteLine("s");
                            if (mPage1.onsale == "V")
                            {

                                List<string> Format = new List<string>();
                                if (EslSize == "01")
                                {
                                    Format = ESLSale29Format;
                                }
                                else if (EslSize == "02")
                                {
                                    Format = ESLSale42Format;
                                }
                                else
                                {
                                    Format = ESLSaleFormat;
                                }
                                mPage1.ESLSize = EslSize;
                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                         mPage1.specification, mPage1.price, mPage1.Special_offer,
                                            mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, Format);
                            }
                            else
                            {

                                foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                                {
                                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && ("0" + dr.Cells[4].Value.ToString()) == EslSize)
                                    {
                                        PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                    }
                                }

                                List<string> Format = new List<string>();
                                if (EslSize == "01")
                                {
                                    Format = ESL29Format;
                                }
                                else if (EslSize == "02")
                                {
                                    Format = ESL42Format;
                                }
                                else
                                {
                                    Format = ESLFormat;
                                }
                                mPage1.ESLSize = EslSize;
                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                         mPage1.specification, mPage1.price, mPage1.Special_offer,
                                            mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, Format);
                            }

                            kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            pictureBoxPage1.Image = bmp;
                            //    bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                        }
                        if (mPage1.actionName == "EslStyleChangeUpdate")
                        {
                            Console.WriteLine("s");
                            if (mPage1.onsale == "V")
                            {
                                foreach (DataGridViewRow dr in this.dataGridView7.Rows)
                                {
                                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && ("0" + dr.Cells[4].Value.ToString()) == EslSize)
                                    {
                                        PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                    }
                                }

                                List<string> Format = new List<string>();
                                if (EslSize == "01")
                                {
                                    Format = ESLSale29Format;
                                }
                                else if (EslSize == "02")
                                {
                                    Format = ESLSale42Format;
                                }
                                else
                                {
                                    Format = ESLSaleFormat;
                                }
                                mPage1.ESLSize = EslSize;
                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                        mPage1.specification, mPage1.price, mPage1.Special_offer,
                                           mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, Format);
                                kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            }
                            else
                            {
                                foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                                {
                                    if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && ("0" + dr.Cells[4].Value.ToString()) == EslSize)
                                    {
                                        PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                    }
                                }

                                List<string> Format = new List<string>();
                                if (EslSize == "01")
                                {
                                    Format = ESL29Format;
                                }
                                else if (EslSize == "02")
                                {
                                    Format = ESL42Format;
                                }
                                else
                                {
                                    Format = ESLFormat;
                                }
                                mPage1.ESLSize = EslSize;
                                bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                       mPage1.specification, mPage1.price, mPage1.Special_offer,
                                          mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, Format);
                                kvp.Value.mSmcEsl.TransformImageToData(bmp);
                            }

                            pictureBoxPage1.Image = bmp;
                            //    bmp = mElectronicPriceData.writeIDimage(mPage1.usingAddress);
                        }


                        dataGridView1.Rows[Convert.ToInt32(mPage1.no) - 1].Cells[17].Value = DateTime.Now.ToString();
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, Convert.ToInt32(mPage1.no) - 1, false, openExcelAddress, excel, excelwb, mySheet);
                        //kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        //  kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress, 0);
                       if(writeType)
                            kvp.Value.mSmcEsl.WriteESLDataWithBle2("FFFFFFFF");
                       else
                            kvp.Value.mSmcEsl.WriteESLDataWithBle();
                        //.Threading.Thread.Sleep(1000);
                        EslUdpTest.SmcEsl mSmcEsl = kvp.Value.mSmcEsl;
                        //  mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3);
                        //   System.Threading.Thread.Sleep(200);
                        //   Console.WriteLine("WWWWWWWWWWWWWWWWWWWWWWWTTTTTTTTT1");
                    }
                }
                //  mSmcEsl.ConnectBleDevice(mPage1.usingAddress);
                foreach (DataGridViewRow dr in this.dataGridView3.Rows)
                {
                    if (macaddress == dr.Cells[0].Value.ToString())
                    {
                        dr.Cells[4].Value = "連線中";
                    }
                }

                // int CurrentRow = dataGridView1.CurrentRow.Index;
                // dataGridView1.Rows[CurrentRow].Cells[17].Value = DateTime.Now.ToString();
                //       richTextBox1.Text = "正連接:" + mPage1.usingAddress + "\r\n" + richTextBox1.Text;
                // mSmcEsl.WriteESLData(PageList[listcount].usingAddress);
            }
            else if (mPage1.actionName == "saletime")
            {
                macaddress = mPage1.usingAddress;
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.usingAddress)
                    {
                        Console.WriteLine("乾 最好進不來");
                        dataGridView1.Rows[dr.Index].Cells[0].Selected = true;

                    }
                }
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {

                    if (kvp.Key.Contains(mPage1.APLink))
                    {

                        string format = "yyyy/MM/dd HH:mm:ss";
                        string start = Convert.ToDateTime(mPage1.onSaleTimeS).ToString("yyyy/MM/dd HH:mm:ss");
                        string end = Convert.ToDateTime(mPage1.onSaleTimeE).ToString("yyyy/MM/dd HH:mm:ss");
                        DateTime strDate = DateTime.ParseExact(start, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                        DateTime endDate = DateTime.ParseExact(end, format, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces);
                        if (DateTime.Compare(strDate, DateTime.Now) < 0 && DateTime.Compare(endDate, DateTime.Now) > 0)
                        {

                            foreach (DataGridViewRow dr in this.dataGridView7.Rows)
                            {
                                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && ("0" + dr.Cells[4].Value.ToString()) == EslSize)
                                {
                                    PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                }
                            }

                            List<string> Format = new List<string>();
                            if (EslSize == "01")
                            {
                                Format = ESLSale29Format;
                            }
                            else if (EslSize == "02")
                            {
                                Format = ESLSale42Format;
                            }
                            else
                            {
                                Format = ESLSaleFormat;
                            }
                            mPage1.ESLSize = EslSize;
                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, Format);
                            pictureBoxPage1.Image = bmp;

                            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                            {
                                if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() == mPage1.product_name)
                                {
                                    dr.Cells[15].Value = "V";
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                }
                            }
                        }
                        else
                        {
                            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                            {
                                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && ("0" + dr.Cells[4].Value.ToString()) == EslSize)
                                {
                                    PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                }
                            }
                            List<string> Format = new List<string>();
                            if (EslSize == "01")
                            {
                                Format = ESL29Format;
                            }
                            else if (EslSize == "02")
                            {
                                Format = ESL42Format;
                            }
                            else
                            {
                                Format = ESLFormat;
                            }
                            mPage1.ESLSize = EslSize;
                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                     mPage1.specification, mPage1.price, mPage1.Special_offer,
                                        mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, Format);
                            pictureBoxPage1.Image = bmp;
                            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                            {
                                if (dr.Cells[6].Value != null && dr.Cells[6].Value.ToString() == mPage1.product_name)
                                {
                                    dr.Cells[15].Value = "X";
                                    dr.Cells[19].Value = DBNull.Value;
                                    dr.Cells[20].Value = DBNull.Value;
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                }
                            }

                        }
                        dataGridView1.Rows[Convert.ToInt32(mPage1.no) - 1].Cells[17].Value = DateTime.Now.ToString();
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, Convert.ToInt32(mPage1.no) - 1, false, openExcelAddress, excel, excelwb, mySheet);
                        kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        // kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.usingAddress, 0);
                        if (writeType)
                            kvp.Value.mSmcEsl.WriteESLDataWithBle2("FFFFFFFF");
                        else
                            kvp.Value.mSmcEsl.WriteESLDataWithBle();
                        //  mSmcEsl.UpdataESLDataFromBuffer(mPage1.usingAddress, 0, 3);
                        //  Console.WriteLine("WWWWWWWWWWWWWWWWWWWWWWWTTTTTTTTT2");
                        // System.Threading.Thread.Sleep(200);

                    }
                }

            }

            else
            {

                macaddress = mPage1.BleAddress;
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == mPage1.BleAddress)
                    {
                        dataGridView1.Rows[dr.Index].Cells[0].Selected = true;
                    }
                }
                Console.WriteLine("ReadType_mPage1.APLink:" + mPage1.APLink);
                foreach (KeyValuePair<string, EslObject> kvp in mDictSocket)
                {

                    if (kvp.Key.Contains(mPage1.APLink))
                    {
                        if (mPage1.onsale == "V")
                        {
                            foreach (DataGridViewRow dr in this.dataGridView7.Rows)
                            {
                                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && ("0" + dr.Cells[4].Value.ToString()) == EslSize)
                                {
                                    PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                }
                            }

                            List<string> Format = new List<string>();
                            if (EslSize == "01")
                            {
                                Format = ESLSale29Format;
                            }
                            else if (EslSize == "02")
                            {
                                Format = ESLSale42Format;
                            }
                            else
                            {
                                Format = ESLSaleFormat;
                            }
                            mPage1.ESLSize = EslSize;
                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                         mPage1.specification, mPage1.price, mPage1.Special_offer,
                                            mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, Format);
                        }
                        else
                        {

                            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
                            {
                                if (dr.Cells[2].Value != null && dr.Cells[2].Value.ToString() == "V" && dr.Cells[4].Value != null && ("0" + dr.Cells[4].Value.ToString()) == EslSize)
                                {
                                    PageList[listcount].ProductStyle = dr.Cells[1].Value.ToString();
                                }
                            }
                            List<string> Format = new List<string>();
                            if (EslSize == "01")
                            {
                                Format = ESL29Format;
                            }
                            else if (EslSize == "02")
                            {
                                Format = ESL42Format;
                            }
                            else
                            {
                                Format = ESLFormat;
                            }
                            mPage1.ESLSize = EslSize;
                            bmp = mElectronicPriceData.setPage1("Calibri", mPage1.product_name, mPage1.Brand,
                                          mPage1.specification, mPage1.price, mPage1.Special_offer,
                                             mPage1.barcode, mPage1.Web, mPage1.usingAddress, mPage1.HeadertextALL, Format);
                        }

                        dataGridView1.Rows[Convert.ToInt32(mPage1.no) - 1].Cells[17].Value = DateTime.Now.ToString();
                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 17, Convert.ToInt32(mPage1.no) - 1, false, openExcelAddress, excel, excelwb, mySheet);
                        kvp.Value.mSmcEsl.TransformImageToData(bmp);
                        if (writeType)
                            kvp.Value.mSmcEsl.WriteESLDataWithBle2("FFFFFFFF");
                        else
                            kvp.Value.mSmcEsl.WriteESLDataWithBle();
                        // kvp.Value.mSmcEsl.writeESLDataBuffer(mPage1.BleAddress,0);
                        //  kvp.Value.mSmcEsl.ConnectBleDevice(mPage1.BleAddress);

                    }
                }
            }
        }


        private   string writeESLsuccess(string deviceIP,string str_data)
        {
            BleWriteTimer.Stop();
            str_data = "全部資料寫入完成";
            //   this.progressBar1.Visible = false; //顯示進度條

            //   stopwatch.Stop();//碼錶停止
            //   TimeSpan ts = stopwatch.Elapsed;

            // Format and display the TimeSpan value.
            //   string elapsedTime = String.Format("{0:00} 分 {1:00} 秒 {2:000} 毫秒", ts.Minutes, ts.Seconds, ts.Milliseconds);

            //ConnectBleTimeOut.Stop();
            //str_data = "AP 更新 ESL 完成";
            updateESLper.Text = (Convert.ToUInt32(updateESLper.Text) + 1).ToString();
            //  richTextBox1.Text = "斷線成功" + macaddress + "\r\n" + richTextBox1.Text;
            Console.WriteLine("writeESLsuccess:" + progressBar1.Maximum);
            progressBar1.Value += progressBar1.Step;
            Console.WriteLine("AP 更新 PageList.Count" + PageList.Count);
            for (int i = 0; i < PageList.Count; i++)
            {
                //  Console.WriteLine(i+"AP 更新 ESL" + PageList[i].usingAddress + deviceIP);
                if (PageList[i].APLink + ":8899" == deviceIP && PageList[i].UpdateState == null)
                {
                    str_data = "AP 更新" + PageList[i].usingAddress + "完成";
                    Console.WriteLine(i + "AP 更新 ESL" + PageList[i].usingAddress + deviceIP);
                    if (PageList[i].actionName == "reset")// 地3業還原
                    {
                        PageList[i].UpdateState = "更新成功";
                        PageList[i].UpdateTime = DateTime.Now.ToString();
                        foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                        {
                            //    Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + macaddress);
                            if (dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                            {
                                //  Console.WriteLine("ININJ");
                                dr4.Cells[2].Value = DateTime.Now.ToString();
                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                dr4.Cells[2].Style.BackColor = Color.Green;
                                dataGridView4.Rows[dr4.Index].Cells[0].Selected = false;
                                dr4.Cells[0].Value = false;
                                dr4.Cells[3].Value = "";
                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                dr4.Cells[6].Value = "未绑定";
                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                //    dr4.Cells[6].Value = dr.Cells[6].Value;
                                break;
                            }
                        }

                        foreach (DataGridViewRow dr in dataGridView1.Rows)
                        {

                            if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                            {
                                if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                {
                                    Console.WriteLine("=========================");
                                    if (dr.Cells[1].Value.ToString().Contains(',' + PageList[i].usingAddress))
                                    {
                                        int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + PageList[i].usingAddress);
                                        dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                        dr.Cells[12].Value = dr.Cells[1].Value;
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    }
                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress + ','))
                                    {

                                        int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress + ',');
                                        dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                        dr.Cells[12].Value = dr.Cells[1].Value;
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    }

                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                    {

                                        int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress);
                                        dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                                        dr.Cells[12].Value = dr.Cells[1].Value;
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    }

                                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == "")
                                    {
                                        MessageBox.Show(dr.Cells[6].Value.ToString() + "無綁定ESL自動下架");
                                        dr.Cells[0].ReadOnly = true;
                                        dr.Cells[0].Value = false;
                                        dr.Cells[13].Value = "";
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        dr.DefaultCellStyle.ForeColor = Color.Gray;
                                    }
                                    break;
                                }
                            }
                            Console.WriteLine("=OOOOOOOOOO===");
                        }
                    }
                    else
                    {//第一頁功能

                        PageList[i].UpdateState = "更新成功";
                        PageList[i].UpdateTime = DateTime.Now.ToString();
                        List<string> nullbeacon = new List<string>();
                        foreach (DataGridViewRow dr in dataGridView1.Rows)
                        {

                            // Console.WriteLine("dr.Cells[1].Value.ToString()");
                            if (PageList[i].actionName == "down" || PageList[i].actionName == "sale" || PageList[i].actionName == "saletime" || PageList[i].actionName == "EslStyleChangeUpdate")
                            {



                                if (dr.Cells[1].Value != null)
                                {
                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                    {
                                        //// Console.WriteLine("macaddress" + macaddress);
                                        // dr.Cells[4].Style.BackColor = Color.Green;
                                        //  dr.Cells[4].Value = DateTime.Now.ToString();
                                        // Console.WriteLine("macaddress" + dr.Cells[6].Value);
                                        // dr.Cells[18].Value = DateTime.Now.ToString();

                                        dr.Cells[4].Style.BackColor = Color.Green;
                                        dr.Cells[4].Value = DateTime.Now.ToString();
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 4, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);

                                        /* PageList[i].UpdateState = "更新成功";
                                         PageList[i].UpdateTime = DateTime.Now.ToString();*/
                                        dataGridView1.Rows[dr.Index].Cells[0].Selected = false;
                                        Page1 mPage1 = PageList[i];
                                        if (PageList[i].actionName == "down")
                                        {

                                            foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                            {
                                                //  Console.WriteLine(dr4.Cells[1].Value.ToString()+"jjjjjjj" + PageList[i].usingAddress);
                                                if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString() == PageList[i].usingAddress)
                                                {
                                                    //Console.WriteLine("ININJ");
                                                    dr4.Cells[2].Value = DateTime.Now.ToString();
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr4.Cells[2].Style.BackColor = Color.Green;
                                                    dataGridView4.Rows[dr4.Index].Cells[0].Selected = false;
                                                    dr4.Cells[0].Value = false;
                                                    dr4.Cells[3].Value = "";
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr4.Cells[6].Value = "未綁定";
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    //    dr4.Cells[6].Value = dr.Cells[6].Value;
                                                    break;
                                                }
                                            }




                                            if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                                            {
                                                if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                {
                                                    Console.WriteLine(PageList[i].usingAddress + "=========================" + dr.Cells[1].Value.ToString());
                                                    if (dr.Cells[1].Value.ToString().Contains(',' + PageList[i].usingAddress))
                                                    {
                                                        int changeaddr = dr.Cells[1].Value.ToString().IndexOf(',' + PageList[i].usingAddress);
                                                        dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                                        dr.Cells[12].Value = dr.Cells[1].Value;
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    }
                                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress + ','))
                                                    {

                                                        int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress + ',');
                                                        dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 13);
                                                        dr.Cells[12].Value = dr.Cells[1].Value;
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    }

                                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                    {

                                                        int changeaddr = dr.Cells[1].Value.ToString().IndexOf(PageList[i].usingAddress);
                                                        dr.Cells[1].Value = dr.Cells[1].Value.ToString().Remove(changeaddr, 12);
                                                        dr.Cells[12].Value = dr.Cells[1].Value;
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    }

                                                    if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() == "")
                                                    {
                                                        dr.Cells[0].ReadOnly = true;
                                                        dr.Cells[0].Value = false;
                                                        dr.Cells[2].Value = DBNull.Value;
                                                        dr.Cells[13].Value = "";
                                                        dr.Cells[19].Value = DBNull.Value;
                                                        dr.Cells[20].Value = DBNull.Value;
                                                        dr.Cells[21].Value = DBNull.Value;
                                                        dr.Cells[22].Value = DBNull.Value;
                                                        dr.Cells[23].Value = DBNull.Value;

                                                        nullbeacon.Add(dr.Cells[5].Value.ToString());
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 19, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 20, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 21, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 22, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 23, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                        dr.DefaultCellStyle.ForeColor = Color.Gray;
                                                    }

                                                }
                                            }

                                            // dr.DefaultCellStyle.ForeColor = Color.Gray;
                                            dr.Cells[15].Value = "X";
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.Cells[16].Value = "X";
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 16, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            // dr.Cells[1].Value = "";
                                            //  mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.Cells[18].Value = DateTime.Now.ToString();
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.Cells[13].Value = "";
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            // dr.Cells[12].Value = "";
                                            //  mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            // dr.Cells[0].ReadOnly = true;
                                            //dr.Cells[0].Value = false;

                                        }

                                        /*   if (immediateUpdate)
                                           {
                                               dr.Cells[13].Value = PageList[i].ProductStyle;
                                               Console.WriteLine(PageList[i].product_name + "PageList[i].ProductStyle" + PageList[i].ProductStyle);
                                               mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                               productState(dr);
                                               dr.Cells[18].Value = DateTime.Now.ToString();
                                               mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);

                                               foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                               {
                                                   Console.WriteLine(dr4.Cells[1].Value.ToString() + "jjjjjjj" + PageList[i].usingAddress);
                                                   if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString() == PageList[i].usingAddress)
                                                   {
                                                       Console.WriteLine("ININJ");
                                                       dr4.Cells[2].Value = DateTime.Now.ToString();
                                                       mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                       dr4.Cells[2].Style.BackColor = Color.Green;
                                                       dr4.Cells[3].Value = dr.Cells[13].Value;
                                                       mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                       dr4.Cells[6].Value = dr.Cells[6].Value;
                                                       mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                       break;
                                                   }
                                               }

                                           }*/

                                      
                                        if (PageList[i].actionName == "sale")
                                        {


                                            string ESLStyle = "";

                                            if (PageList[i].onsale == "V")
                                            {
                                                foreach (DataGridViewRow dr7 in dataGridView7.Rows)
                                                {
                                                    if (PageList[i].ESLSize == "01" && dr7.Cells[4].Value.ToString() == "1" && dr7.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr7.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "02" && dr7.Cells[4].Value.ToString() == "2" && dr7.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr7.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "00" && dr7.Cells[4].Value.ToString() == "0" && dr7.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr7.Cells[1].Value.ToString();
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                foreach (DataGridViewRow dr2 in dataGridView2.Rows)
                                                {
                                                    if (PageList[i].ESLSize == "01" && dr2.Cells[4].Value.ToString() == "1" && dr2.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr2.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "02" && dr2.Cells[4].Value.ToString() == "2" && dr2.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr2.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "00" && dr2.Cells[4].Value.ToString() == "0" && dr2.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr2.Cells[1].Value.ToString();
                                                    }
                                                }
                                            }
                                                dr.Cells[13].Value = ESLStyle;
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 5, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 6, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 7, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 8, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 9, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 10, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 11, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.Cells[18].Value = DateTime.Now.ToString();
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);

                                            dr.Cells[5].Style.ForeColor = Color.Empty;
                                            dr.Cells[6].Style.ForeColor = Color.Empty;
                                            dr.Cells[7].Style.ForeColor = Color.Empty;
                                            dr.Cells[8].Style.ForeColor = Color.Empty;
                                            dr.Cells[9].Style.ForeColor = Color.Empty;
                                            dr.Cells[10].Style.ForeColor = Color.Empty;
                                            dr.Cells[11].Style.ForeColor = Color.Empty;
                                            dr.DefaultCellStyle.ForeColor = Color.Black;
                                            foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                            {
                                                // Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + PageList[i].usingAddress);
                                                if (dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                {
                                                    Console.WriteLine("ININJ");
                                                    dr4.Cells[2].Value = DateTime.Now.ToString();
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr4.Cells[2].Style.BackColor = Color.Green;
                                                    dr4.Cells[3].Value = ESLStyle;
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr4.Cells[6].Value = dr.Cells[6].Value;
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    break;
                                                }
                                            }

                                        }

                                        if (PageList[i].actionName == "saletime")
                                        {
                                            string ESLStyle = "";
                                            if (PageList[i].onsale == "V")
                                            {
                                                foreach (DataGridViewRow dr7 in dataGridView7.Rows)
                                                {
                                                    if (PageList[i].ESLSize == "01" && dr7.Cells[4].Value.ToString() == "1" && dr7.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr7.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "02" && dr7.Cells[4].Value.ToString() == "2" && dr7.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr7.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "00" && dr7.Cells[4].Value.ToString() == "0" && dr7.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr7.Cells[1].Value.ToString();
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                foreach (DataGridViewRow dr2 in dataGridView2.Rows)
                                                {
                                                    if (PageList[i].ESLSize == "01" && dr2.Cells[4].Value.ToString() == "1" && dr2.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr2.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "02" && dr2.Cells[4].Value.ToString() == "2" && dr2.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr2.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "00" && dr2.Cells[4].Value.ToString() == "0" && dr2.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr2.Cells[1].Value.ToString();
                                                    }
                                                }
                                            }
                                            dr.Cells[13].Value = ESLStyle;
                                            Console.WriteLine(PageList[i].product_name + "PageList[i].ProductStyle" + PageList[i].ProductStyle);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            productState(dr);
                                            dr.Cells[18].Value = DateTime.Now.ToString();
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.Cells[3].Value = false;



                                            foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                            {
                                                //    Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + PageList[i].usingAddress);
                                                if (dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                {
                                                    //   Console.WriteLine("ININJ");
                                                    dr4.Cells[2].Value = DateTime.Now.ToString();
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr4.Cells[2].Style.BackColor = Color.Green;
                                                    dr4.Cells[3].Value = ESLStyle;
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr4.Cells[6].Value = dr.Cells[6].Value;
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    break;
                                                }
                                            }
                                        }

                                        if (PageList[i].actionName == "EslStyleChangeUpdate")
                                        {
                                            string ESLStyle = "";
                                                  if (PageList[i].onsale == "V")
                                            {
                                                foreach (DataGridViewRow dr7 in dataGridView7.Rows)
                                                {
                                                    if (PageList[i].ESLSize == "01" && dr7.Cells[4].Value.ToString() == "1" && dr7.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr7.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "02" && dr7.Cells[4].Value.ToString() == "2" && dr7.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr7.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "00" && dr7.Cells[4].Value.ToString() == "0" && dr7.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr7.Cells[1].Value.ToString();
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                foreach (DataGridViewRow dr2 in dataGridView2.Rows)
                                                {
                                                    if (PageList[i].ESLSize == "01" && dr2.Cells[4].Value.ToString() == "1" && dr2.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr2.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "02" && dr2.Cells[4].Value.ToString() == "2" && dr2.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr2.Cells[1].Value.ToString();
                                                    }
                                                    else if (PageList[i].ESLSize == "00" && dr2.Cells[4].Value.ToString() == "0" && dr2.Cells[2].Value.ToString() == "V")
                                                    {
                                                        ESLStyle = dr2.Cells[1].Value.ToString();
                                                    }
                                                }
                                            }
                                            dr.Cells[13].Value = ESLStyle;
                                            Console.WriteLine(PageList[i].product_name + "PageList[i].ProductStyle" + PageList[i].ProductStyle);
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            dr.Cells[18].Value = DateTime.Now.ToString();
                                            mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                            foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                            {
                                                //   Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + PageList[i].usingAddress);
                                                if (dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                                {
                                                    //    Console.WriteLine("ININJ");
                                                    dr4.Cells[2].Value = DateTime.Now.ToString();
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr4.Cells[2].Style.BackColor = Color.Green;
                                                    dr4.Cells[3].Value = ESLStyle;
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    dr4.Cells[6].Value = dr.Cells[6].Value;
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                    break;
                                                }
                                            }
                                        }


                                        foreach (DataGridViewRow dr3 in dataGridView3.Rows)
                                        {
                                            if (dr3.Cells[0].Value.ToString().Contains(PageList[i].usingAddress))
                                            {
                                                dr3.Cells[4].Value = "已完成";
                                                dr3.Cells[6].Value = DateTime.Now.ToString();
                                                // UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                                break;
                                            }
                                        }

                                        /*   foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                           {
                                               Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + PageList[i].usingAddress);
                                               if (dr4.Cells[1].Value.ToString().Contains(PageList[i].usingAddress))
                                               {
                                                   Console.WriteLine("ININJ");
                                                   dr4.Cells[2].Value = DateTime.Now.ToString();
                                                   mExcelData.dataGridViewRowCellUpdate(dataGridView4,2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                   dr4.Cells[2].Style.BackColor = Color.Green;
                                                   dr4.Cells[3].Value = styleName;
                                                   mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                   dr4.Cells[6].Value = dr.Cells[6].Value;
                                                   mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                   break;
                                               }
                                           }*/
                                    }
                                }



                            }
                            else
                            {
                                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString().Contains(PageList[i].BleAddress))
                                {
                                    foreach (DataGridViewRow drnew in dataGridView1.Rows)
                                    {
                                        if (drnew.Cells[12].Value != null && drnew.Cells[12].Value.ToString() != "")
                                        {
                                            if (drnew.Cells[12].Value.ToString().Contains(PageList[i].BleAddress))
                                            {
                                                Console.WriteLine("cccccccc" + drnew.Cells[12].Value);
                                                if (drnew.Cells[12].Value != null)
                                                    a = drnew.Cells[12].Value;
                                                if (drnew.Cells[13].Value != null)
                                                    b = drnew.Cells[13].Value;
                                                if (drnew.Cells[14].Value != null)
                                                    c = drnew.Cells[14].Value;
                                                if (drnew.Cells[15].Value != null)
                                                    d = drnew.Cells[15].Value;
                                                easd = drnew.Cells[16].Value;
                                                //  drnew.Cells[12].Value = DBNull.Value;
                                                if (drnew.Cells[12].Value.ToString().Contains(',' + PageList[i].usingAddress))
                                                {
                                                    int changeaddr = drnew.Cells[12].Value.ToString().IndexOf(',' + PageList[i].usingAddress);
                                                    drnew.Cells[12].Value = drnew.Cells[12].Value.ToString().Remove(changeaddr, 13);


                                                }
                                                if (drnew.Cells[12].Value.ToString().Contains(PageList[i].usingAddress + ','))
                                                {

                                                    int changeaddr = drnew.Cells[12].Value.ToString().IndexOf(PageList[i].usingAddress + ',');
                                                    drnew.Cells[12].Value = drnew.Cells[12].Value.ToString().Remove(changeaddr, 13);

                                                }

                                                if (drnew.Cells[12].Value.ToString().Contains(PageList[i].usingAddress))
                                                {

                                                    int changeaddr = drnew.Cells[12].Value.ToString().IndexOf(PageList[i].usingAddress);
                                                    drnew.Cells[12].Value = drnew.Cells[12].Value.ToString().Remove(changeaddr, 12);

                                                }
                                                if (drnew.Cells[12].Value.ToString().Length == 0)
                                                {
                                                    drnew.DefaultCellStyle.ForeColor = Color.Gray;
                                                    drnew.Cells[0].Value = false;
                                                    drnew.Cells[0].ReadOnly = true;
                                                }
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                drnew.Cells[13].Value = DBNull.Value;
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                drnew.Cells[14].Value = DBNull.Value;
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView1, 14, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                if (drnew.Cells[12].Value != null && drnew.Cells[12].Value.ToString() == "")
                                                {
                                                    drnew.Cells[16].Value = "X";
                                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 16, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);

                                                    // drnew.Cells[15].Value = "X";
                                                    //     mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, drnew.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                }
                                                break;
                                            }

                                        }

                                    }

                                    Console.WriteLine("b" + b.ToString());
                                    Console.WriteLine("c" + c.ToString());
                                    Console.WriteLine("d" + d.ToString());
                                    // dr.Cells[12].Value = dr.Cells[1].Value;
                                    if (dr.Cells[12].Value.ToString().Length > 0)
                                        dr.Cells[12].Value = dr.Cells[12].Value.ToString() + "," + PageList[i].usingAddress;
                                    else
                                        dr.Cells[12].Value = PageList[i].usingAddress;

                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 1, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 12, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    dr.Cells[13].Value = b;
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    dr.Cells[14].Value = c;
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 14, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    /* if (dr.Cells[12].Value != null && dr.Cells[12].Value.ToString().Length == 12) {
                                         dr.Cells[15].Value = "X";
                                         mExcelData.dataGridViewRowCellUpdate(dataGridView1, 15, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                     }*/
                                    dr.Cells[16].Value = "V";
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 16, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);

                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 5, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 6, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 7, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 8, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 9, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 10, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    mExcelData.dataGridViewRowCellUpdate(dataGridView1, 11, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                    dr.DefaultCellStyle.ForeColor = Color.Black;
                                    dr.Cells[1].Style.ForeColor = Color.Black;
                                    dr.Cells[0].ReadOnly = false;
                                    dr.Cells[0].Value = true;
                                    /*   foreach (DataGridViewRow dr4 in this.dataGridView4.Rows)
                                       {
                                           if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString() == dr.Cells[12].Value.ToString())
                                           {
                                               dr4.Cells[6].Value = dr.Cells[6].Value;
                                           }
                                       }*/
                                }
                                if (dr.Cells[1].Value != null && dr.Cells[1].Value.ToString() != "")
                                {
                                    Console.WriteLine("macaddressAAS" + dr.Cells[1].Value.ToString());
                                    if (dr.Cells[1].Value.ToString().Contains(PageList[i].BleAddress))
                                    {
                                        Console.WriteLine("TTTTTTTTTTTT" + PageList[i].ESLSize+ " "+ PageList[i].BleAddress+ " "+ PageList[i].onsale);
                                        string ESLStyle = "";
                                        if (PageList[i].onsale == "V")
                                        {
                                            foreach (DataGridViewRow dr7 in dataGridView7.Rows)
                                            {
                                               
                                                if (PageList[i].ESLSize == "01" && dr7.Cells[4].Value.ToString() == "1" && dr7.Cells[2].Value.ToString() == "V")
                                                {
                                                    ESLStyle = dr7.Cells[1].Value.ToString();
                                                }
                                                else if (PageList[i].ESLSize == "02" && dr7.Cells[4].Value.ToString() == "2" && dr7.Cells[2].Value.ToString() == "V")
                                                {
                                                    ESLStyle = dr7.Cells[1].Value.ToString();
                                                }
                                                else if(PageList[i].ESLSize == "00" && dr7.Cells[4].Value.ToString() == "0" && dr7.Cells[2].Value.ToString() == "V")
                                                {
                                                    ESLStyle = dr7.Cells[1].Value.ToString();
                                                }
                                            }
                                        }
                                        else
                                        {
                                            foreach (DataGridViewRow dr2 in dataGridView2.Rows)
                                            {
                                                Console.WriteLine(dr2.Cells[4].Value.ToString() + "  "+dr2.Cells[2].Value.ToString());
                                                if (PageList[i].ESLSize == "01" && dr2.Cells[4].Value.ToString() == "1" && dr2.Cells[2].Value.ToString() == "V")
                                                {
                                                    ESLStyle = dr2.Cells[1].Value.ToString();
                                                }
                                                else if (PageList[i].ESLSize == "02" && dr2.Cells[4].Value.ToString() == "2" && dr2.Cells[2].Value.ToString() == "V")
                                                {
                                                    ESLStyle = dr2.Cells[1].Value.ToString();
                                                }
                                                else if (PageList[i].ESLSize == "00" && dr2.Cells[4].Value.ToString() == "0" && dr2.Cells[2].Value.ToString() == "V")
                                                {
                                                    ESLStyle = dr2.Cells[1].Value.ToString();
                                                }
                                            }
                                        }

                                        dr.Cells[4].Style.BackColor = Color.Green;
                                        dr.Cells[4].Value = DateTime.Now.ToString();
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 4, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        dr.Cells[18].Value = DateTime.Now.ToString();
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 18, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        dr.Cells[13].Value = ESLStyle;
                                        mExcelData.dataGridViewRowCellUpdate(dataGridView1, 13, dr.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                        if (dr.Index % 2 == 1)
                                        {
                                            dr.DefaultCellStyle.BackColor = Color.Beige;
                                        }
                                        else
                                        {
                                            dr.DefaultCellStyle.BackColor = Color.Bisque;
                                        }
                                        //  UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                        PageList[i].UpdateState = "更新成功";
                                        PageList[i].UpdateTime = DateTime.Now.ToString();
                                        foreach (DataGridViewRow dr3 in dataGridView3.Rows)
                                        {
                                            if (dr3.Cells[0].Value.ToString().Contains(PageList[i].BleAddress))
                                            {
                                                dr3.Cells[4].Value = "已完成";
                                                dr3.Cells[6].Value = DateTime.Now.ToString();
                                                /// UpdateESLDen.Text = (Convert.ToInt32(UpdateESLDen.Text) + 1).ToString();
                                                break;
                                            }
                                        }

                                        foreach (DataGridViewRow dr4 in dataGridView4.Rows)
                                        {
                                            // Console.WriteLine("jjjjjjj" + dr4.Cells[1].Value.ToString() + PageList[i].BleAddress);
                                            if (dr4.Cells[1].Value != null && dr4.Cells[1].Value.ToString().Contains(PageList[i].BleAddress))
                                            {
                                                // Console.WriteLine("AA" + dr4.Cells[1].Value.ToString());
                                                dr4.Cells[2].Value = DateTime.Now.ToString();
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 2, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                dr4.Cells[2].Style.BackColor = Color.Green;
                                                dr4.Cells[3].Value = ESLStyle;
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 3, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                dr4.Cells[6].Value = dr.Cells[6].Value;
                                                mExcelData.dataGridViewRowCellUpdate(dataGridView4, 6, dr4.Index, false, openExcelAddress, excel, excelwb, mySheet);
                                                break;

                                            }
                                        }

                                    }

                                }
                            }



                        }

                        if (nullbeacon.Count != 0)
                            beacon_data_set(nullbeacon, "", "", "");

                    }
                    deviceIPData = PageList[i].APLink;
                    // PageList.RemoveAt(i);
                    Console.WriteLine("aaaaaaa");
                    break;
                }
            }

            return str_data;
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl2.SelectedIndex == 1)
            {
                for (int i = 0; i < this.dataGridView2.RowCount; i++)
                {
                    dataGridView2.CurrentCell = null;
                    for (int ee = 0; ee < this.dataGridView2.ColumnCount; ee++)
                    {

                        if (ee != 0 && ee != 1 && ee != 2)
                            this.dataGridView2.Columns[ee].Visible = false;
                    }
                    Console.WriteLine("selectChange"+dataGridView2.Rows[i].Cells[3].Value.ToString());
                    if (dataGridView2.Rows[i].Cells[4].Value.ToString() != "0")
                    {

                        dataGridView2.Rows[i].Visible = false;
                    }
                    else
                    {
                        dataGridView2.Rows[i].Visible = true;
                    }
                        
                }

                for (int i = 0; i < this.dataGridView7.RowCount; i++)
                {
                    dataGridView7.CurrentCell = null;
                    for (int ee = 0; ee < this.dataGridView7.ColumnCount; ee++)
                    {

                        if (ee != 0 && ee != 1 && ee != 2)
                            this.dataGridView7.Columns[ee].Visible = false;
                    }
                    if (dataGridView7.Rows[i].Cells[4].Value.ToString() != "0")
                    {

                        dataGridView7.Rows[i].Visible = false;
                    }
                    else
                    {
                        dataGridView7.Rows[i].Visible = true;
                    }

                }
            }
            else if (tabControl2.SelectedIndex == 2)
            {
                for (int i = 0; i < this.dataGridView2.RowCount; i++)
                {
                    dataGridView2.CurrentCell = null;
                    for (int ee = 0; ee < this.dataGridView2.ColumnCount; ee++)
                    {

                        if (ee != 0 && ee != 1 && ee != 2)
                            this.dataGridView2.Columns[ee].Visible = false;
                    }
                    if (dataGridView2.Rows[i].Cells[4].Value.ToString() != "1")
                    {

                        dataGridView2.Rows[i].Visible = false;
                    }
                    else
                    {
                        dataGridView2.Rows[i].Visible = true;
                    }
                }


                for (int i = 0; i < this.dataGridView7.RowCount; i++)
                {
                    dataGridView7.CurrentCell = null;
                    for (int ee = 0; ee < this.dataGridView7.ColumnCount; ee++)
                    {

                        if (ee != 0 && ee != 1 && ee != 2)
                            this.dataGridView7.Columns[ee].Visible = false;
                    }
                    if (dataGridView7.Rows[i].Cells[4].Value.ToString() != "1")
                    {

                        dataGridView7.Rows[i].Visible = false;
                    }
                    else
                    {
                        dataGridView7.Rows[i].Visible = true;
                    }

                }
            }
            else if (tabControl2.SelectedIndex == 3)
            {
                for (int i = 0; i < this.dataGridView2.RowCount; i++)
                {
                    dataGridView2.CurrentCell = null;
                    for (int ee = 0; ee < this.dataGridView2.ColumnCount; ee++)
                    {

                        if (ee != 0 && ee != 1 && ee != 2)
                            this.dataGridView2.Columns[ee].Visible = false;
                    }
                    if ( dataGridView2.Rows[i].Cells[4].Value.ToString() != "2")
                    {

                        dataGridView2.Rows[i].Visible = false;
                    }
                    else
                    {
                        dataGridView2.Rows[i].Visible = true;
                    }
                }


                for (int i = 0; i < this.dataGridView7.RowCount; i++)
                {
                    dataGridView7.CurrentCell = null;
                    for (int ee = 0; ee < this.dataGridView7.ColumnCount; ee++)
                    {

                        if (ee != 0 && ee != 1 && ee != 2)
                            this.dataGridView7.Columns[ee].Visible = false;
                    }
                    if (dataGridView7.Rows[i].Cells[4].Value.ToString() != "2")
                    {

                        dataGridView7.Rows[i].Visible = false;
                    }
                    else
                    {
                        dataGridView7.Rows[i].Visible = true;
                    }

                }
            }
            else
            {
                for (int i = 0; i < this.dataGridView2.RowCount; i++)
                {
                    for (int ee = 0; ee < this.dataGridView2.ColumnCount; ee++)
                    {
                        if (ee != 0 && ee != 1 && ee != 2)
                            this.dataGridView2.Columns[ee].Visible = false;
                    }
                    this.dataGridView2.Rows[i].Visible = true;
                }


                for (int i = 0; i < this.dataGridView7.RowCount; i++)
                {
                    for (int ee = 0; ee < this.dataGridView7.ColumnCount; ee++)
                    {
                        if (ee != 0 && ee != 1 && ee != 2)
                            this.dataGridView7.Columns[ee].Visible = false;
                    }
                    this.dataGridView7.Rows[i].Visible = true;
                }
            }
        }
    }
}
