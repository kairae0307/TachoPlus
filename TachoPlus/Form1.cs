using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Net.Sockets;
using System.IO;
using System.IO.Ports;
using System.Net;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Data.SqlClient;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using Microsoft.Win32;
using System.Management;
using System.Reflection;
using System.Net.Mail;
using System.Net.Mime;
using System.Net.NetworkInformation;
using System.Globalization;
using System.Security.Cryptography;
using System.Security.AccessControl;


//It IS mandatory to include this reference
//using System.Runtime.InteropServices;

//
//   You can add any code you want…
//   HERE …..
//


namespace TachoPlus
{
   
    public partial class Form1 : Form
    {
        //1. Printable Data Handling APIs
        [DllImport("WoosimPrinter.DLL")]
        public static extern void ClearSpool();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void CompressedBmpSaveSpool(String bmpFilePath);
        [DllImport("WoosimPrinter.DLL")]
        public static extern int ControlCommand(byte[] Cmds, int iLength);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void DataMatrixSaveSpool(int width, int height, int module, String barcodeData);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void GS1DatabarSaveSpool(int type, int n, String barcodeData);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void LoadLogoSaveSpool(int n);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void MaxicodeSaveSpool(int mode, String barcodeData);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void MicroPDF417SaveSpool(int width, int column, int row, int ratio, String barcodeData, bool HRI);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void NormalBmpSaveSpool(String bmpFilePath);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void OneDimensionBarcodeSaveSpool(int ucBarcodeType,
                                                    int ucBarWidth, int ucBarHeight, bool HRI_On, string BarcodeData);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_DotFeed(int dots);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_DrawBox(int iXPos, int iYPos, int width, int height, int thickness);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_DrawEllipse(int iXPos, int iYPos, int A_lenght, int B_lenght, int thickness);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_DrawLine(int iXPos, int iYPos, int iX2Pos, int iY2Pos, int thickness);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_LineFeed(int lines);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_Newline();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_Print();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_Print_StandardMode();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void PDF417SaveSpool(int width, int column, int level, int ratio, String barcodeData, bool HRI);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void PrintData();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void PrintDotFeed(int dots);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void PrintLineFeed(int lines);
        [DllImport("WoosimPrinter.DLL")]
        public static extern int PrintSpool(bool bDelete_Spool);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void PrintSpoolForTTF(String sData, int iXFontSize, int iYFontSize);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void QRCodeSaveSpool(int version, char level, int module, String barcodeData);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void TextSaveSpool(String textData);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void TruncatedPDF417SaveSpool(int width, int column, int level, int ratio, String barcodeData, bool HRI);

        //2. Printer Mode & Setting APIs
        [DllImport("WoosimPrinter.DLL")]
        public static extern void GetPrinterModelName();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void GetFirmwareVersion();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void GetPrinterStatus(int iTimeoutMsec);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void InitLineSpace();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void InitPageMode(int iXPos, int iYPos, int width, int height);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void InitPrinterStatus();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_ClearCurrentData();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_SetArea(int iXPos, int iYPos, int width, int height);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_SetDirection(int n);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_SetPosition(int iXPos, int iYPos);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void Page_SetStandardMode();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void SetAbsPosition(int distance);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void SetCharCodeTable(int n, int MCU);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void SetCharSpace(int n);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void SetFontForTTF(String ttfFile);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void SetFontSize(int n);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void SetLineSpace(int n);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void SetPageMode();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void SetTextAlignment(int n);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void SetTextStyle(int underline, bool emphasize, int width, int height, bool reverse);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void SetUpsideDown(bool set);

        //3. Misc Device Handling APIs
        [DllImport("WoosimPrinter.DLL")]
        public static extern void CancelMSRMode();
        [DllImport("WoosimPrinter.DLL")]
        public static extern bool ClosePrinterConnection();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void CancelSCRMode();
        [DllImport("WoosimPrinter.DLL")]
        public static extern int ConnectSerialPrinter(String sPortName, int iBaudRate, int iTimeoutMsec, bool bProtocol);
        [DllImport("WoosimPrinter.DLL")]
        public static extern int ConnectUSBPrinter(int iTimeoutMsec, bool Protocol);
        [DllImport("WoosimPrinter.DLL")]
        public static extern int ConnectWirelessPrinter(String sIP_ADDR, int iPortNum, int iTimeoutMsec, bool bProtcol);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void CutPaper(int mode);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void EnterMSRMode(int n);
        [DllImport("WoosimPrinter.DLL")]
        public static extern void EnterSCRMode();
        [DllImport("WoosimPrinter.DLL")]
        public static extern void FeedToMark();

        [DllImport("user32.dll")]
        public static extern uint RegisterWindowMessage(string lpString);

        private uint UWM_RECEIVE_MSG = RegisterWindowMessage("WOOSIM_PRT_OK");

        public byte[] Daily_Distance = new byte[4];
        public double Money_Check = 0;



        string m_strText = "";
        int m_nNum = 0;
        bool m_bCheck = false;

        public bool Recipt_Using = false;
        public string Encrypt = "";
        public string Dencrypt = "";
        public int StartDay = 1;
        // Return Message        
        private const int SUCCESS = 1;
        private const int ALREADY_OPENED = -1;
        private const int NO_RESPONSE_FROM_PRINTER = -5;
        private const int TIMEOUT = -7;
        private const int NOT_OPEN_THE_PORT = -11;


        //Serial, Bluetooth Err message
        private const int UNABLE_TO_OPEN_THE_PORT = -2;
        private const int UNABLE_TO_CONFIGURE_THE_SERIAL_PORT = -3;
        private const int UNABLE_TO_SET_THE_TIMEOUT_PARAMETERS = -4;

        //Wireless Err message
        private const int SOCKET_ERROR = -2;
        private const int CONNECT_FAIL = -3;

        //USB Err message
        private const int UNABLE_TO_MAKE_CONNECTION = -2;
        private const int UNABLE_TO_GET_INFOMATION = -3;
        private const int NOT_WOOSIM_PRINTER = -4;
         //Thread  Debug Exception ignore	


       public  DateTime StartDate = new DateTime(2016, 11, 11, 0, 0, 0);
       public  DateTime EndDate = new DateTime(2016, 11, 11, 0, 0, 0);
       public bool GetTime_Receipt = false;


       public string  Print_Text = "";
        public bool MALAYSIA_Set = false;
        public bool UAE_Set = false;
        public bool THAILAND_Set = false;
        public bool MiniPrintCon = false;
        public bool MiniPrintEnable = false;
        public string MiniPrintPort = "";

        public bool SerialSend = false;
        public delegate void Add(string dir);
        public delegate void Add_a(TreeNode child);
        string name = "★ Tacho Data ★";
        delegate void LoadingCallBack(bool visible);
        LoginForm Login;
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct NETRESOURCE
        {
            public uint dwScope;
            public uint dwType;
            public uint dwDisplayType;
            public uint dwUsage;
            public string lpLocalName;
            public string lpRemoteName;
            public string lpComment;
            public string lpProvider;
        }

        [DllImport("mpr.dll", CharSet = CharSet.Auto)]
        public static extern int WNetUseConnection(
                    IntPtr hwndOwner,
                    [MarshalAs(UnmanagedType.Struct)] ref NETRESOURCE lpNetResource,
                    string lpPassword,
                    string lpUserID,
                    uint dwFlags,
                    StringBuilder lpAccessName,
                    ref int lpBufferSize,
                    out uint lpResult);

        public struct Total
        {
            public double Money;        
            public double Distance;
          
        }
        [STAThread]
        public void Loading_Visible(bool visible)
        {
            if (this.pictureBox17.InvokeRequired)
            {
                LoadingCallBack d = new LoadingCallBack(Loading_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox17.Visible = visible;
            }
        }     

        [StructLayout(LayoutKind.Sequential)]
        public struct DEV_BROADCAST_VOLUME
        {
            public int dbcv_size;
            public int dbcv_devicetype;
            public int dbcv_reserved;
            public int dbcv_unitmask;
        }
     //   public TMEX TMEXLibrary;
     /*   public short portNum;
        public short portType;
        public int hSess;
        public int sessionOptions;
        public byte[] stateBuffer = new byte[15360];						// internal for TMX interface
        public short[] ROM = new short[8];*/

       // public TransactionForm transactionform;
        public string formname = "";
        public bool LoginOK = false;
        public bool isBackup = false;
        public string MdbName = "";
        public string CashierID = "";
        public bool CashierMode = false;
        public string ShareIP = "";
        public bool iButtonMode = false;
        public bool ViewerMode = false;
        public bool UserLogin = false;
        public bool AdminLogin = false;
        public bool UserIdCreate = false;

        public string Taxi_ID_temp = "";
        public string Driver_ID_temp = "";
        public string Driver_Name_temp = "";

        public string Vacant_dist_temp = "";
        public string Hired_dist_temp = "";
        public string Total_dist_temp = "";
        public string Total_income_temp = "";
        public string Monthly_temp = "";

        public string Call_temp = "";
        public string Lugg_temp = "";
        public string AP_temp = "";
        public string Extra_temp = "";
        public string Toll_temp = "";


        public string ServerMessage = "";
        public int SelectItem = 0;
        public bool TransactionForm_Run = false;
        public bool PaidTimeWrite = false;
         public bool Ibutton_touched= false;
         public bool Ibutton_newButton= true;
         public string Hdd_Serial = "";
         public string mdb_Hdd_Serial = "";
         public byte[] TimLimitByte = new byte[2];

        // 검색 시작일
        public bool bSearchStartSel;
        public DateTime dtSearchStartDay;
        public bool bSearchStartAM;
        public int nSearchStartHour;
        // 검색 종료일
        public bool bSearchEndSel;
        public DateTime dtSearchEndDay;
        public string searchCarNum = "";
        public string searchDreverID = "";
        public string searchCashierID = "";


        public double SearchIncome = 0;

        private bool bConnectServer = false;
        private Socket EY_ChatClient = null;
        private int intPortNum = 3333;
        private int intSize = 0;
        private string strLog = "";
        private string strErr = "";
        //  private string strCurTime = "";
        private string strSendMsg = "";
        private string strReceiveMsg = "";
        private byte[] byteReceiveMsg = new byte[1024];
        private byte[] byteSendMsg = new byte[1024];
        public string ServerIP = "";
        public Total total;
        static public TcpServerSocket svr;
        public int listviewID = 0;
        public int recvTotal_Money = 0;
        public double recvTotal_Dist = 0;
        public double reevTotal_SalesDist = 0;

        public bool IbuttonSetting = false;
        public bool IbuttonReadCheck = false;
        public int sessionOptions;
        public byte[] stateBuffer = new byte[15360];						// internal for TMX interface
       // public bool dotcheck = false;
        public byte DotValue = 0;
        public bool ibutonCheck = false;
        public bool Serverstart = false;
        public bool TIME1_ENABLE = false;
        public bool TIME2_ENABLE = false;
        public string CurTime = "";
        public string Time1_str = "";
        public string Time2_str = "";
        public string Car_Num = "";                 
        public string CarSum_Path = "";
        public string CarSum_DBName = "";
        public string CarArea = "";
        public char CarSign;
        public string Amf_path = "";
        public string dirname = "";

        public string mdbfilename = "";
        public TimeSpan[] SalesTimeTotal;
        public DateTime InTImeBackup;
        public string CarNumber_backup = "";
        public int SavedTime = 0;
        public int e7cnt = 13;
        public int f9cnt = 0;
        public int oldf9cnt = 0;
        public int hh = 0;
        public int mm = 0;
        public int ss = 0;
        public int TimeCnt = 0;
        public int TimeCount = 0;
        public int LoadingTimeOut = 0;
        public int ClientTimeoutcnt1 = 0;
        public int ClientTimeoutcnt2 = 0;
        public int ClientTimeoutcnt3 = 0;
        public int ClientTimeoutcnt4 = 0;

     //   public bool Client1Check = false;
    //    public bool Client2Check = false;
    //    public bool Client3Check = false;
    //    public bool Client4Check = false;
        public bool[] ClientCheck = new bool[4];

        public bool TachoPass = false;
        public bool smf_data_chk = false;
        public bool NewFA = false;
        public bool SDcard_chk = false;
        public bool pasingEnd = false;
        public bool pasing1 = false;
        public bool pasing2 = false;
        public bool pasing3 = false;
        public bool pasing4 = false;
        public bool distCheck = false;
        public bool YearCheck = false;
        public bool empty_chk = false;
        public bool NandCheck = false;
        public bool ExternalCon = false;
        public bool tmpopen = false;
        public bool NewVersion = false;
        public bool Tacho_braek = false;
        public bool start_tree = false;
        public bool Second_tree = false;
        public bool third_tree = false;
        public bool E7Mark = false;
        public bool meterzero = false;
        public int timecnt = 0;
        public bool fillCheck = false;
        public bool Tacho_dailly = false;
        public bool TAcho_tong = false;
        public bool Tacho_2dailly = false;
        public bool Tacho_auto = false;
        public bool mdbfilesel = false;
        public bool nodeclick = false;
        public bool textboxdig = false;
        public bool FaFb = false;
        public bool auto_serial = false;
        public bool auto_Repeat = false;
        public bool Fare_0_Dist_0_Enble = false;
        // 세부 검색창 활성화 여부
        public bool bIsDetail = false;

        public int selectedColumnIndex = 4;
        public int selectedOrder = 1;
        public int nOpenedindex = 0;
        private int richTextCnt = 0;
        public int DataCnt = 0;
        public int rcvCnt = 0;
        public int linecnt = 0;
        public int pasingcnt = 0;
        public int TimeOut1 = 0;
        public int TimeOut2 = 0;
        public int TimeOut3 = 0;
        public int TimeOut4 = 0;

        public bool Data_Start = false;
        public bool Data_Start2 = false;
        public bool Data_Start3 = false;
        public bool Data_Start4 = false;

        public bool Data_rcving = false;
        public bool Data_rcving2 = false;
        public bool Data_rcving3 = false;
        public bool Data_rcving4 = false;



        public bool IRD_Enable = false;

        public bool ServerOpen = false;
        private Socket ServerSocket = null;
        private Thread ServerThread = null;

        public string TACHO2_path = "";
        public Socket TCPListnerAcceptSocket;

      
        private bool bSocketEnd = false;
   
        public int intCientID = 0;
      
     
        private string strCurTime = "";
        private string strClientID = "";
   
        public string xValue = "";
        public int ClientCnt = 0;

        private Socket ClientSocket = null;
        public Queue que = new Queue();
        public List<string> IDList = new List<string>();
       // public Hashtable ClientList = new Hashtable();

        private iniClass inicls = new iniClass();

        private List<byte> FileData = new List<byte>();
        private List<byte> Image1 = new List<byte>();
        private List<byte> Image2 = new List<byte>();
        private List<byte> SwapData = new List<byte>();

        delegate void textCallbak(String txt);
        delegate void progressCallBack(int nProgressBar);

        private List<byte> rcvList = new List<byte>();
        private List<byte> rcvList2 = new List<byte>();
        private List<byte> rcvList3 = new List<byte>();
        private List<byte> rcvList4 = new List<byte>();

        private List<byte> hisList = new List<byte>();

        public Thread Ibuttonthread;
        public Thread Sharethread; 
        public Thread getSerialData;
        public Thread getSerialData2;
        public Thread getSerialData3;
        public Thread getSerialData4;

        delegate void labelCallBack(bool visible);
        delegate void pictureCallBack(bool visible);

        public class NodeSorter : IComparer
        {
            // Compare the length of the strings, or the strings
            // themselves, if they are the same length.
            public int Compare(object x, object y)
            {
                TreeNode tx = x as TreeNode;
                TreeNode ty = y as TreeNode;

                // Compare the length of the strings, returning the difference.
                if (tx.Text.Length != ty.Text.Length)
                    return tx.Text.Length - ty.Text.Length;

                // If they are the same length, call Compare.
                return string.Compare(ty.Text, tx.Text);
            }
        }


        public struct TachoRamData
        {
            public int PremiumBasicDistance;
            public int PremiumAfterDistance;
            public int DriveBasicDistance;
            public int DriveAfterDistance;
            public int PremiumBasicMoney;
            public int PremiumAfterMoney;
            public int DriveBasicMoney;
            public int DriveAfterMoney;
            public int CallMoney;
            public int FreightMoney;
            public int DiscountMoney;
            public int TotalDriveDistance;
            public int TotalTradeDistance;
            public int TodayIncomeMoney;
            public int DistanceBy1Pulse;
            public int TodayTotalDriveDistance;
            public int TodayTotalTradeDistance;
            public DateTime InWarehouseTime;
            public DateTime OutWarehouseTime;
            public int RealIncomeMoney;
            public int DriverNumber;
            public double Fuel;
            public string CarNumber;
            public ushort _start;
            public ushort _size;
            public ushort _pointer;
            public byte _overflag;
            public int VersionNumber;
            public int TachoSavedTime;
            public int OffSize;
            public int TVersion;
            public int NandPiece;
            public int Nandsize;
        }

        public struct TachoDataCode
        {
            public int moneyTblTacho;
            public double salesKmTblTacho;
            public double driveDistanceTblTacho;
            public DateTime overrunTime;
            public int emerBreakCnt;
            public int driveBasicCnt;
            public int driveAfterCnt;
            public int premiumBasicCnt;
            public int premiumAfterCnt;
            public int doorOpenCnt;

            public int driveCount;
            public DateTime yymmdd;
            public DateTime beforeTime;
            public DateTime afterTime;
            public double salesKm;
            public int money;
            public int SalesTotalMoney;
            public double empty;
            public int emptyTime;
            public bool notuse;
            public bool add;
            public bool key;
            public bool emerBreak;
            public DateTime emerTime;
            public DateTime emptyStartTime;
            public int emerSpeed;

            // '10. 7.19 추가
            public int celldriveBasicCnt;
            public int celldriveAfterCnt;
            public int cellpremiumBasicCnt;
            public int cellpremiumAfterCnt;
            public int cellsalesCnt;
            public int cellsalesTime;
            public int cellcarEmptyTime;
            public int cellkeyUseCnt;
            public int cellemerBreakCnt;
            public DateTime celloverrunTime;
            //

            public bool MKdoor;
            public double TotalDriveDistanceSaved;
            public double TotalDriveDistance;
            public byte speed;
            public double distance;
            public bool sales;
            public bool engine;

            public int salesCnt;        // 영업회수
            public int carEmptyTime;    // 공차시간
            public int keyUseCnt;       // 키사용회수
            public int salesTime;       // 영업시간
        }
        #region Calc_Func
        private int BcdToDecimalByLsb(byte[] arr, int cnt)
        {
            int rtValue = 0, mulValue = 0;

            for (int i = 0; i < cnt; i++)
            {
                if (i == 0) mulValue = 1;
                else if (i == 1) mulValue = 100;
                else if (i == 2) mulValue = 10000;
                else if (i == 3) mulValue = 1000000;
                else mulValue = 0;

                rtValue += (((arr[i] >> 4) * 10) + (arr[i] & 0x0F)) * mulValue;
            }

            return rtValue;
        }

        private int BcdToDecimal(byte bTemp)
        {
            return (((bTemp >> 4) * 10) + (bTemp & 0x0F));
        }

        int BinToBcd8(int n)
        {
            return ((int)((n / 10) << 4) + (n % 10));
        }

        public byte[] BinToBcd24P(byte[] arr, int n)
        {

            int rulTmp;

            rulTmp = n % 1000000;
            arr[2] = (byte)BinToBcd8(rulTmp / 10000);
            rulTmp = n % 10000;
            arr[1] = (byte)BinToBcd8(rulTmp / 100);
            arr[0] = (byte)BinToBcd8(rulTmp % 100);
            return arr;
        }




        [STAThread]
        public void client1_Visible(bool visible)
        {
            if (this.Client1_Icon.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(client1_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.Client1_Icon.Visible = visible;
            }
        }
        [STAThread]
        public void client2_Visible(bool visible)
        {
            if (this.Client2_Icon.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(client2_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.Client2_Icon.Visible = visible;
            }
        }
        [STAThread]
        public void client3_Visible(bool visible)
        {
            if (this.Client3_Icon.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(client3_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.Client3_Icon.Visible = visible;
            }
        }
        [STAThread]
        public void client4_Visible(bool visible)
        {
            if (this.Client4_Icon.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(client4_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.Client4_Icon.Visible = visible;
            }
        }     


        [STAThread]
        private void pictureBox2_Visible(bool visible)
        {
            if (this.pictureBox2.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox2_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox2.Visible = visible;
            }
        }


        [STAThread]
        private void pictureBox1_Visible(bool visible)
        {
            if (this.pictureBox1.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox1_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox1.Visible = visible;
            }
        }


        [STAThread]
        private void pictureBox3_Visible(bool visible)
        {
            if (this.pictureBox3.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox3_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox3.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox4_Visible(bool visible)
        {
            if (this.pictureBox4.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox4_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox4.Visible = visible;
            }
        }
        [STAThread]
        private void pictureBox5_Visible(bool visible)
        {
            if (this.pictureBox5.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox5_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox5.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox6_Visible(bool visible)
        {
            if (this.pictureBox6.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox6_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox6.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox7_Visible(bool visible)
        {
            if (this.pictureBox7.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox7_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox7.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox8_Visible(bool visible)
        {
            if (this.pictureBox8.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox8_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox8.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox9_Visible(bool visible)
        {
            if (this.pictureBox9.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox9_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox9.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox10_Visible(bool visible)
        {
            if (this.pictureBox10.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox10_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox10.Visible = visible;
            }
        }



        [STAThread]
        private void pictureBox11_Visible(bool visible)
        {
            if (this.pictureBox11.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox11_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox11.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox12_Visible(bool visible)
        {
            if (this.pictureBox12.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox12_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox12.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox13_Visible(bool visible)
        {
            if (this.pictureBox13.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox13_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox13.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox14_Visible(bool visible)
        {
            if (this.pictureBox14.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox14_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox14.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox15_Visible(bool visible)
        {
            if (this.pictureBox15.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox15_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox15.Visible = visible;
            }
        }

        [STAThread]
        private void pictureBox16_Visible(bool visible)
        {
            if (this.pictureBox16.InvokeRequired)
            {
                pictureCallBack d = new pictureCallBack(pictureBox16_Visible);
                this.Invoke(d, new object[] { visible });
            }
            else
            {
                this.pictureBox16.Visible = visible;
            }
        }


        #endregion


        public Form1()
        {
            InitializeComponent();
            m_list = new PrintableListView.PrintableListView();
            CheckForIllegalCrossThreadCalls = false; 
          //  transactionform = new TransactionForm(this);
            total = new Total();
         //   TMEXLibrary = new TMEX();
        }
        private static string EncryptString(string InputText, string Password)
        {

            // Rihndael class를 선언하고, 초기화
            RijndaelManaged RijndaelCipher = new RijndaelManaged();

            // 입력받은 문자열을 바이트 배열로 변환
            byte[] PlainText = System.Text.Encoding.Unicode.GetBytes(InputText);

            // 딕셔너리 공격을 대비해서 키를 더 풀기 어렵게 만들기 위해서 
            // Salt를 사용한다.
            byte[] Salt = Encoding.ASCII.GetBytes(Password.Length.ToString());

            // PasswordDeriveBytes 클래스를 사용해서 SecretKey를 얻는다.
            PasswordDeriveBytes SecretKey = new PasswordDeriveBytes(Password, Salt);

            // Create a encryptor from the existing SecretKey bytes.
            // encryptor 객체를 SecretKey로부터 만든다.
            // Secret Key에는 32바이트
            // (Rijndael의 디폴트인 256bit가 바로 32바이트입니다)를 사용하고, 
            // Initialization Vector로 16바이트
            // (역시 디폴트인 128비트가 바로 16바이트입니다)를 사용한다.
            ICryptoTransform Encryptor = RijndaelCipher.CreateEncryptor(SecretKey.GetBytes(32), SecretKey.GetBytes(16));

            // 메모리스트림 객체를 선언,초기화 
            MemoryStream memoryStream = new MemoryStream();

            // CryptoStream객체를 암호화된 데이터를 쓰기 위한 용도로 선언
            CryptoStream cryptoStream = new CryptoStream(memoryStream, Encryptor, CryptoStreamMode.Write);

            // 암호화 프로세스가 진행된다.
            cryptoStream.Write(PlainText, 0, PlainText.Length);

            // 암호화 종료
            cryptoStream.FlushFinalBlock();

            // 암호화된 데이터를 바이트 배열로 담는다.
            byte[] CipherBytes = memoryStream.ToArray();

            // 스트림 해제
            memoryStream.Close();
            cryptoStream.Close();

            // 암호화된 데이터를 Base64 인코딩된 문자열로 변환한다.
            string EncryptedData = Convert.ToBase64String(CipherBytes);

            // 최종 결과를 리턴
            return EncryptedData;
        }
        private static string DecryptString(string InputText, string Password)
        {
            RijndaelManaged RijndaelCipher = new RijndaelManaged();

            byte[] EncryptedData = Convert.FromBase64String(InputText);
            byte[] Salt = Encoding.ASCII.GetBytes(Password.Length.ToString());

            PasswordDeriveBytes SecretKey = new PasswordDeriveBytes(Password, Salt);

            // Decryptor 객체를 만든다.
            ICryptoTransform Decryptor = RijndaelCipher.CreateDecryptor(SecretKey.GetBytes(32), SecretKey.GetBytes(16));

            MemoryStream memoryStream = new MemoryStream(EncryptedData);

            // 데이터 읽기(복호화이므로) 용도로 cryptoStream객체를 선언, 초기화
            CryptoStream cryptoStream = new CryptoStream(memoryStream, Decryptor, CryptoStreamMode.Read);

            // 복호화된 데이터를 담을 바이트 배열을 선언한다.
            // 길이는 알 수 없지만, 일단 복호화되기 전의 데이터의 길이보다는
            // 길지 않을 것이기 때문에 그 길이로 선언한다.
            byte[] PlainText = new byte[EncryptedData.Length];

            // 복호화 시작
            int DecryptedCount = cryptoStream.Read(PlainText, 0, PlainText.Length);

            memoryStream.Close();
            cryptoStream.Close();

            // 복호화된 데이터를 문자열로 바꾼다.
            string DecryptedData = Encoding.Unicode.GetString(PlainText, 0, DecryptedCount);

            // 최종 결과 리턴
            return DecryptedData;
        }
     

        private void buttonOpen_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                    serialPort1.Close();


            }
            catch
            {
                MessageBox.Show("Can't close already opened port", "Error");
                return;
            }

            serialPort1.PortName = comboBoxPort.SelectedItem.ToString();
            serialPort1.BaudRate = Convert.ToInt32(comboBoxBaud.SelectedItem);

            try
            {
                serialPort1.Open();
  
                label13.Text = comboBoxPort.SelectedItem.ToString() + " Open";
               // if (IRD_Enable == true)
              //  {
                    getSerialData = new Thread(new ThreadStart(Run_SerialThread));
                    getSerialData.IsBackground = true;
                    Thread.Sleep(100);
                    getSerialData.Start();
             //   }
              
              
            }
            catch
            {
                MessageBox.Show("Can't open port", "Error");
                return;
            }
        }
        public void Run_SerialThread()
        {
            byte status = 0;
            byte bInData;
            string txt;
            while (serialPort1.IsOpen)
            {
             
                try
                {
                    switch (status)
                    {
                        case 0:
                                  
                                    byte[] Data = new byte[6];

                                      Data[0] = 0x55;
                                      Data[1] = 0xA6;
                                  //  Data[2] = 0x00;
                                  //  Data[3] = 0xcc;
                                  //  Data[4] = 0xfd;
                                  //  Data[5] = 0x00;

                               

                                    pictureBox9_Visible(true);
                                    pictureBox10_Visible(false);
                                        serialPort1.DiscardInBuffer();
                                        serialPort1.DiscardOutBuffer();
                                        rcvList.Clear();

                                //    pictureBox2_Visible(false);
                                //    pictureBox1_Visible(true);
                                    if (IRD_Enable == true || SerialSend ==true)
                                    {
                                     
                                        SerialSend = false;
                                        serialPort1.Write(Data, 0, 2);
                                     //   status = 1;
                                    }
                                 //  label4.ForeColor = Color.Black;
                                 //  label4.Text = "NORMAL";

                                //  Thread.Sleep(10);
                               //     
                                  status++;
                                  TimeOut1 = 70;
                          
                         
                            break;
                        case 1:
                          
                            TimeOut1--;

                            if (TimeOut1 == 0)
                            {
                                status = 0;
                                rcvList.Clear();
                                Data_Start = false;
                            }
                            Data_rcving = true;
                            
                            for (int g = 0; g < 100; g++)
                           {
                                if (serialPort1.IsOpen)
                                {
                                    if (serialPort1.BytesToRead > 0)
                                    {

                                      
                                        bInData = (byte)serialPort1.ReadByte();
                                        txt = String.Format("{0:X2} ", bInData);

                                       

                                        if (bInData == 0xA3)
                                        {
                                            if (Data_Start == false)
                                            {
                                                Data_Start = true;
                                            //    rcvList.Clear();
                                            }
                                        }


                                        if (Data_Start == true)
                                        {  
                                           // label1_Ani_0_Visible(true);
                                           // label1_Base_Visible(false);

                                            pictureBox9_Visible(false);
                                            pictureBox10_Visible(true);

                                            TimeOut1 = 70;
                                            rcvList.Add(bInData);
                                          //  label4.ForeColor = Color.Red;
                                          //  label4.Text = "Data Receive";
                                         //   richTextCnt++;
                                            //  txt = String.Format("{0:X2} ", bInData);
                                        }
                                        else
                                        {
                                            Data_Start = false;
                                            label4.ForeColor = Color.Black;
                                            label4.Text = "NORMAL";
                                        }


                                        if (rcvList.Count >= 256+64)
                                        {
                                            if (rcvList[rcvList.Count - 2] == 0xfd && rcvList[0] == 0xA3)
                                            {

                                                try
                                                {
                                                    Data_Start = false;
                                                    //   label3.Text = String.Format("{0:X2} ", rcvList[rcvList.Count - 1]);
                                                    byte[] Data_ = new byte[1];

                                                    if (rcvList.Count != 0)
                                                    {

                                                        Data_[0] = rcvList[rcvList.Count-1];
                                                    //    rcvList.RemoveAt(rcvList.Count - 1);
                                                    }
                                                    else
                                                    {
                                                       
                                                        label4.Text = "NORMAL";
                                                        status = 0;
                                                        rcvList.Clear();
                                                    }

                                                    byte CheckSum = 0;
                                                    for (int a = 0; a < rcvList.Count - 1; a++)
                                                    {
                                                        CheckSum += rcvList[a];
                                                    }



                                                    if (CheckSum == Data_[0])
                                                    {



                                                        if (serialPort1.IsOpen)
                                                        {

                                                             Data = new byte[1];

                                                            Data[0] = CheckSum;
                                                       

                                                            serialPort1.DiscardOutBuffer();
                                                            serialPort1.DiscardInBuffer();
                                                            serialPort1.Write(Data, 0, Data.Length);
                                                        }
                                                        //////////////////////////////// 한개 짜리 tmf 저장 /////////////////////////////////////////

                                                        string NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}",
                                                                                                   (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                                                        string carnum = "";
                                                        string Model = "";

                                                        for (int i = 180; i <189; i++)
                                                        {
                                                            if (rcvList[i] < 0x20)
                                                            {
                                                                rcvList[i] = 0x20;
                                                            }
                                                        }

                                                        if (rcvList.Count > 256+64)
                                                        {
                                                            carnum = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}{8:C2}"
                                                                                     ,Convert.ToChar( rcvList[188]), Convert.ToChar(rcvList[187]), Convert.ToChar(rcvList[186]), Convert.ToChar(rcvList[185])
                                                                                     ,Convert.ToChar( rcvList[184]), Convert.ToChar(rcvList[183])
                                                                                  , Convert.ToChar(rcvList[182]), Convert.ToChar(rcvList[181]), Convert.ToChar(rcvList[180]));

                                                        

                                                            Model = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}"
                                                                                  , Convert.ToChar(rcvList[240]), Convert.ToChar(rcvList[241]), Convert.ToChar(rcvList[242]), Convert.ToChar(rcvList[243]), Convert.ToChar(rcvList[244]));
                 
                                                           //  strMN += String.Format("{0:C}", Convert.ToChar(newProInHeader.ModelName[i]));
                                                          //  strMN += String.Format("{0:C}", Convert.ToChar(newProInHeader.ModelName[i]));
                                                        }
                                                  
                                                      //  string TmpFile = Application.StartupPath + "\\" + NowReceiveTime + ".TMF";


                                                        string newPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "AMF");
                                                        // Create the subfolder
                                                        System.IO.Directory.CreateDirectory(newPath);

                                                        string TmpFile = TACHO2_path + "\\TMF\\AMF\\" + Model + "_" + NowReceiveTime + "_" + carnum + ".AMF";

                                                        Amf_path = TmpFile;

                                                        byte[] rcvByte = new byte[rcvList.Count];
                                                        rcvList.CopyTo(rcvByte);

                                                        FileStream fs = new FileStream(TmpFile, FileMode.OpenOrCreate, FileAccess.Write);
                                                        BinaryWriter bw = new BinaryWriter(fs);
                                                        
                                                        bw.Write(rcvByte);

                                                        serialPort1.DiscardInBuffer();
                                                        serialPort1.DiscardOutBuffer();
                                                       // serialPort1.Write(Data_, 0, Data_.Length);

                                                        fs.Close();
                                                        bw.Close();
                                                        AMF_Data(TmpFile);
                                                     //   Tacho_Run(TmpFile);

                                                        if (svr != null)
                                                        {

                                                            svr.TachoSend1 = true;
                                                        }

                                                        //////////////////////////////////////   2번째 저장 total tmf  기존 데이터와  이어 붙히자 !!///////////////////

                                                        string TMFPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "TransData");
                                                        // Create the subfolder
                                                        System.IO.Directory.CreateDirectory(TMFPath);

                                                        NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                                                                (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day);
                                                        TmpFile = TACHO2_path + "\\TMF\\TransData\\" + NowReceiveTime + ".AMF";

                                                       // rcvList.RemoveAt(0);
                                                        rcvByte = new byte[rcvList.Count];
                                                        rcvList.CopyTo(rcvByte);

                                                          fs = new FileStream(TmpFile, FileMode.Append, FileAccess.Write);                                                     
                                                          bw = new BinaryWriter(fs);

                                                          bw.Write(rcvByte);
                                                          fs.Close();
                                                          bw.Close();

                                                                      
                                                    }
                                                    else
                                                    {
                                                        Data_Start = false;
                                                        label4.Text = "NORMAL";
                                                        richTextCnt = 0;
                                                        linecnt = 0;
                                                        // richTextBoxMsg.Clear();

                                                        status = 0;
                                                        rcvList.Clear();
                                                        Data_Start = false;
                                                        serialPort1.DiscardInBuffer();
                                                        serialPort1.DiscardOutBuffer();
                                                       // MessageBox.Show("CheckSum Error");

                                                        Data = new byte[2];

                                                    //    Data[0] = 0x55;
                                                     //   Data[1] = 0xA5;

                                                      //  serialPort1.DiscardOutBuffer();
                                                      //  serialPort1.DiscardInBuffer();
                                                      //  serialPort1.Write(Data, 0, Data.Length);
                                                    }


                                                        richTextCnt = 0;
                                                        linecnt = 0;
                                                        // richTextBoxMsg.Clear();
                                                     
                                                        status = 0;
                                                        rcvList.Clear();
                                                        Data_Start = false;


                                                  
                                                    



                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                    string path = Application.StartupPath + "\\ErrorLog.jie";
                                                    using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                                                    {
                                                        sw.WriteLine("Run_SerialThread_TMFmake :[" + DateTime.Now.ToString() + "] " + ex.StackTrace);
                                                    }
                                                }
                                                pictureBox9_Visible(true);
                                                pictureBox10_Visible(false);


                                                pictureBox2_Visible(false);
                                                pictureBox1_Visible(true);


                                                Data_Start = false;
                                                label4.ForeColor = Color.Black;
                                                label4.Text = "NORMAL";


                                             //   AMF_Data(Amf_path);

                                            }

                                            
                                        }

                                        //     richTextBoxMsg.AppendText(txt);



                                    }
                                }
                            }
                          
                            Thread.Sleep(10);
                        
                            break;
                        case 2:
                          
                            break;
                    } //switch(status)
                }
                catch (Exception ex)
                {
                 //   MessageBox.Show(ex.Message);
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    string path = Application.StartupPath + "\\ErrorLog.jie";
                    using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                    {
                        sw.WriteLine("Run_SerialThread :[" + DateTime.Now.ToString() + "] " + ex.StackTrace);
                    }
                }
            }

        }

        public void Run_SerialThread2()
        {
            byte status = 0;
            byte bInData;
            string txt;
            while (serialPort2.IsOpen)
            {
                try
                {
                    switch (status)
                    {
                        case 0:
                            serialPort2.DiscardInBuffer();
                            serialPort2.DiscardOutBuffer();
                            byte[] Data = new byte[6];

                            Data[0] = 0x55;
                            Data[1] = 0x15;
                            Data[2] = 0x00;
                            Data[3] = 0xcc;
                            Data[4] = 0xfd;
                            Data[5] = 0x00;

                           


                            pictureBox11_Visible(true);
                            pictureBox12_Visible(false);

                            pictureBox4_Visible(false);
                            pictureBox3_Visible(true);

                            serialPort2.Write(Data, 0, Data.Length);
                            label10.ForeColor = Color.Black;
                            label10.Text = "NORMAL";

                            //  Thread.Sleep(50);
                            status++;
                            TimeOut2 = 50;

                            break;
                        case 1:

                            TimeOut2--;

                            if (TimeOut2 == 0)
                            {
                                status = 0;
                                rcvList2.Clear();
                                Data_Start2 = false;
                            }
                            Data_rcving2 = true;

                            for (int g = 0; g < 100; g++)
                            {
                                if (serialPort2.IsOpen)
                                {
                                    if (serialPort2.BytesToRead > 0)
                                    {


                                        bInData = (byte)serialPort2.ReadByte();
                                        txt = String.Format("{0:X2} ", bInData);



                                        if (bInData == 0xA3)
                                        {
                                            if (Data_Start2 == false)
                                            {
                                                Data_Start2 = true;
                                                rcvList2.Clear();
                                            }
                                        }


                                        if (Data_Start2 == true)
                                        {
                                          

                                            pictureBox11_Visible(false);
                                            pictureBox12_Visible(true);

                                            TimeOut2 = 50;
                                            rcvList2.Add(bInData);
                                            label10.ForeColor = Color.Red;
                                            label10.Text = "Data Receive";

                                          
                                            //  txt = String.Format("{0:X2} ", bInData);
                                        }
                                        else
                                        {
                                            Data_Start2 = false;
                                            label10.ForeColor = Color.Black;
                                            label10.Text = "NORMAL";
                                        }


                                        if (rcvList2.Count > 15)
                                        {
                                            if (rcvList2[rcvList2.Count - 2] == 0xfd && rcvList2[0] == 0xA3)
                                            {

                                                try
                                                {
                                                    Data_Start2 = false;
                                                    //   label3.Text = String.Format("{0:X2} ", rcvList[rcvList.Count - 1]);
                                                    byte[] Data_ = new byte[1];





                                                    if (rcvList2.Count != 0)
                                                    {
                                                        Data_[0] = rcvList2[rcvList2.Count - 1];
                                                        rcvList2.RemoveAt(rcvList2.Count - 1);
                                                    }
                                                    else
                                                    {
                                                        label10.Text = "NORMAL";
                                                        status = 0;
                                                        rcvList2.Clear();
                                                    }

                                                    byte CheckSum = 0;

                                                    for (int q = 385; q < rcvList2.Count; q++)
                                                    {
                                                        CheckSum += rcvList2[q];
                                                    }



                                                    if (CheckSum == Data_[0])
                                                    {

                                                        //////////////////////////////// 한개 짜리 tmf 저장 

                                                        string NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}",
                                                                                                   (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

                                                           string carnum = "";
                                                           if (rcvList2.Count > 384)
                                                           {
                                                               carnum = String.Format("{0:X2}-{1:X2}{2:X2}", rcvList2[336], rcvList2[335], rcvList2[334]);
                                                           }

                                                        //  string TmpFile = Application.StartupPath + "\\" + NowReceiveTime + ".TMF";
                                                           string newPath = System.IO.Path.Combine(TACHO2_path + "\\TMF_1", "AMF");
                                                           // Create the subfolder
                                                           System.IO.Directory.CreateDirectory(newPath);

                                                        string TmpFile = TACHO2_path + "\\TMF_1\\auto\\" + NowReceiveTime + "_" + carnum + ".TMF";


                                                        byte[] rcvByte = new byte[rcvList2.Count];
                                                        rcvList2.CopyTo(rcvByte);

                                                        FileStream fs = new FileStream(TmpFile, FileMode.OpenOrCreate, FileAccess.Write);
                                                        BinaryWriter bw = new BinaryWriter(fs);

                                                        bw.Write(rcvByte);

                                                        serialPort2.DiscardInBuffer();
                                                        serialPort2.DiscardOutBuffer();
                                                        serialPort2.Write(Data_, 0, Data_.Length);

                                                        fs.Close();
                                                        bw.Close();
                                                        Tacho_Run(TmpFile);
                                                        svr.TachoSend2 = true;

                                                        //////////////////////////////////////   2번째 저장 total tmf  기존 데이터와  이어 붙히자 !!

                                                        string TMFPath = System.IO.Path.Combine(TACHO2_path + "\\TMF_1", "TransData");
                                                        // Create the subfolder
                                                        System.IO.Directory.CreateDirectory(TMFPath);

                                                        NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                                                                (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day);
                                                        TmpFile = TACHO2_path + "\\TMF_1\\TransData\\" + NowReceiveTime + ".TMF";

                                                        rcvList2.RemoveAt(0);
                                                        rcvByte = new byte[rcvList2.Count];
                                                        rcvList2.CopyTo(rcvByte);

                                                        fs = new FileStream(TmpFile, FileMode.Append, FileAccess.Write);
                                                        bw = new BinaryWriter(fs);

                                                        bw.Write(rcvByte);
                                                        fs.Close();
                                                        bw.Close();


                                                    }
                                                    else
                                                    {
                                                        Data_Start2 = false;
                                                        label10.Text = "NORMAL";
                                                    }


                                                   

                                                    status = 0;
                                                    rcvList2.Clear();
                                                    Data_Start2 = false;



                                                    status = 0;



                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                    string path = Application.StartupPath + "\\ErrorLog.jie";
                                                    using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                                                    {
                                                        sw.WriteLine("Run_SerialThread2_TMFmake :[" + DateTime.Now.ToString() + "] " + ex.StackTrace);
                                                    }
                                                }

                                              


                                            }
                                        }

                                        //     richTextBoxMsg.AppendText(txt);



                                    }
                                }
                            }
                          //  Thread.Sleep(50);

                            break;
                        case 2:

                            break;
                    } //switch(status)
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    string path = Application.StartupPath + "\\ErrorLog.jie";
                    using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                    {
                        sw.WriteLine("Run_SerialThread2 :[" + DateTime.Now.ToString() + "] " + ex.StackTrace);
                    }
                }
            }

        }

        public void Run_SerialThread3()
        {
            byte status = 0;
            byte bInData;
            string txt;
            while (serialPort3.IsOpen)
            {
                try
                {
                    switch (status)
                    {
                        case 0:
                            serialPort3.DiscardInBuffer();
                            serialPort3.DiscardOutBuffer();
                            byte[] Data = new byte[6];

                            Data[0] = 0x55;
                            Data[1] = 0x15;
                            Data[2] = 0x00;
                            Data[3] = 0xcc;
                            Data[4] = 0xfd;
                            Data[5] = 0x00;

                            pictureBox13_Visible(false);
                            pictureBox14_Visible(true);

                            pictureBox6_Visible(false);
                            pictureBox5_Visible(true);

                            serialPort3.Write(Data, 0, Data.Length);
                            label19.ForeColor = Color.Black;
                            label19.Text = "NORMAL";

                            //  Thread.Sleep(50);
                            status++;
                            TimeOut3 = 50;

                            break;
                        case 1:

                            TimeOut3--;

                            if (TimeOut3 == 0)
                            {
                                status = 0;
                                rcvList3.Clear();
                                Data_Start3 = false;
                            }
                            Data_rcving3 = true;

                            for (int g = 0; g < 100; g++)
                            {
                                if (serialPort3.IsOpen)
                                {
                                    if (serialPort3.BytesToRead > 0)
                                    {


                                        bInData = (byte)serialPort3.ReadByte();
                                        txt = String.Format("{0:X2} ", bInData);



                                        if (bInData == 0xA3)
                                        {
                                            if (Data_Start3 == false)
                                            {
                                                Data_Start3 = true;
                                                rcvList3.Clear();
                                            }
                                        }


                                        if (Data_Start3 == true)
                                        {
                                            pictureBox13_Visible(true);
                                            pictureBox14_Visible(false);

                                            TimeOut3 = 50;
                                            rcvList3.Add(bInData);
                                            label19.ForeColor = Color.Red;
                                            label19.Text = "Data Receive";


                                            //  txt = String.Format("{0:X2} ", bInData);
                                        }
                                        else
                                        {
                                            Data_Start3 = false;
                                            label19.ForeColor = Color.Black;
                                            label19.Text = "NORMAL";
                                        }


                                        if (rcvList3.Count > 15)
                                        {
                                            if (rcvList3[rcvList3.Count - 2] == 0xfd && rcvList3[0] == 0xA3)
                                            {

                                                try
                                                {
                                                    Data_Start3 = false;
                                                    //   label3.Text = String.Format("{0:X2} ", rcvList[rcvList.Count - 1]);
                                                    byte[] Data_ = new byte[1];





                                                    if (rcvList3.Count != 0)
                                                    {
                                                        Data_[0] = rcvList3[rcvList3.Count - 1];
                                                        rcvList3.RemoveAt(rcvList3.Count - 1);
                                                    }
                                                    else
                                                    {
                                                        label19.Text = "NORMAL";
                                                        status = 0;
                                                        rcvList3.Clear();
                                                    }

                                                    byte CheckSum = 0;

                                                    for (int q = 385; q < rcvList3.Count; q++)
                                                    {
                                                        CheckSum += rcvList3[q];
                                                    }



                                                    if (CheckSum == Data_[0])
                                                    {

                                                        //////////////////////////////// 한개 짜리 tmf 저장 

                                                        string NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}",
                                                                                                   (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

                                                           string carnum = "";
                                                           if (rcvList3.Count > 384)
                                                           {
                                                               carnum = String.Format("{0:X2}-{1:X2}{2:X2}", rcvList3[336], rcvList3[335], rcvList3[334]);
                                                           }

                                                        //  string TmpFile = Application.StartupPath + "\\" + NowReceiveTime + ".TMF";
                                                           string newPath = System.IO.Path.Combine(TACHO2_path + "\\TMF_2", "AMF");
                                                           // Create the subfolder
                                                           System.IO.Directory.CreateDirectory(newPath);

                                                        string TmpFile = TACHO2_path + "\\TMF_2\\AMF\\" + NowReceiveTime + "_" + carnum + ".TMF";


                                                        byte[] rcvByte = new byte[rcvList3.Count];
                                                        rcvList3.CopyTo(rcvByte);

                                                        FileStream fs = new FileStream(TmpFile, FileMode.OpenOrCreate, FileAccess.Write);
                                                        BinaryWriter bw = new BinaryWriter(fs);

                                                        bw.Write(rcvByte);

                                                        serialPort3.DiscardInBuffer();
                                                        serialPort3.DiscardOutBuffer();
                                                        serialPort3.Write(Data_, 0, Data_.Length);

                                                        fs.Close();
                                                        bw.Close();
                                                        Tacho_Run(TmpFile);
                                                        svr.TachoSend3 = true;

                                                        //////////////////////////////////////   2번째 저장 total tmf  기존 데이터와  이어 붙히자 !!



                                                        string TMFPath = System.IO.Path.Combine(TACHO2_path + "\\TMF_2", "TransData");
                                                        // Create the subfolder
                                                        System.IO.Directory.CreateDirectory(TMFPath);

                                                        NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                                                                (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day);
                                                        TmpFile = TACHO2_path + "\\TMF_2\\TransData\\" + NowReceiveTime + ".TMF";

                                                        rcvList3.RemoveAt(0);
                                                        rcvByte = new byte[rcvList3.Count];
                                                        rcvList3.CopyTo(rcvByte);

                                                        fs = new FileStream(TmpFile, FileMode.Append, FileAccess.Write);
                                                        bw = new BinaryWriter(fs);

                                                        bw.Write(rcvByte);
                                                        fs.Close();
                                                        bw.Close();


                                                    }
                                                    else
                                                    {
                                                        Data_Start3 = false;
                                                        label19.Text = "NORMAL";
                                                    }




                                                    status = 0;
                                                    rcvList3.Clear();
                                                    Data_Start3 = false;



                                                    status = 0;



                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                    string path = Application.StartupPath + "\\ErrorLog.jie";
                                                    using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                                                    {
                                                        sw.WriteLine("Run_SerialThread3_TMFmake :[" + DateTime.Now.ToString() + "] " + ex.StackTrace);
                                                    }
                                                }

                                              


                                            }
                                        }

                                        //     richTextBoxMsg.AppendText(txt);



                                    }
                                }
                            }
                          //  Thread.Sleep(50);

                            break;
                        case 2:

                            break;
                    } //switch(status)
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    string path = Application.StartupPath + "\\ErrorLog.jie";
                    using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                    {
                        sw.WriteLine("Run_SerialThread3 :[" + DateTime.Now.ToString() + "] " + ex.StackTrace);
                    }
                }
            }

        }

        public void Run_SerialThread4()
        {
            byte status = 0;
            byte bInData;
            string txt;
            while (serialPort4.IsOpen)
            {
                try
                {
                    switch (status)
                    {
                        case 0:
                            serialPort4.DiscardInBuffer();
                            serialPort4.DiscardOutBuffer();
                            byte[] Data = new byte[6];

                            Data[0] = 0x55;
                            Data[1] = 0x15;
                            Data[2] = 0x00;
                            Data[3] = 0xcc;
                            Data[4] = 0xfd;
                            Data[5] = 0x00;

                            pictureBox15_Visible(false);
                            pictureBox16_Visible(true);

                            pictureBox8_Visible(false);
                            pictureBox7_Visible(true);

                            serialPort4.Write(Data, 0, Data.Length);
                            label27.ForeColor = Color.Black;
                            label27.Text = "NORMAL";

                            //  Thread.Sleep(50);
                            status++;
                            TimeOut4 = 50;

                            break;
                        case 1:

                            TimeOut4--;

                            if (TimeOut4 == 0)
                            {
                                status = 0;
                                rcvList4.Clear();
                                Data_Start4 = false;
                            }
                            Data_rcving4 = true;

                            for (int g = 0; g < 100; g++)
                            {
                                if (serialPort4.IsOpen)
                                {
                                    if (serialPort4.BytesToRead > 0)
                                    {


                                        bInData = (byte)serialPort4.ReadByte();
                                        txt = String.Format("{0:X2} ", bInData);



                                        if (bInData == 0xA3)
                                        {
                                            if (Data_Start4 == false)
                                            {
                                                Data_Start4 = true;
                                                rcvList4.Clear();
                                            }
                                        }


                                        if (Data_Start4 == true)
                                        {
                                          

                                            pictureBox15_Visible(true);
                                            pictureBox16_Visible(false);

                                            TimeOut4 = 50;
                                            rcvList4.Add(bInData);
                                            label27.ForeColor = Color.Red;
                                            label27.Text = "Data Receive";


                                            //  txt = String.Format("{0:X2} ", bInData);
                                        }
                                        else
                                        {
                                            Data_Start4 = false;
                                            label27.ForeColor = Color.Black;
                                            label27.Text = "NORMAL";
                                        }


                                        if (rcvList4.Count > 15)
                                        {
                                            if (rcvList4[rcvList4.Count - 2] == 0xfd && rcvList4[0] == 0xA3)
                                            {

                                                try
                                                {
                                                    Data_Start4 = false;
                                                    //   label3.Text = String.Format("{0:X2} ", rcvList[rcvList.Count - 1]);
                                                    byte[] Data_ = new byte[1];





                                                    if (rcvList4.Count != 0)
                                                    {
                                                        Data_[0] = rcvList4[rcvList4.Count - 1];
                                                        rcvList4.RemoveAt(rcvList4.Count - 1);
                                                    }
                                                    else
                                                    {
                                                        label27.Text = "NORMAL";
                                                        status = 0;
                                                        rcvList4.Clear();
                                                    }

                                                    byte CheckSum = 0;

                                                    for (int q = 385; q < rcvList4.Count; q++)
                                                    {
                                                        CheckSum += rcvList4[q];
                                                    }



                                                    if (CheckSum == Data_[0])
                                                    {

                                                        //////////////////////////////// 한개 짜리 tmf 저장 

                                                        string NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}",
                                                                                                   (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

                                                           string carnum = "";
                                                           if (rcvList4.Count > 384)
                                                           {
                                                               carnum = String.Format("{0:X2}-{1:X2}{2:X2}", rcvList4[336], rcvList4[335], rcvList4[334]);
                                                           }

                                                        //  string TmpFile = Application.StartupPath + "\\" + NowReceiveTime + ".TMF";
                                                           string newPath = System.IO.Path.Combine(TACHO2_path + "\\TMF_3", "AMF");
                                                           // Create the subfolder
                                                           System.IO.Directory.CreateDirectory(newPath);

                                                        string TmpFile = TACHO2_path + "\\TMF_3\\AMF\\" + NowReceiveTime + "_" + carnum + ".TMF";


                                                        byte[] rcvByte = new byte[rcvList4.Count];
                                                        rcvList4.CopyTo(rcvByte);

                                                        FileStream fs = new FileStream(TmpFile, FileMode.OpenOrCreate, FileAccess.Write);
                                                        BinaryWriter bw = new BinaryWriter(fs);

                                                        bw.Write(rcvByte);

                                                        serialPort4.DiscardInBuffer();
                                                        serialPort4.DiscardOutBuffer();
                                                        serialPort4.Write(Data_, 0, Data_.Length);

                                                        fs.Close();
                                                        bw.Close();
                                                        Tacho_Run(TmpFile);
                                                        svr.TachoSend4 = true;

                                                        //////////////////////////////////////   2번째 저장 total tmf  기존 데이터와  이어 붙히자 !!




                                                        string TMFPath = System.IO.Path.Combine(TACHO2_path + "\\TMF_3", "TransData");
                                                        // Create the subfolder
                                                        System.IO.Directory.CreateDirectory(TMFPath);

                                                        NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                                                                (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day);
                                                        TmpFile = TACHO2_path + "\\TMF_3\\TransData\\" + NowReceiveTime + ".TMF";

                                                        rcvList4.RemoveAt(0);
                                                        rcvByte = new byte[rcvList4.Count];
                                                        rcvList4.CopyTo(rcvByte);

                                                        fs = new FileStream(TmpFile, FileMode.Append, FileAccess.Write);
                                                        bw = new BinaryWriter(fs);

                                                        bw.Write(rcvByte);
                                                        fs.Close();
                                                        bw.Close();


                                                    }
                                                    else
                                                    {
                                                        Data_Start4 = false;
                                                        label27.Text = "NORMAL";
                                                    }




                                                    status = 0;
                                                    rcvList3.Clear();
                                                    Data_Start3 = false;



                                                    status = 0;



                                                }
                                                catch (Exception ex)
                                                {
                                                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                                    string path = Application.StartupPath + "\\ErrorLog.jie";
                                                    using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                                                    {
                                                        sw.WriteLine("Run_SerialThread4_TMFmake :[" + DateTime.Now.ToString() + "] " + ex.StackTrace);
                                                    }
                                                }



                                            }
                                        }

                                        //     richTextBoxMsg.AppendText(txt);



                                    }
                                }
                            }
                          //  Thread.Sleep(50);

                            break;
                        case 2:

                            break;
                    } //switch(status)
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    string path = Application.StartupPath + "\\ErrorLog.jie";
                    using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                    {
                        sw.WriteLine("Run_SerialThread4 :[" + DateTime.Now.ToString() + "] " + ex.StackTrace);
                    }
                }
            }

        }
        private void buttonClose_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort1.IsOpen)
                    serialPort1.Close();
                label13.Text = comboBoxPort.SelectedItem.ToString() + " Close";

                label4.ForeColor = Color.Black;
                label4.Text = "NORMAL";

              

                pictureBox9_Visible(true);
                pictureBox10_Visible(false);

                pictureBox2_Visible(true);
                pictureBox1_Visible(false);

            }
            catch
            {
                MessageBox.Show("Can't close already opened port", "Error");
                return;
            }
        }


        public void Add1(string dir)
        {


            //tn = treeView1.Nodes.Add(dir);

        }
        TreeNode TACHO;
       
        TreeNode Month;
        TreeNode Year;

        public void Add2(TreeNode child)
        {
            string FileName = "";
            string Year_str = "";
            string Year_ = "20";
            string Month_str = "";
            string Day_str = "";

            if (start_tree == true)
            {
              //  dirname = "KEY";
                TACHO = new TreeNode("DATE");
                Year = new TreeNode(Year_str);
                Month = new TreeNode(Month_str);
                treeView1.Nodes.Add(TACHO);
                //	a.Nodes.Add(Year);


            }
            else if (Second_tree == true)
            {

               // dirname += "년";
                Year = new TreeNode(dirname);
                TACHO.Nodes.Add(Year);

                Second_tree = false;
               
            }
            else if (third_tree ==true)
            {
               
                  Month = new TreeNode(dirname);
                  Year.Nodes.Add(Month);
                  third_tree = false;
            }
            else
            {
                FileName = child.Text;

                for (int i = 0; i < FileName.Length; i++)
                {

                    if (i < 2)
                    {
                        Year_ += FileName[i];
                    }
                    else if (i > 1 && i < 4)
                    {
                        Month_str += FileName[i];
                    }
                    else
                    {
                        Day_str += FileName[i];
                    }

                }

              //  Year_ += "년";

                Month_str += "월";
                Day_str += "일";
                Month.Nodes.Add(child);
                 ///////////////////////////////////////////////  년도  노드 를 찾은후 -> 월 노드를 추가 할것인지 판단 해야함 
                /*
                for (int i = 0; i < TACHO.Nodes.Count; i++)
                {
                    bool nodeCheck = false;  // 월노드가 0이 아니지만  월노드를 새로 추가 해야 할경우를 위해....
                    if (TACHO.Nodes[i].Text == Year_)
                    {
                        if (TACHO.Nodes[i].Nodes.Count == 0)   //  월 노드가 0 
                        {
                            Month = new TreeNode(Month_str);
                            TACHO.Nodes[i].Nodes.Add(Month);
                            Month.Nodes.Add(child);
                            //  TACHO.Nodes[i].Nodes[i].Nodes.Add(child);
                            break;
                        }
                        else
                        {
                            for (int j = 0; j < TACHO.Nodes[i].Nodes.Count; j++)  // 월 노드 찾기  
                            {
                                if (TACHO.Nodes[i].Nodes[j].Text == Month_str)  // 월노드가 있음
                                {
                                    nodeCheck = true;
                                    Month.Nodes.Add(child);
                                }
                            }


                            if (nodeCheck == false)
                            {
                              
                                Month = new TreeNode(Month_str);
                                TACHO.Nodes[i].Nodes.Add(Month);
                                Month.Nodes.Add(child);
                            }

                        }


                       
                    }
                }*/
              
                /*

                    if (Year.Text != Year_)
                    {


                       // Year = new TreeNode(Year_);
                      //  b.Nodes.Add(Year);
                        Month = new TreeNode(Month_str);
                        Year.Nodes.Add(Month);

                        Month.Nodes.Add(child);

                    }
                    else
                    {
                        if (Month.Text == Month_str)
                        {
                            Month.Nodes.Add(child);

                        }
                        else
                        {
                            Month = new TreeNode(Month_str);
                            Year.Nodes.Add(Month);
                            Month.Nodes.Add(child);

                        }
                    }
                */


            }


        }
        public void GetDirectoryNodes(TreeNode root, DirectoryInfo dirs, bool isLoop)
        {

            try
            {
                DirectoryInfo[] DIRS = dirs.GetDirectories();
                FileInfo[] files = dirs.GetFiles();

                /*
                DirectoryInfo[] DIRS = dirs.GetDirectories();
                FileInfo[] files = dirs.GetFiles();

                string[] file_str = new string[files.Length];

                char[] trimChars = { '.', 'm', 'd', 'b' };
                int cnt = 0;
                for (int i = 0; i < files.Length; i++)
                {

                    if (files[i].Extension != ".ldb")
                    {

                        file_str[i] = files[i].ToString();
                        file_str[i] = file_str[i].TrimEnd(trimChars);

                    }


                }

                */


                    foreach (DirectoryInfo dir in DIRS)
                    {
                        TreeNode child = new TreeNode(dir.Name);

                        if (dir.Name == "TACHO")
                        {

                            dirname = dir.Name;

                            start_tree = true;
                            treeView1.Invoke(new Add_a(Add2), new object[] { child });

                            string path = "";
                            if (ViewerMode == true)
                            {
                                path = @"\\" + ShareIP + "\\tacho2\\TACHO\\" ;
                            }
                            else
                            {

                                 path = TACHO2_path + "\\TACHO\\";
                            }


                            DirectoryInfo dir1 = new DirectoryInfo(path);
                            DirectoryInfo[] tempDIRS = dir1.GetDirectories();

                           for (int i = tempDIRS.Length - 1; i > 0; i--)  //  Year폴더 정렬 하자 높은월이 맨위로 
                            {

                                for (int j = 0; j < i; j++)
                                {
                                    int a = Int32.Parse(tempDIRS[j].Name);
                                    int b = Int32.Parse(tempDIRS[j + 1].Name);

                               
                                    if (a < b)
                                    {

                                        DirectoryInfo temp = tempDIRS[j];

                                        tempDIRS[j] = tempDIRS[j + 1];

                                        tempDIRS[j + 1] = temp;

                                    }

                                }

                            }


                          
                            foreach (DirectoryInfo d in tempDIRS)
                            {

                             

                              //  path = TACHO2_path + "\\TACHO\\" + d.Name;

                                if (ViewerMode == true)
                                {
                                    path =@"\\" + ShareIP + "\\tacho2\\TACHO\\" + d.Name;
                                }
                                else
                                {

                                    path = TACHO2_path + "\\TACHO\\" + d.Name;
                                }

                                  dirname = d.Name;
                           
                                start_tree = false;
                                Second_tree = true;
                                treeView1.Invoke(new Add_a(Add2), new object[] { child });
                         
                          
                                DirectoryInfo dir_ = new DirectoryInfo(path);

                              //   DirectoryInfo[] MonthFolder_tmp = dir_.GetDirectories();

                                 DirectoryInfo[] MonthFolder = dir_.GetDirectories();
                                
                                 for (int i = MonthFolder.Length-1; i > 0; i--)  // Month 폴더 정렬 하자 높은월이 맨위로 
                                 {

                                     for (int j = 0; j < i; j++)
                                     {
                                         int a = Int32.Parse(MonthFolder[j].Name);
                                         int b = Int32.Parse(MonthFolder[j+1].Name);
                                         if (a < b)
                                         {

                                             DirectoryInfo temp = MonthFolder[j];

                                             MonthFolder[j] = MonthFolder[j + 1];

                                             MonthFolder[j + 1] = temp;

                                         }

                                     }

                                 }

                              
                                foreach (DirectoryInfo m_folder in MonthFolder)
                                {
                                    dirname = m_folder.Name;
                                    start_tree = false;
                                    Second_tree = false;
                                    third_tree = true;
                                    treeView1.Invoke(new Add_a(Add2), new object[] { child });


                                    files = m_folder.GetFiles();

                                    string[] file_str = new string[files.Length];

                                    char[] trimChars = { '.', 'm', 'd', 'b' };
                                    int cnt = 0;
                                    for (int i = 0; i < files.Length; i++)
                                    {

                                        if (files[i].Extension != ".ldb")
                                        {

                                            file_str[i] = files[i].ToString();
                                            file_str[i] = file_str[i].TrimEnd(trimChars);

                                        }


                                    }
                                    for (int i = file_str.Length - 1; i > 0; i--)  //Day 파일 정렬 하자 높은일이 맨위로 
                                    {

                                        for (int j = 0; j < i; j++)
                                        {
                                           // if (file_str[j] != null && file_str[j + 1] != null)
                                          //  {
                                            int a = 0;
                                            int b = 0;
                                            if (file_str[j] != null)
                                            {
                                                a = Int32.Parse(file_str[j]);
                                            }
                                        
                                            

                                            if (file_str[j + 1] != null)
                                            {
                                                 b = Int32.Parse(file_str[j + 1]);
                                            }

                                               
                                               
                                             //   if (files[i].Extension != ".ldb")
                                             //   {
                                                    if (a < b)
                                                    {

                                                        string temp = file_str[j];

                                                        file_str[j] = file_str[j + 1];

                                                        file_str[j + 1] = temp;

                                                    }
                                             //   }
                                          //  }

                                        }

                                    }
                                    foreach (string dir3 in file_str)
                                    {
                                        if (dir3 != null)
                                        {
                                            TreeNode child1 = new TreeNode(dir3);
                                            dirname = dir3;
                                            start_tree = false;
                                            Second_tree = false;
                                            third_tree = false;
                                            treeView1.Invoke(new Add_a(Add2), new object[] { child1 });

                                        }
                                    }
                                }
                              
                                

                            }
                          


                        }                       
                      

                         /*    path = TACHO2_path + "\\TACHO\\2015\\";
                            dir1 = new DirectoryInfo(path);
                            start_tree = false;
                            GetDirectoryNodes(child, dir1, false);*/
                      
                    }



                    



                
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message);
            }
           

        }
      
       public void GetDataNodes(TreeNode root, DirectoryInfo dirs, bool isLoop)
       {
           




       }
  

        public void Treeview_Refresh()
        {

            //	treeView1.Nodes.Clear();
            //tn = new TreeNode(name);
            //treeView1.Nodes.Add(name);


            treeView1.Nodes.Clear();
            //	string path = Application.StartupPath;
            string path = "";

            if (ViewerMode == true)
            {
                path = @"\\" + ShareIP + "\\tacho2\\";
            }
            else
            {
                path = TACHO2_path;
            }
            string filename = "Information.mdb";
            string filename1 = "Information1.mdb";
            string filename2 = "Information2.mdb";
            string filename3 = "Information3.mdb";
            if (Directory.Exists(TACHO2_path))
            {
            
                DirectoryInfo dir = new DirectoryInfo(path);


                DirectorySecurity dSecurity = dir.GetAccessControl();

                dSecurity.AddAccessRule(new FileSystemAccessRule("Users",
                                      FileSystemRights.Modify,      // ==> 수정 권한 부여
                                      InheritanceFlags.ObjectInherit,
                                      PropagationFlags.InheritOnly,
                                    AccessControlType.Allow));




                dir = new DirectoryInfo(Application.StartupPath);

                dSecurity = dir.GetAccessControl();

                dSecurity.AddAccessRule(new FileSystemAccessRule("Users",
                                      FileSystemRights.Modify,      // ==> 수정 권한 부여
                                      InheritanceFlags.ObjectInherit,
                                      PropagationFlags.InheritOnly,
                                    AccessControlType.Allow));


                TreeNode tt = new TreeNode(name);
                dir = new DirectoryInfo(path);
                GetDirectoryNodes(tt, dir, true);
                //	 treeView1.ExpandAll();

            }
            else
            {
                System.IO.Directory.CreateDirectory(path);

               
                string Filesource = System.IO.Path.Combine(Application.StartupPath, filename);
                System.IO.File.Copy(Filesource, path + "\\" + filename);   //

                DirectoryInfo dir = new DirectoryInfo(path);


                DirectorySecurity dSecurity = dir.GetAccessControl();

                dSecurity.AddAccessRule(new FileSystemAccessRule("USERS",
                                      FileSystemRights.Modify,      // ==> 수정 권한 부여
                                      InheritanceFlags.ObjectInherit,
                                      PropagationFlags.InheritOnly,
                                    AccessControlType.Allow));


                dir = new DirectoryInfo(Application.StartupPath);

                dSecurity = dir.GetAccessControl();

                dSecurity.AddAccessRule(new FileSystemAccessRule("USERS",
                                      FileSystemRights.Modify,      // ==> 수정 권한 부여
                                      InheritanceFlags.ObjectInherit,
                                      PropagationFlags.InheritOnly,
                                    AccessControlType.Allow));

            }



            if (System.IO.File.Exists(path + "\\" + filename))  // information 같은 파일의이름이 존재 함
            {

            }
            else
            {
                string Filesource = System.IO.Path.Combine(Application.StartupPath, filename);
                System.IO.File.Copy(Filesource, path + "\\" + filename);   //
            }
            ////////////////////////////////////////////////////////////////////////////////////////////////////

     


            ////////////////////////////////////////////////////////////////////////////////////////////////////
            //	Thread treeThread = new Thread(FillTree);
            //	treeThread.Start();
            //	start_tree = false;

           /*
                for (int i = 0; i < treeView1.Nodes.Count; i++)
                {
                    if (treeView1.Nodes[i].Text == "TACHO")
                    {

                        treeView1.Nodes[i].ExpandAll();

                        // this.listView1.Items[formTachoReceive1.listView1.Items.Count - 1].EnsureVisible(); 
                        treeView1.Nodes[0].EnsureVisible();

                    }
                }

           */
            try
            {
                if (treeView1.Nodes.Count != 0)
                {


                    DateTime DbTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day); //현재 시간 
                    for (int i = 0; i < treeView1.Nodes.Count; i++)
                    {
                        if (treeView1.Nodes[i].Text == "DATE" || treeView1.Nodes[i].Text == "KEY")
                        {

                            treeView1.Nodes[i].Expand();

                            for (int j = 0; j < treeView1.Nodes[i].Nodes.Count; j++)
                            {
                                if (treeView1.Nodes[i].Nodes[j].Text == (DbTime.Year.ToString()))
                                {
                                    treeView1.Nodes[i].Nodes[j].ExpandAll();
                                }
                            }
                            //  treeView1.Nodes[i].Nodes[treeView1.Nodes[i].Nodes.Count-1].ExpandAll();

                        }
                    }

                    treeView1.Nodes[0].EnsureVisible();

                    if (ViewerMode == true)
                    {
                        path = @"\\" + ShareIP + "\\tacho2\\" + "TACHO" + "\\" + treeView1.Nodes[0].Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Nodes[0].Text;
                    }
                    else
                    {
                        path = TACHO2_path + "TACHO" + "\\" + treeView1.Nodes[0].Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Nodes[0].Text;
                    }
                    //   path = TACHO2_path + e.Node.Parent.Parent.Parent.Text + "\\" + e.Node.Parent.Parent.Text + "\\" + e.Node.Parent.Text;
                    string[] files = Directory.GetFiles(path, "*.mdb");


                    for (int i = 0; i < files.Length; i++)
                    {
                        FileInfo file = new FileInfo(files[i]);

                        files[i] = file.Name;
                        if (files[i] == treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes[0].Text + ".mdb")
                        {

                            mdbfilename = path + "\\" + files[i];

                            DB_ReadData(0, 1);


                        }

                    }
              
                         
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
           
          
           
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            RegistryKey reg;
            reg = Registry.LocalMachine.CreateSubKey("SOFTWARE").CreateSubKey("TachoPlus");


          /*  reg.DeleteSubKey("Text", false);
            reg.DeleteSubKey("Num", false);
            reg.DeleteSubKey("Check", false);
            Registry.LocalMachine.DeleteSubKey("SOFTWARE\\TachoPlus");*/

            m_strText = Convert.ToString(reg.GetValue("Text", ""));
            m_nNum = Convert.ToInt32(reg.GetValue("Num", 0));
            m_bCheck = Convert.ToBoolean(reg.GetValue("Check", false));

            CheckForIllegalCrossThreadCalls = false;
            string[] ports = System.IO.Ports.SerialPort.GetPortNames();
            label31.Visible = false;
            this.richTextBox1.Visible = false;
            button5.Visible = false;
            //////////////////////////////////////////////////////////////////
            pictureBox17.Visible = false;
            string drive = "";
            if (drive == "" || drive == null)
            {

                drive = "C";

            }
            ManagementObject disk = new ManagementObject("win32_logicaldisk.deviceid=\"" + drive + ":\"");
            disk.Get();
            //   MessageBox.Show(disk["VolumeSerialNumber"].ToString());

            Hdd_Serial = disk["VolumeSerialNumber"].ToString();

//


      

            groupBox1.Visible = false;
            groupBox4.Visible = false;
            groupBox6.Visible = false;
          
            button10.Visible = false;

          
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            groupBox5.Visible = false;
            groupBox7.Visible = false;

            Client1_Icon.Visible = false;
            Client2_Icon.Visible = false;
            Client3_Icon.Visible = false;
            Client4_Icon.Visible = false;

            textBox1.Visible = false;

         //   Server_icon.Visible = false;
            button10.Visible = false;
            button14.Visible = false;
            button12.Visible = false;
        //    button18.Visible = false;
            button17.Visible = false;
          //  button13.Visible = false;

            button8.Visible = false;  // total eeprom
        
            label32.Text = "";

            que.Enqueue(1);
            que.Enqueue(2);
            que.Enqueue(3);
            que.Enqueue(4);

            ImageList dummyImageList = new ImageList();
            dummyImageList.ImageSize = new System.Drawing.Size(1, 18);
            listView1.SmallImageList = dummyImageList;
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;

            string path = Application.StartupPath + "\\TachoPlus.ini";
            TACHO2_path = inicls.GetIniValue("Tacho Init", "path", path); // 타코 루트
            string filename = "Information.mdb";

            System.IO.Directory.CreateDirectory(TACHO2_path);

            string Japan = "";
            Japan = inicls.GetIniValue("Tacho Init", "JAPAN", path);  //  

            if (Japan == "1")
            {
                label17.Visible = false;
                label31.Visible = false;
                linkLabel1.Visible = false;
            }
            else
            {

            }


            if (ports.Length > 0)
            {

                foreach (string port in ports)//PC 에 있는 시리얼 포트 찾아서 저장
                {
                    comboBoxPort.Items.Add(port);
                    comboBox2.Items.Add(port);
                    comboBox4.Items.Add(port);
                 
                }
                comboBoxBaud.Items.Add("9600");
                comboBoxBaud.Items.Add("19200");
                comboBoxBaud.Items.Add("38400");
                comboBoxBaud.Items.Add("57600");
                comboBoxBaud.Items.Add("115200");


                if (Japan == "1")
                {

                    comboBox1.Items.Add("9600");
                    comboBox3.Items.Add("9600");
                }
                else
                {
                    comboBox1.Items.Add("38400");
                    comboBox3.Items.Add("38400");
                }



           
            



                if (ports.Length == 4)
                {

                    comboBoxPort.SelectedIndex = 0;
                    comboBox2.SelectedIndex = 1;
                    comboBox4.SelectedIndex = 2;
                 
               
                }

                comboBoxPort.SelectedIndex = 0;

                comboBoxBaud.Enabled = false;
                comboBox1.Enabled = false;
                comboBox3.Enabled = false;


                if (Japan == "1")
                {
                    serialPort1.BaudRate = 9600;
                    serialPort1.Encoding = Encoding.Default;
                    serialPort1.DataBits = 8;
                    serialPort1.StopBits = StopBits.One;


                    serialPort2.BaudRate = 9600;
                    serialPort2.Encoding = Encoding.Default;
                    serialPort2.DataBits = 8;
                    serialPort2.StopBits = StopBits.One;


                    serialPort3.BaudRate = 9600;
                    serialPort3.Encoding = Encoding.Default;
                    serialPort3.DataBits = 8;
                    serialPort3.StopBits = StopBits.One;


                    serialPort4.BaudRate = 9600;
                    serialPort4.Encoding = Encoding.Default;
                    serialPort4.DataBits = 8;
                    serialPort4.StopBits = StopBits.One;
                }
                else
                {
                    serialPort1.BaudRate = 9600;
                    serialPort1.Encoding = Encoding.Default;
                    serialPort1.DataBits = 8;
                    serialPort1.StopBits = StopBits.One;


                    serialPort2.BaudRate = 9600;
                    serialPort2.Encoding = Encoding.Default;
                    serialPort2.DataBits = 8;
                    serialPort2.StopBits = StopBits.One;


                    serialPort3.BaudRate = 9600;
                    serialPort3.Encoding = Encoding.Default;
                    serialPort3.DataBits = 8;
                    serialPort3.StopBits = StopBits.One;


                    serialPort4.BaudRate = 9600;
                    serialPort4.Encoding = Encoding.Default;
                    serialPort4.DataBits = 8;
                    serialPort4.StopBits = StopBits.One;
                }

                if (Japan == "1")
                {
                    comboBoxBaud.SelectedIndex = comboBoxBaud.Items.IndexOf("9600");
                    //   comboBoxBaud.SelectedIndex = comboBoxBaud.Items.IndexOf("9600");
                    //  comboBoxBaud.SelectedIndex = comboBoxBaud.Items.IndexOf("115200");
                    comboBox1.SelectedIndex = comboBox1.Items.IndexOf("9600");
                    comboBox3.SelectedIndex = comboBox3.Items.IndexOf("9600");
                }
                else
                {
                    comboBoxBaud.SelectedIndex = comboBoxBaud.Items.IndexOf("38400");
                    //   comboBoxBaud.SelectedIndex = comboBoxBaud.Items.IndexOf("9600");
                    //  comboBoxBaud.SelectedIndex = comboBoxBaud.Items.IndexOf("115200");
                    comboBox1.SelectedIndex = comboBox1.Items.IndexOf("38400");
                    comboBox3.SelectedIndex = comboBox3.Items.IndexOf("38400");
                }
         

            }
            

         


            if (System.IO.File.Exists(TACHO2_path + "\\" + filename))  // information 같은 파일의이름이 존재 함
            {

            }
            else
            {
                string Filesource = System.IO.Path.Combine(Application.StartupPath, filename);
                System.IO.File.Copy(Filesource, TACHO2_path + "\\" + filename);   //
            }



             Encrypt = EncryptString(Hdd_Serial, "JIE");
            // string code = inicls.GetIniValue("Tacho Init", "Registration", path);
             Dencrypt = DecryptString(Encrypt, "JIE");

             string code = m_strText;
            /*
           //  if (code == null || code =="")
             if (m_strText == "" || m_nNum == 0 || !m_bCheck)           
             {
                 RegistrationForm registsration = new RegistrationForm();

                 DialogResult dialog1 = registsration.ShowDialog();
                 if (dialog1== DialogResult.OK)
                 {


                     registsration.Close();
                     code =  Convert.ToString(reg.GetValue("Text", ""));;
                 
                 }
                 else
                 {
                     this.Close();
                     return;
                 }
             }
          
            if (code != Encrypt)
            {
                RegistrationForm registsration = new RegistrationForm();

                DialogResult dialog_1 = registsration.ShowDialog();

                if (dialog_1 == DialogResult.OK)
                {
                  

                        registsration.Close();
                        code =  Convert.ToString(reg.GetValue("Text", ""));;
                        if (code != Encrypt)
                        {
                             //  MessageBox.Show("인증실패");
                              this.Close();
                          
                              return;
                        }

                  

                }
                else
                {
                      this.Close();
                      this.Dispose();
                      return;
                }
               
             //   MessageBox.Show("인증실패");
              //  this.Close();
              //  return;
            }
            */


          
            string ServerStr = "";
            ServerStr = inicls.GetIniValue("Tacho Init", "ServerOpen", path);  //  Server Open

            if (ServerStr == "1")
            {
                ServerOpen = true;
             
                
            }
            else
            {

              //  Server_icon.Visible = false;
                button10.Visible = false;
                     label32.Text = "";
                ServerOpen = false;
                Client_label.Visible = false;
            }

            if (ServerOpen == true)
            {
                if (ServerSocket != null)
                {
                    //strErr = string.Format("[서버] : 현재 서버 가동중입니다.!!");
                    //   Add_Log(strErr);
                    return;
                }
                ServerThread = new Thread(new ThreadStart(Thread_Run));
                ServerThread.IsBackground = true;
                ServerThread.Start();
            }
            
            string TransTimeEnable = "";
            TransTimeEnable = inicls.GetIniValue("Tacho Init", "TrnasTimeEnble", path);  //  Server Open

            if (TransTimeEnable == "1")
            {
                TIME1_ENABLE = true;
            }
            else
            {
                TIME1_ENABLE = false;
            }
           
            Time1_str = inicls.GetIniValue("Tacho Init", "TrnasTime", path);

            string str = inicls.GetIniValue("Tacho Init", "IbuttonEnable", path);

            if (str == "1")
            {
                button9.Visible = false;
                groupBox10.Visible = false;
               // label17.Visible = false;
                //    button15.Visible = true;
                iButtonMode = true;
                label31.Visible = true;
               // Ibuttonthread = new Thread(new ThreadStart(Run_Ibutton));
               // Ibuttonthread.IsBackground = true;
               // Thread.Sleep(100);
               // Ibuttonthread.Start();
             //   linkLabel1.Text = "www.taximeter.net";
                ServerIP = inicls.GetIniValue("Tacho Init", "Manager IP", path);


            }
            else
            {
                groupBox9.Visible = false;
           ///     linkLabel1.Text = "www.taximeter.net";
                label17.Visible = false;
                label31.Visible = true;
                button16.Visible = false;
                iButtonMode = false;
                button9.Visible = true;
                button6.Visible = false;
            }
        
            str = inicls.GetIniValue("Tacho Init", "IRD_Enable", path);

            if (str == "1")
            {
                IRD_Enable = true;
                label32.Text = "IRD Mode";
                serialPort1.RtsEnable = true;
                comboBoxBaud.SelectedIndex = comboBoxBaud.Items.IndexOf("9600");
            }
            else
            {
                IRD_Enable = false;
            }

            str = inicls.GetIniValue("Tacho Init", "ViewerMode", path);

            if (str == "1")
            {
                ViewerMode = true;
            }
            else
            {
                ViewerMode = false;
            }



            str = inicls.GetIniValue("Tacho Init", "CashierMode", path);

            if (str == "1")
            {
                CashierMode = true;
                ViewerMode = false;

                ShareIP = inicls.GetIniValue("Tacho Init", "Share IP", path);
            }
            else
            {
                CashierMode = false;
            }

          


            if (ViewerMode == false)
            {
                button17.Enabled = true;
               // button6.Enabled = false;
            
                if (button9.Visible == false)
                {
                   // LoginForm registsration = new LoginForm(this);

                  //  DialogResult dialog = registsration.ShowDialog();
                }
               
            }
            else
            {
               // button9.Visible = false;
                button6.Visible = false;
                ShareIP = inicls.GetIniValue("Tacho Init", "Share IP", path);
                registrationRToolStripMenuItem1.Enabled = false;
              //  button17.Enabled = false;
              //  button16.Enabled = false;
                button6.Enabled = false;
            }
            string temp = inicls.GetIniValue("Tacho Init", "MALAYSIA", path);
             if (temp == "1")
             {
                 MALAYSIA_Set = true;
                 label17.Visible = false;
                 button6.Visible = false;
             }

              temp = inicls.GetIniValue("Tacho Init", "UAE", path);
              if (temp == "1")
              {
                  MALAYSIA_Set = false;
                  UAE_Set = true;
                  label17.Visible = true;
                  button6.Visible = false;
                  button18.Visible = true;
                
              }
              else
              {
                  label17.Visible = false;
                  button18.Visible = false;
              }

              temp = inicls.GetIniValue("Tacho Init", "THAILAND", path);
            if (temp == "1")
            {
                THAILAND_Set = true;
            }



              temp = inicls.GetIniValue("Tacho Init", "Monthly Start Date", path);

              StartDay = Int32.Parse(temp);  

          
              temp = inicls.GetIniValue("Tacho Init", "Print_Enable", path);
             if (temp == "1")
             {
                 MiniPrintEnable = true;
                 MiniPrintPort = inicls.GetIniValue("Tacho Init", "Print_Port", path);

                 ///////////////////////////////////////////////////////////


                 long val = ConnectSerialPrinter(MiniPrintPort, 9600, 1000, false);
                 try
                 {
                     switch (val)
                     {
                         case SUCCESS:
                             MiniPrintCon = true;
                             break;
                         case ALREADY_OPENED:
                             this.label32.Text = "Already opened";
                             break;
                         case UNABLE_TO_OPEN_THE_PORT:
                             this.label32.Text = "Unable to open the port";
                             MiniPrintCon = false;
                             break;
                         case UNABLE_TO_CONFIGURE_THE_SERIAL_PORT:
                             this.label32.Text = "Unable to configure the serial port";
                             MiniPrintCon = false;
                             break;
                         case UNABLE_TO_SET_THE_TIMEOUT_PARAMETERS:
                             this.label32.Text = "Unable to set the timeout parameters";
                             MiniPrintCon = false;
                             break;
                         case NO_RESPONSE_FROM_PRINTER:
                             this.label32.Text = "No response from printer";
                             MiniPrintCon = false;
                             break;
                         case TIMEOUT:
                             this.label32.Text = "TIMEOUT";
                             MiniPrintCon = false;
                             break;
                     }
                 }
                 catch
                 {
                     MessageBox.Show(e.ToString());
                 }
                 ////////////////////////////////////////////////////////////////

             }
             else
             {
                 MiniPrintEnable = false;
             }



             Print_Text = inicls.GetIniValue("Tacho Init", "Print_Text", path);


             if (iButtonMode == false)
             {
                 Treeview_Refresh();
                 label1.Visible = false;
                 label5.Visible = false;
                 label8.Visible = false;
                 label9.Visible = false;
                 label18.Visible = false;
                 Taxiid_textBox.Visible = false;
                 Driverid_textBox.Visible = false;
                 Drivername_textBox.Visible = false;
                 textBox2.Visible = false;
                 textBox3.Visible = false;
                 groupBox8.Visible = false;
                 button17.Visible = false;
                 this.Text = "TachoPlusV0.14";
              //   this.listView1.Columns[3].Width = 0;
                 this.listView1.Columns[8].Width = 0;
                 this.listView1.Columns[9].Width = 0;
                 this.listView1.Columns[14].Width = 0;
                 this.listView1.Columns[15].Width = 0;
                 registrationRToolStripMenuItem1.Visible = false;
                 this.Text = "TachoPlusV0.14";
                 linkLabel1.Text = "www.taximeter.net";

             }
             else
             {
                // button17.Visible = false;

                 if (UAE_Set == true)
                 {

                     this.Text = "CashierSystemV0.14";
                     linkLabel1.Text = "www.luxumborj.com";
                 }
                 else
                 {
                     this.Text = "TachoPlusV0.14";
                     linkLabel1.Text = "www.taximeter.net";
                 }
                  Login = new LoginForm(this);

                 DialogResult dialog = Login.ShowDialog();
                 try
                 {

                    // Treeview_Refresh();
                 }
                 catch (Exception ex)
                 {
                     MessageBox.Show(ex.Message);
                 }

                 if (LoginOK == false)
                 {
                     Application.Exit();
                 }
                 Treeview_Refresh();
             }


             
         
        
          //  treeView1.TreeViewNodeSorter = new NodeSorter();

        }


     //   EY_ChatChild[] eyChild = new EY_ChatChild[4];
        private void SetDirectorySecurity(string linePath)
        {
            DirectorySecurity dSecurity = Directory.GetAccessControl(linePath);
            dSecurity.AddAccessRule(new FileSystemAccessRule("Users",
                                                                        FileSystemRights.FullControl,
                                                                        InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit,
                                                                        PropagationFlags.None,
                                                                        AccessControlType.Allow));
            Directory.SetAccessControl(linePath, dSecurity);
        }
        public byte[] OnWire_Read(byte[] Buffer, int Length, byte[] Address, TMEX TMEXLibrary, int hSess, short[] ROM)
        {
            int n = 0;

            TMEXLibrary.TMRom(hSess, stateBuffer, ROM);  // Selecet
            n = TMEXLibrary.TMStrongAccess(hSess, stateBuffer);



            TMEXLibrary.TMAccess(hSess, stateBuffer);
            TMEXLibrary.TMTouchByte(hSess, 0xF0);
            TMEXLibrary.TMTouchByte(hSess, (short)Address[0]);
            TMEXLibrary.TMTouchByte(hSess, (short)Address[1]);
            for (int i = 0; i < Buffer.Length; i++)
            {
                Buffer[i] = (byte)TMEXLibrary.TMTouchByte(hSess, 0xFF);
            }

            return Buffer;
        }

        public void OneWire_Write(byte[] Buffer, int Length, byte[] Address, TMEX TMEXLibrary, int hSess)
        {

            byte[] auth = new byte[3];

            TMEXLibrary.TMAccess(hSess, stateBuffer);
            TMEXLibrary.TMTouchByte(hSess, 0x0F);
            TMEXLibrary.TMTouchByte(hSess, (short)Address[0]);
            TMEXLibrary.TMTouchByte(hSess, (short)Address[1]);
            TMEXLibrary.TMBlockStream(hSess, Buffer, (short)Buffer.Length);

            // get target address and ending offset/data status byte
            TMEXLibrary.TMAccess(hSess, stateBuffer);
            TMEXLibrary.TMTouchByte(hSess, 0xAA);

            for (int i = 0; i < 3; i++)
            {
                auth[i] = (byte)TMEXLibrary.TMTouchByte(hSess, 0xFF);
            }

            // copy scratchpad to memory
            TMEXLibrary.TMAccess(hSess, stateBuffer);
            TMEXLibrary.TMTouchByte(hSess, 0x55);
            for (int i = 0; i < 3; i++)
            {
                TMEXLibrary.TMTouchByte(hSess, auth[i]);
            }
        }

        string TaxiID = "";
        string DriverID = "";
        string DriverName = "";
        int ID = 0;
        DateTime Outtime = new DateTime(1, 1, 1, 1, 1, 1);
        DateTime Intime = new DateTime(1, 1, 1, 1, 1, 1);

        public void PaidTime_Write()
        {



         
            short[] ROM = new short[8];
            //      byte[] block = new byte[322];
            byte[] TachoAddr = new byte[6]; // 타코 시작 주소와 끝 주소를 찾는다. 
            Int16 StartAddr = 0;
            Int16 EndAddr = 0;
            int n = 0;
            byte[] Address = new byte[2];
            byte[] TachoStartByte = new byte[1];
            byte[] TachoOverFlow = new byte[1];
            short ret = 0;
            short portNum = 0;
            short portType = 0;
            int hSess = 0;

            /////////////////////////////////////////////////////////////////
            TMEX TMEXLibrary = new TMEX();

            TMEXLibrary.TMReadDefaultPort(ref portNum, ref portType);

            hSess = TMEXLibrary.TMExtendedStartSession(portNum, portType, ref sessionOptions);
            bool check=true;

            do{
                if (hSess != 0)
                {
                    check = false;
                }
                else
                {
                    Thread.Sleep(1000);
                    TMEXLibrary = new TMEX();

                    TMEXLibrary.TMReadDefaultPort(ref portNum, ref portType);

                    hSess = TMEXLibrary.TMExtendedStartSession(portNum, portType, ref sessionOptions);
                }

            }while(check);


            if (hSess != 0)
            {

                // must be called before any non-session functions can be called
                TMEXLibrary.TMSetup(hSess);

                ret = TMEXLibrary.TMFirst(hSess, stateBuffer);

                if (ret != 1)
                {
                    MessageBox.Show("Disconnected  to ibutton.");
                    pictureBox17.Visible = false;
                    return;
                }

                n = 0;

                //   while ( (ret==1) && (n<32) )
                //   {
                // if MSB of ROM[0] != 0, then write, else read
                //   short[] ROM  = new short[8];

                TMEXLibrary.TMRom(hSess, stateBuffer, ROM);  // Selecet
                n = TMEXLibrary.TMStrongAccess(hSess, stateBuffer);

                if (ROM[0] != 0x0c)
                {
                    MessageBox.Show("Disconnected  to ibutton.");
                    pictureBox17.Visible = false;
                    ibutonCheck = true;
                    return;
                }
            }

            //  }

            ///////////////////////////////////////////////////////////////

            byte[] CarAddr = new byte[2];
            CarAddr[0] = 0xb4;
            CarAddr[1] = 0x00;


            byte[] DriverAddr = new byte[2];
            DriverAddr[0] = 0xc0;
            DriverAddr[1] = 0x00;


            byte[] TimeLimitAddr = new byte[2];
            TimeLimitAddr[0] = 0x20;
            TimeLimitAddr[1] = 0x00;


                    byte[] temp = new byte[2];



                    OneWire_Write(TimLimitByte, TimLimitByte.Length, TimeLimitAddr, TMEXLibrary, hSess);



                    byte[] Temp = new byte[TimLimitByte.Length];


                    Temp = OnWire_Read(Temp, Temp.Length, TimeLimitAddr, TMEXLibrary, hSess, ROM);

               

                    for (int i = 0; i < Temp.Length; i++)
                    {

                        if (Temp[i] != TimLimitByte[i])
                        {
                          /* bool Check=true;

                            do
                            {
                            OneWire_Write(TimLimitByte, TimLimitByte.Length, TimeLimitAddr, TMEXLibrary, hSess);


                            Temp = new byte[TimLimitByte.Length];


                            Temp = OnWire_Read(Temp, Temp.Length, TimeLimitAddr, TMEXLibrary, hSess, ROM);
                            if (Temp[i] == TimLimitByte[i])
                            {
                                Check = false;
                            }
                            else
                            {
                                Thread.Sleep(1000);
                                /////////////////////////////////////////////////////////////////
                                 TMEXLibrary = new TMEX();

                                TMEXLibrary.TMReadDefaultPort(ref portNum, ref portType);

                                hSess = TMEXLibrary.TMExtendedStartSession(portNum, portType, ref sessionOptions);
                                if (hSess != 0)
                                {

                                    // must be called before any non-session functions can be called
                                    TMEXLibrary.TMSetup(hSess);

                                    ret = TMEXLibrary.TMFirst(hSess, stateBuffer);

                                    if (ret != 1)
                                    {
                                        MessageBox.Show("Disconnected  to ibutton.");
                                        pictureBox17.Visible = false;
                                        return;
                                    }

                                    n = 0;

                                    //   while ( (ret==1) && (n<32) )
                                    //   {
                                    // if MSB of ROM[0] != 0, then write, else read
                                    //   short[] ROM  = new short[8];

                                    TMEXLibrary.TMRom(hSess, stateBuffer, ROM);  // Selecet
                                    n = TMEXLibrary.TMStrongAccess(hSess, stateBuffer);

                                    if (ROM[0] != 0x0c)
                                    {
                                        MessageBox.Show("Disconnected  to ibutton.");
                                        pictureBox17.Visible = false;
                                        ibutonCheck = true;
                                        return;
                                    }
                                }

                                //  }

                                ///////////////////////////////////////////////////////////////
                            }


                            } while (Check);*/

                            EY_ChatClient.Close();
                      
                            MessageBox.Show("Write Fail!");
                            pictureBox17.Visible = false;
                        
                            return;
                        }
                    }

                    byte[] TempStartTime = new byte[6];
                    byte[] TempStartTime_Addr = new byte[2];
                    TempStartTime_Addr[0] = 0x03;
                    TempStartTime_Addr[1] = 0x01;
                    TempStartTime = OnWire_Read(TempStartTime, TempStartTime.Length, TempStartTime_Addr, TMEXLibrary, hSess, ROM);

               
                   

                    EY_ChatClient.Close();
                    string DBstring = "";

                    try
                    {
                       
                        string NameDB = mdbfilename;
                        DBstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + NameDB;
                        //					 Db_backup = false;
                        OleDbConnection conn = new OleDbConnection(@DBstring);

                        conn.Open();

                        DateTime nowTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                        string queryUpdate = "UPDATE TblTacho SET " + "PaidTime" + "='" + nowTime
                            + "' WHERE ID=" + SelectItem.ToString();

                        OleDbCommand commUpdate = new OleDbCommand(queryUpdate, conn);

                        commUpdate.ExecuteNonQuery();


                         queryUpdate = "UPDATE TblTacho SET " + "RealIncome" + "='" + textBox3.Text
                            + "' WHERE ID=" + SelectItem.ToString();
                          commUpdate = new OleDbCommand(queryUpdate, conn);

                         commUpdate.ExecuteNonQuery();


                        string queryRead = "SELECT * FROM TblTacho ORDER BY ID";

              
              

                                    


                    OleDbCommand commRead = new OleDbCommand(queryRead, conn);
                    OleDbDataReader srRead = commRead.ExecuteReader();

                      
                        while (srRead.Read())
                        {
                            ID = srRead.GetInt32(0);

                            if (ID == SelectItem)
                            {

                                TaxiID = srRead.GetString(1);
                                DriverID = srRead.GetString(2);
                                DriverName = srRead.GetString(3);
                                Outtime = srRead.GetDateTime(4);
                                Intime = srRead.GetDateTime(5);
                            }
                        }


                        conn.Close();
                        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////// 
                        // 영수증 모드를 위하여 Amf_data 파싱을 해야 한다.!!!
                        if (MiniPrintCon == true)
                        {
                            GetTime_Receipt = true;
                            int year = BcdToDecimal(TempStartTime[0]);
                            year += 2000;
                            int month = BcdToDecimal(TempStartTime[1]);
                            //   StartDay 

                            StartDate = new DateTime(year, month, StartDay, 0, 0, 0);
                            EndDate = StartDate;

                            EndDate = EndDate.AddMonths(1).AddDays(-1);

                            // 영수증 모드를 위하여 Amf_data 파싱을 해야 한다.!!!

                            ///////////////////////////// 타코 길이 계산하기 

                            Address[0] = 0x40;
                            Address[1] = 0x00;


                            TachoAddr = OnWire_Read(TachoAddr, TachoAddr.Length, Address, TMEXLibrary, hSess, ROM);   // 타코 길이 알기 위해 주소를 타코 저장 위치 주소를 읽느나.

                            if (TachoAddr[0] == 0xff && TachoAddr[1] == 0xff && TachoAddr[2] == 0xff && TachoAddr[3] == 0xff && TachoAddr[4] == 0xff)
                            {
                                // IbuttonReadCheck = false;
                                //  errorstatus = true;
                                MessageBox.Show("Disconnected  to ibutton.");
                                return;
                            }

                            EndAddr = (Int16)(TachoAddr[1] << 8 & 0xff00);
                            EndAddr += TachoAddr[0];

                            StartAddr = (Int16)(TachoAddr[4] << 8 & 0xff00);
                            StartAddr += TachoAddr[3];

                            StartAddr = 0x100;  // 시작 주소는 고정으로 사용 !

                            int num = EndAddr - StartAddr;  // 타코 트랜젝션의 갯수를 구한다. 


                            if (EndAddr == 0x00)
                            {

                                //  IbuttonReadCheck = false;
                                //   errorstatus = true;
                                pictureBox17.Visible = false;
                                MessageBox.Show("Empty Tacho!");
                                return;
                            }



                            ///////////////////////////////////////////////////////////


                            byte[] Amf_array;  // AMf_file 총 싸이즈 data +headr +0xfd + checkSum

                            if (TachoOverFlow[0] == 0x01)  // Memory overflow 발생 
                            {
                                Amf_array = new byte[8194]; // 8192 +2(0xfd + checkSum)
                            }
                            else
                            {
                                Amf_array = new byte[num + 256 + 2]; // transaction length.
                            }


                            Address[0] = 0x00;
                            Address[1] = 0x00;



                            Amf_array = OnWire_Read(Amf_array, Amf_array.Length, Address, TMEXLibrary, hSess, ROM); // data Read

                            int checkcnt = 0;
                            for (int i = 0; i < 256; i++) //data 검사 
                            {

                                if (Amf_array[i] == 0xff)
                                {
                                    checkcnt++;
                                }
                            }

                            if (checkcnt == 256)
                            {
                                pictureBox17.Visible = false;
                                //   continue;
                            }

                            //////////////////////////////////////////////  data 검사 15.08.17
                            int index = 256;
                            int TachoLength = num / 64;

                            bool ErrCheck = false;
                            while (TachoLength != 0)
                            {
                                checkcnt = 0;
                                for (int i = 0; i < 64; i++)
                                {
                                    if (Amf_array[index] == 0xff)
                                    {
                                        checkcnt++;
                                    }
                                    index++;
                                    if (checkcnt > 5)
                                    {
                                        ErrCheck = true;
                                        // continue;
                                    }
                                }
                                TachoLength--;
                            }
                            if (ErrCheck == true)
                            {
                                pictureBox17.Visible = false;
                                //continue;
                                return;
                            }
                            ///////////////////////////////////////////////


                            string NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}",
                                                                        (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);


                            string carnum = "";
                            carnum = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}{8:C2}"
                                                       , Convert.ToChar(Amf_array[188]), Convert.ToChar(Amf_array[187]), Convert.ToChar(Amf_array[186]), Convert.ToChar(Amf_array[185])
                                                       , Convert.ToChar(Amf_array[184]), Convert.ToChar(Amf_array[183])
                                                    , Convert.ToChar(Amf_array[182]), Convert.ToChar(Amf_array[181]), Convert.ToChar(Amf_array[180]));
                            Taxi_ID_temp = carnum;


                            string Model = "";
                            if (Amf_array.Length > 256 + 64)
                            {
                                for (int i = 0; i < 9; i++)
                                {
                                    if (Amf_array[180 + i] < 0x20)
                                    {
                                        Amf_array[180 + i] = 0x20;
                                    }
                                }






                                Model = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}"
                                                      , Convert.ToChar(Amf_array[240]), Convert.ToChar(Amf_array[241]), Convert.ToChar(Amf_array[242]), Convert.ToChar(Amf_array[243]), Convert.ToChar(Amf_array[244]));
                                //  strMN += String.Format("{0:C}", Convert.ToChar(newProInHeader.ModelName[i]));
                            }

                            //  string TmpFile = Application.StartupPath + "\\" + NowReceiveTime + ".TMF";

                            if (Model == "\0\0\0\0\0")
                            {
                                Model = "Pro1+";
                            }
                            Model = "Pro1+";

                            string newPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "AMF");
                            // Create the subfolder
                            System.IO.Directory.CreateDirectory(newPath);

                            string TmpFile = TACHO2_path + "\\TMF\\AMF\\" + Model + "_" + NowReceiveTime + "_" + carnum + ".AMF";

                            Amf_path = TmpFile;

                            // byte[] rcvByte = new byte[mStreamBuffer.Length];
                            // rcvList.CopyTo(rcvByte);
                            ///////////////////////////////////////////////////////////////////////////
                            /*  byte Checksum = 0;

                              for (int i = 0; i < Amf_array.Length - 2; i++)
                              {
                                  Checksum += Amf_array[i];
                              }*/


                            FileStream fs = new FileStream(TmpFile, FileMode.OpenOrCreate, FileAccess.Write);
                            BinaryWriter bw = new BinaryWriter(fs);

                            bw.Write(Amf_array);



                            fs.Close();
                            bw.Close();



                            AMF_Data(TmpFile);







                            Thread.Sleep(1000);
                            lock (this)
                            {
                                Driver_Receipt();
                            }
                        }
                        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        if (CashierMode == true)
                        {
                            if (NetworkInterface.GetIsNetworkAvailable())
                            {

                                //  string vAlarmOffSQL = "UPDATE TblTacho SET rele_date = '{4}' WHERE mach_code='{0}' and pipe_code='{1}' and erro_code='{2}' and erro_date=#{3}#";

                                DateTime DbTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day); //현재 시간 

                                string FolderName = "TACHO\\" + DbTime.Year.ToString() + "\\" + DbTime.Month.ToString();

                                string newPath = TACHO2_path + FolderName;

                                newPath = @"\\" + ShareIP + "\\tacho2\\" + FolderName;


                                NameDB = newPath + "\\" + MdbName;
                                DBstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + NameDB;
                                //					 Db_backup = false;
                                conn = new OleDbConnection(@DBstring);

                                conn.Open();


                                //   string  strSQL = "SELECT * FROM Table WHERE Time BETWEEN #" + strStart + "# AND #" + strEnd + "#";

                                commRead = new OleDbCommand(queryRead, conn);
                                srRead = commRead.ExecuteReader();

                                string TaxiID_ = "";
                                string DriverID_ = "";
                                string DriverName_ = "";
                                int ID_ = 0;
                                DateTime Outtime_ = new DateTime(1, 1, 1, 1, 1, 1);
                                DateTime Intime_ = new DateTime(1, 1, 1, 1, 1, 1);


                                while (srRead.Read())
                                {
                                    TaxiID_ = srRead.GetString(1);

                                    if (TaxiID_ == TaxiID)
                                    {


                                        DriverID_ = srRead.GetString(2);
                                        if (DriverID_ == DriverID)
                                        {
                                            DriverName_ = srRead.GetString(3);
                                            if (DriverName_ == DriverName)
                                            {
                                                Outtime_ = srRead.GetDateTime(4);
                                                if (Outtime_ == Outtime)
                                                {
                                                    Intime_ = srRead.GetDateTime(5);
                                                    if (Intime_ == Intime)
                                                    {
                                                        ID_ = srRead.GetInt32(0);
                                                    }
                                                }
                                            }
                                        }



                                    }
                                }


                                //  nowTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

                                queryUpdate = "UPDATE TblTacho SET " + "PaidTime" + "='" + nowTime
                                 + "' WHERE ID=" + ID_.ToString();

                                commUpdate = new OleDbCommand(queryUpdate, conn);

                                commUpdate.ExecuteNonQuery();


                                queryUpdate = "UPDATE TblTacho SET " + "RealIncome" + "='" + textBox3.Text
                                + "' WHERE ID=" + ID_.ToString();
                                commUpdate = new OleDbCommand(queryUpdate, conn);

                                commUpdate.ExecuteNonQuery();

                                conn.Close();

                                if (isBackup == true)
                                {
                                    // collect한 백업을 데이터를 저장한다.  아직 서버에 전송전이기때문에
                                    // 업데이트 필드에 대한 정보를 파일에 메모!!  -> 차량번호 기사 번호, 이름 , 입, 출고 시간을 기록하자. 후에 네트워크 정상시 다시 업데이트 
                                }


                            }
                            else
                            {

                                // 서버에 연결 할수 없음으로 불가능!!

                                // 업데이트 필드에 대한 정보를 파일에 메모!!  -> 차량번호 기사 번호, 이름 , 입, 출고 시간을 기록하자. 후에 네트워크 정상시 다시 업데이트 
                            






                                /*
                             //   string BackupPath = System.IO.Path.Combine(TACHO2_path + "\\TMF\\", "BACKUP");
                                
                                // Create the subfolder
                             //   System.IO.Directory.CreateDirectory(BackupPath);
                                string newPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "BACKUP");
                                // Create the subfolder
                                System.IO.Directory.CreateDirectory(newPath);

                                string TmpFile = TACHO2_path + "\\TMF\\BACKUP\\PaidTime.PTF";

                                byte[] paid_array = new byte[48];

                                byte[] carnum_char = new byte[TaxiID.Length];
                             
                            //    paid_array[0] = TaxiID

                                // byte[] rcvByte = new byte[mStreamBuffer.Length];
                                // rcvList.CopyTo(rcvByte);

                                FileStream fs = new FileStream(TmpFile, FileMode.OpenOrCreate, FileAccess.Write);
                                BinaryWriter bw = new BinaryWriter(fs);

                                bw.Write(paid_array);



                                fs.Close();
                                bw.Close();
                                */
                               

                            }
                        }
                    

                    }
                    catch (Exception e)
                    {

                    }
                


                    pictureBox17.Visible = false;

                    DB_ReadData(0, 1);

                    if (ServerMessage != "")
                    {
                        this.richTextBox1.Visible = true;
                        button16.Enabled = false;
                        button5.Visible = true;
                        richTextBox1.Text = "------------------------------------------ From. Office Message --------------------------------------------";
                        richTextBox1.Text += "\n\n   To. " + Drivername_textBox.Text +"\n\n";
                        richTextBox1.Text += "  " +ServerMessage;
                      //  MessageBox.Show(ServerMessage);
                        ServerMessage = "";
                    }

              
            
        }
        
        public void Run_Ibutton()
        {

            byte status = 0;

            //GetTime_Receipt
         
            while (true)
            {

                if (IbuttonSetting == true)
                {
                    Thread.Sleep(100);
                    continue;
                }
              
                if (this.Visible == false)
                {
                    Thread.Sleep(100);
                    continue;
                }

                if (PaidTimeWrite == true)
                {
                    Thread.Sleep(1000);
                    lock (this)
                    {
                        PaidTime_Write();
                        PaidTimeWrite = false;
                    }
                    Thread.Sleep(1000);
                    Client_label.Visible = false;
                    Thread.Sleep(10);
                    continue;
                }
                Thread.Sleep(100);
                try
                {
                  //  switch (status)
                  //  {
                      //  case 0:

                                           
                            /////////////////////////////////////////////////////////////////
                                short[] ROM = new short[8];
                                //      byte[] block = new byte[322];
                                byte[] TachoAddr = new byte[6]; // 타코 시작 주소와 끝 주소를 찾는다. 
                                Int16 StartAddr = 0;
                                Int16 EndAddr = 0;
                                int n = 0;
                                byte[] Address = new byte[2];
                                byte[] TachoStartByte = new byte[1];
                                byte[] TachoOverFlow = new byte[1];
                                short ret = 0;
                                short portNum = 0;
                                short portType = 0;
                                int hSess = 0;
                                bool errorstatus = false;

                              

                            TMEX TMEXLibrary = new TMEX();
                            TMEXLibrary.TMReadDefaultPort(ref portNum, ref portType);
                       
                            hSess = TMEXLibrary.TMExtendedStartSession(portNum, portType, ref sessionOptions);
                            
                            if (hSess != 0)
                            {

                                // must be called before any non-session functions can be called
                                TMEXLibrary.TMSetup(hSess);  // 1


                                ///// Check Ibutton
                                ret = TMEXLibrary.TMFirst(hSess, stateBuffer);

                                if (ret != 1)
                                {
                                    IbuttonReadCheck = false;
                                    errorstatus = true;
                                   // MessageBox.Show("Disconnected  to ibutton.");
                                  //  return;
                                }

                                n = 0;


                                TMEXLibrary.TMRom(hSess, stateBuffer, ROM);  // Selecet  2
                                ret = TMEXLibrary.TMNext(hSess, stateBuffer); // 3

                                n = TMEXLibrary.TMStrongAccess(hSess, stateBuffer);

                                if (ROM[0] != 0x0c)
                                {
                                    // continue;
                                    label32.Text = "";
                                    Ibutton_touched = false;   // 아이버튼 연결이 되지 않음
                                    Ibutton_newButton = true;  // 새로 인식할수 있도록 셋 해줌
                                    IbuttonReadCheck = false;
                                    errorstatus = true;
                                    // MessageBox.Show("Disconnected  to ibutton.");
                                    //  ibutonCheck = true;
                                    //   return;
                                }
                                else
                                {
                                    label32.Text = "iButton : Present";
                                    Ibutton_touched = true;
                                }
                            }


                          


                            if (Ibutton_touched == true && Ibutton_newButton == true)
                            {

                                pictureBox17.Visible = true;
                                if (Ibutton_newButton == true)
                                {
                                    Ibutton_newButton = false;
                                }
                                TMEXLibrary.TMSetup(hSess);

                                ret = TMEXLibrary.TMFirst(hSess, stateBuffer);


                                if (ROM[0] != 0x0c)
                                {
                                    Ibutton_touched = false;
                                    Ibutton_newButton = true;
                                    continue;
                                    IbuttonReadCheck = false;
                                    errorstatus = true;
                                    // MessageBox.Show("Disconnected  to ibutton.");
                                    // this.Close();
                                    //  return;
                                    //  this.Close();
                                }


                                if (ret != 1)
                                {
                                    errorstatus = true;
                                    // MessageBox.Show("Ibutton Fail");
                                    // return;
                                }

                                //Daily_Distance////////////////////////////////////////////////////////


                                Address[0] = 0x5B;
                                Address[1] = 0x00;
                                Daily_Distance = OnWire_Read(Daily_Distance, Daily_Distance.Length, Address, TMEXLibrary, hSess, ROM);

                                /////////////////////////////////////////////  타코 Start Byte 읽기
                                Address[0] = 0x00;
                                Address[1] = 0x00;

                                TachoStartByte = OnWire_Read(TachoStartByte, TachoStartByte.Length, Address, TMEXLibrary, hSess, ROM);

                                if (TachoStartByte[0] == 0xA3)
                                {
                                    MessageBox.Show("Tacho data has not yet been Logout.");
                                    pictureBox17.Visible = false;
                                    IbuttonReadCheck = false;   // 아이버튼을 읽는중 아이버튼을 제거 하여 읽다가 데이터 끝까지 못읽었을경우 
                                    continue;
                                }
                                
                                /////////////////// 0x11 -> 1 memory overflow 메모리 전체 읽는다.\\\\\\\

                                Address[0] = 0x11;
                                Address[1] = 0x00;

                                TachoOverFlow = OnWire_Read(TachoOverFlow, TachoOverFlow.Length, Address, TMEXLibrary, hSess, ROM);



                                /////////////////////////////////////////////////////////////////////

                                ///////////////////////////// 타코 길이 계산하기 

                                Address[0] = 0x40;
                                Address[1] = 0x00;


                                TachoAddr = OnWire_Read(TachoAddr, TachoAddr.Length, Address, TMEXLibrary, hSess, ROM);   // 타코 길이 알기 위해 주소를 타코 저장 위치 주소를 읽느나.

                                if (TachoAddr[0] == 0xff && TachoAddr[1] == 0xff && TachoAddr[2] == 0xff && TachoAddr[3] == 0xff && TachoAddr[4] == 0xff)
                                {
                                    IbuttonReadCheck = false;
                                    errorstatus = true;
                                    //  MessageBox.Show("Disconnected  to ibutton.");
                                    //  return;
                                }

                                EndAddr = (Int16)(TachoAddr[1] << 8 & 0xff00);
                                EndAddr += TachoAddr[0];

                                StartAddr = (Int16)(TachoAddr[4] << 8 & 0xff00);
                                StartAddr += TachoAddr[3];

                                StartAddr = 0x100;  // 시작 주소는 고정으로 사용 !

                                int num = EndAddr - StartAddr;  // 타코 트랜젝션의 갯수를 구한다. 


                                if (EndAddr == 0x00)
                                {

                                    IbuttonReadCheck = false;
                                    errorstatus = true;
                                    pictureBox17.Visible = false;
                                    MessageBox.Show("Empty Tacho!");
                                    //   return;
                                }



                                ///////////////////////////////////////////////////////////

                                if (errorstatus == false && IbuttonReadCheck == false)
                                {
                                    byte[] Amf_array;  // AMf_file 총 싸이즈 data +headr +0xfd + checkSum

                                    if (TachoOverFlow[0] == 0x01)  // Memory overflow 발생 
                                    {
                                        Amf_array = new byte[8194]; // 8192 +2(0xfd + checkSum)
                                    }
                                    else
                                    {
                                        Amf_array = new byte[num + 256 + 2]; // transaction length.
                                    }


                                    Address[0] = 0x00;
                                    Address[1] = 0x00;



                                    Amf_array = OnWire_Read(Amf_array, Amf_array.Length, Address, TMEXLibrary, hSess, ROM); // data Read

                                    int checkcnt = 0;
                                    for (int i = 0; i < 256; i++) //data 검사 
                                    {

                                        if (Amf_array[i] == 0xff)
                                        {
                                            checkcnt++;
                                        }
                                    }

                                    if (checkcnt == 256)
                                    {
                                        pictureBox17.Visible = false;
                                        continue;
                                    }
                                   
                                    //////////////////////////////////////////////  data 검사 15.08.17
                                    int index = 256;
                                    int TachoLength = num / 64;

                                    bool ErrCheck = false;
                                    while (TachoLength != 0)
                                    {
                                        checkcnt = 0;
                                        for (int i = 0; i < 64; i++)
                                        {
                                            if (Amf_array[index] == 0xff)
                                            {
                                                checkcnt++;
                                            }
                                            index++;
                                            if (checkcnt > 5)
                                            {
                                                ErrCheck = true;
                                               // continue;
                                            }
                                        }
                                        TachoLength--;
                                    }
                                    if (ErrCheck == true)
                                    {
                                        pictureBox17.Visible = false;
                                        continue;
                                    }
                                    ///////////////////////////////////////////////


                                    string NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}",
                                                                                (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);


                                    string carnum = "";
                                    carnum = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}{8:C2}"
                                                               , Convert.ToChar(Amf_array[188]), Convert.ToChar(Amf_array[187]), Convert.ToChar(Amf_array[186]), Convert.ToChar(Amf_array[185])
                                                               , Convert.ToChar(Amf_array[184]), Convert.ToChar(Amf_array[183])
                                                            , Convert.ToChar(Amf_array[182]), Convert.ToChar(Amf_array[181]), Convert.ToChar(Amf_array[180]));
                                    Taxi_ID_temp = carnum;

                              
                                    string Model = "";
                                    if (Amf_array.Length > 256 + 64)
                                    {
                                        for (int i = 0; i < 9; i++)
                                        {
                                            if (Amf_array[180 + i] < 0x20)
                                            {
                                                Amf_array[180 + i] = 0x20;
                                            }
                                        }

                                      


                                       

                                        Model = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}"
                                                              , Convert.ToChar(Amf_array[240]), Convert.ToChar(Amf_array[241]), Convert.ToChar(Amf_array[242]), Convert.ToChar(Amf_array[243]), Convert.ToChar(Amf_array[244]));
                                        //  strMN += String.Format("{0:C}", Convert.ToChar(newProInHeader.ModelName[i]));
                                    }

                                    //  string TmpFile = Application.StartupPath + "\\" + NowReceiveTime + ".TMF";

                                    if (Model == "\0\0\0\0\0")
                                    {
                                        Model = "Pro1+";
                                    }
                                    Model = "Pro1+";

                                    string newPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "AMF");
                                    // Create the subfolder
                                    System.IO.Directory.CreateDirectory(newPath);

                                    string TmpFile = TACHO2_path + "\\TMF\\AMF\\" + Model + "_" + NowReceiveTime + "_" + carnum + ".AMF";

                                    Amf_path = TmpFile;

                                    // byte[] rcvByte = new byte[mStreamBuffer.Length];
                                    // rcvList.CopyTo(rcvByte);
                                    ///////////////////////////////////////////////////////////////////////////
                                  /*  byte Checksum = 0;

                                    for (int i = 0; i < Amf_array.Length - 2; i++)
                                    {
                                        Checksum += Amf_array[i];
                                    }*/


                                    FileStream fs = new FileStream(TmpFile, FileMode.OpenOrCreate, FileAccess.Write);
                                    BinaryWriter bw = new BinaryWriter(fs);

                                    bw.Write(Amf_array);



                                    fs.Close();
                                    bw.Close();



                                    AMF_Data(TmpFile);
                                    
                                    
                                 /*   if (MiniPrintCon == true)
                                    {
                                        Thread.Sleep(1000);
                                        Driver_Receipt();
                                    }*/

                                    if (CashierMode == true)   // 네트 워크 체크 기능이 필요하다. 오류시 백업모드로 변환하자. 
                                    {
                                        if (NetworkInterface.GetIsNetworkAvailable())
                                        {
                                            ManagerAmfparsing(TmpFile);
                                            //  MessageBox.Show("네트워크 사용가능");

                                            /******* 백업 파일 이 존재 한다면 네트워크 DB에 파싱 **************
                                             파싱후 날짜시간폴더를 만들어 보관********************************/


                                        }
                                        else
                                        {

                                           // BackupAmfparsing(TmpFile);


                                            /***************AMF file 백업기능 ********************/

                                            isBackup = true;  // 백업발생 체크 
                                            
                                            string BackupPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "BACKUP");
                                            // Create the subfolder
                                            System.IO.Directory.CreateDirectory(newPath);

                                            string BackupFile = TACHO2_path + "\\TMF\\BACKUP\\" + Model + "_" + NowReceiveTime + "_" + carnum + ".AMF";

                                          
                                             fs = new FileStream(BackupFile, FileMode.OpenOrCreate, FileAccess.Write);
                                             bw = new BinaryWriter(fs);

                                            bw.Write(Amf_array);


                                            fs.Close();
                                            bw.Close();
                                            
                                            //////////////////////////////////////////////////////////////////////////////
                                            // BackupDB? 사용?


                                        }
                                       
                                    }


                                    //////////////////////////////////////   2번째 저장 total tmf  기존 데이터와  이어 붙히자 !!///////////////////

                                    string TMFPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "TransData");
                                    // Create the subfolder
                                    System.IO.Directory.CreateDirectory(TMFPath);

                                    NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                                            (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day);
                                    TmpFile = TACHO2_path + "\\TMF\\TransData\\" + NowReceiveTime + ".AMF";

                                    // rcvList.RemoveAt(0);
                                    //  rcvByte = new byte[rcvList.Count];
                                    //    rcvList.CopyTo(rcvByte);

                                    fs = new FileStream(TmpFile, FileMode.Append, FileAccess.Write);
                                    bw = new BinaryWriter(fs);

                                    bw.Write(Amf_array);
                                    fs.Close();
                                    bw.Close();

                                  
                                    if ( TachoStartByte[0] == 0xA4)
                                    {
                                        TachoStartByte[0] = 0xff;

                                        Address[0] = 0x00;
                                        Address[1] = 0x00;
                                        OneWire_Write(TachoStartByte, TachoStartByte.Length, Address, TMEXLibrary, hSess);
                                    }

                                    if (IbuttonReadCheck == true)
                                    {
                                        IbuttonReadCheck = false;   // 아이버튼을 읽는중 아이버튼을 제거 하여 읽다가 데이터 끝까지 못읽었을경우 
                                        pictureBox17.Visible = false;
                                        continue;
                                    }
                                    IbuttonReadCheck = true;
                                    //   MessageBox.Show("Sucessfully received tacho data! ");

                                    byte[] TimeLimitAddr = new byte[2];
                                    TimeLimitAddr[0] = 0x20;
                                    TimeLimitAddr[1] = 0x00;


                                    TimLimitByte[0] = 0x00;
                                    TimLimitByte[1] = 0x00;


                                    byte[] temp = new byte[2];






                                    OneWire_Write(TimLimitByte, TimLimitByte.Length, TimeLimitAddr, TMEXLibrary, hSess);
                                    pictureBox17.Visible = false;
                                 
                                    /*
                                    /////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                    // 서버에 접속하여 Paid Time을 설정해야함

                                    try
                                    {
                                        if (bConnectServer)
                                        {
                                          //  strLog = string.Format("[SYSTEM] : 현재 서버와 접속중입니다.!!");
                                            //    Add_Log(strLog);
                                          //  return;
                                        }
                                        Client_label.Visible = true;
                                        Client_label.Text = "Connecting to the server ...";

                                        EY_ChatClient = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.IP);
                                        IPAddress ipServer = IPAddress.Parse(ServerIP);

                                        //  while (true)
                                        // {
                                        try
                                        {
                                            EY_ChatClient.Connect(new IPEndPoint(ipServer, intPortNum));


                                            if (EY_ChatClient.Connected)
                                            {
                                                // strLog = string.Format("[SYSTEM] : 서버 접속 성공!!");
                                                //  Add_Log(strLog);
                                                bConnectServer = true;

                                                //  pictureBox1.Visible = true;

                                                //  label7.Text = "서버 접속중";
                                                byte CheckSum = 0;

                                                byteSendMsg[0] = 0x55;
                                                byteSendMsg[1] = 0x15;

                                                byteSendMsg[2] = Amf_array[180];  // Taxi Id 
                                                byteSendMsg[3] = Amf_array[181];
                                                byteSendMsg[4] = Amf_array[182];
                                                byteSendMsg[5] = Amf_array[183];
                                                byteSendMsg[6] = Amf_array[184];
                                                byteSendMsg[7] = Amf_array[185];
                                                byteSendMsg[8] = Amf_array[186];
                                                byteSendMsg[9] = Amf_array[187];
                                                byteSendMsg[10] = Amf_array[188];


                                                byteSendMsg[11] = Amf_array[192];  // Driver ID
                                                byteSendMsg[12] = Amf_array[193];
                                                byteSendMsg[13] = Amf_array[194];
                                                byteSendMsg[14] = Amf_array[195];
                                                byteSendMsg[15] = Amf_array[196];
                                                byteSendMsg[16] = Amf_array[197];
                                                byteSendMsg[17] = Amf_array[198];
                                                byteSendMsg[18] = Amf_array[199];
                                                byteSendMsg[19] = Amf_array[200];

                                                for (int a = 0; a < 19; a++)
                                                {
                                                    CheckSum += byteSendMsg[a];
                                                }
                                                byteSendMsg[20] = CheckSum;
                                                byteSendMsg[21] = 0xFD;


                                                // 접속 메세지 전송
                                                EY_ChatClient.BeginSend(byteSendMsg, 0, 22, SocketFlags.None, new AsyncCallback(CallBack_SendMsg), EY_ChatClient);
                                                //이제 메시지를 전송받는다.
                                                EY_ChatClient.BeginReceive(byteReceiveMsg, 0, byteReceiveMsg.Length, SocketFlags.None, new AsyncCallback(CallBack_ReceiveMsg), byteReceiveMsg);
                                                //ReceiveStart();
                                                // break;

                                              
                                                Client_label.Text = "Complete access to the server !";
                                                Thread.Sleep(1000);
                                            }
                                        }
                                        catch (Exception err)
                                        {
                                            // strErr = string.Format("[SYSTEM] : {0}", err.Message);
                                            //   Add_Log(strErr);
                                            ///////////////////////////////////////////////////////////////
                                            pictureBox17.Visible = false;
                                            Client_label.Text = "Server connection fails!!";
                                           MessageBox.Show(err.Message);
                                            bConnectServer = false;
                                            Client_label.Visible = false;
                                        }
                                        //  }

                                    }
                                    catch (Exception e)
                                    {
                                     
                                    }




                                    ////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                     */

                                }
                                  //pictureBox17.Visible = false;
                            }
                          
                           // break;
                    
                  //  }
                }
                catch (Exception ex)
                {
                  
                    MessageBox.Show(ex.Message);
                    continue;
                }


            
            

            }
        }

        private void CallBack_ReceiveMsg(IAsyncResult ar)
        {
            byte[] bytes = (byte[])ar.AsyncState;
            int datalen = 0;

            if (EY_ChatClient.Connected == true)
            {
                intSize = EY_ChatClient.EndReceive(ar);
                strReceiveMsg = "";
                strReceiveMsg = Encoding.Default.GetString(bytes, 0, intSize);
                EY_ChatClient.BeginReceive(byteReceiveMsg, 0, byteReceiveMsg.Length, SocketFlags.None, new AsyncCallback(CallBack_ReceiveMsg), byteReceiveMsg);
               
             //   if (bytes[0] == 0x55 && bytes[1] == 0xA5 && bytes[5] ==0xFD)
                if (bytes[0] == 0x55 && bytes[1] == 0xA5)
                {

                    datalen = bytes[2];

                    if (bytes[datalen - 1] == 0xFD)
                    {
                        TimLimitByte[0] = bytes[3];
                        TimLimitByte[1] = bytes[4];

                        for (int i = 0; i < datalen - 7; i++)
                        {
                            ServerMessage += (char)bytes[i + 5];
                        }

                        PaidTimeWrite = true;
                    }
                }
                else if(bytes[0] == 0x55 && bytes[1] == 0xB5 && bytes[5] ==0xFD)
                {
                    TimLimitByte[0] = 0x00;
                    TimLimitByte[1] = 0x00;
                    PaidTimeWrite = true;
                    MessageBox.Show("Data Error!");
                }
               



            }
        }
        private void CallBack_SendMsg(IAsyncResult ar)
        {
            EY_ChatClient = (Socket)ar.AsyncState;

            try
            {
                intSize = EY_ChatClient.EndSend(ar);
                if (intSize == 0)
                {
                    Disconnet();
                }
                else
                {
                    strSendMsg = Encoding.Default.GetString(byteSendMsg, 0, intSize);
                    EY_ChatClient.BeginReceive(byteReceiveMsg, 0, byteReceiveMsg.Length, SocketFlags.None, new AsyncCallback(CallBack_ReceiveMsg), byteReceiveMsg);
                }
            }
            catch (Exception err)
            {
                strErr = string.Format("[SYSTEM] : {0}", err.Message);
                //   MessageBox.Show(err.Message);
                //    Add_Log(strErr);
                Disconnet();
            }
        }
  
        public void ClientMessage_Process(Socket socTClient)
        {
            /*
            lock (this)
            {

                try
                {

                    if (que.Count > 0)
                    {



                        intCientID = (int)que.Dequeue();

                        strClientID = intCientID.ToString();

                        if (!IDList.Contains(strClientID))
                        {
                            IDList.Add(intCientID.ToString());
                        }
                        else
                        {
                            que.Enqueue(intCientID);
                            strClientID = "@";
                            return;
                        }



                        if (intCientID == 1)
                        {
                            eyChild[0] = new EY_ChatChild(socTClient, strClientID, this);
                        }
                        else if (intCientID == 2)
                        {
                            eyChild[1] = new EY_ChatChild(socTClient, strClientID, this);
                        }
                        else if (intCientID == 3)
                        {
                            eyChild[2] = new EY_ChatChild(socTClient, strClientID, this);
                        }
                        else if (intCientID == 4)
                        {
                            eyChild[3] = new EY_ChatChild(socTClient, strClientID, this);
                        }


                        if (intCientID == 1)
                        {
                            client1_Visible(true);
                            // Client1_Icon.Visible = true;
                            //    eyChild1 = new EY_ChatChild(socTClient, strClientID, this);
                            //      ClientList.Add(strClientID, eyChild1);




                        }
                        else if (intCientID == 2)
                        {

                            client2_Visible(true);
                            // Client2_Icon.Visible = true;
                            //   eyChild2 = new EY_ChatChild(socTClient, strClientID, this);
                            //   ClientList.Add(strClientID, eyChild2);

                        }
                        else if (intCientID == 3)
                        {
                            client3_Visible(true);
                            //Client3_Icon.Visible = true;
                            //    eyChild3 = new EY_ChatChild(socTClient, strClientID, this);
                            //   ClientList.Add(strClientID, eyChild3);

                        }
                        else if (intCientID == 4)
                        {
                            client4_Visible(true);
                            // Client4_Icon.Visible = true;
                            //    eyChild4 = new EY_ChatChild(socTClient, strClientID, this);
                            //    ClientList.Add(strClientID, eyChild4);

                        }




                    }
                    else
                    {
                        strClientID = "0";
                    }
                }
                catch (Exception ex)
                {
                    //   MessageBox.Show(ex.Message);
                }
            }
             * */


        }

        public void Thread_Run()
        {

            svr = new TcpServerSocket(true);
            svr.SetForm1(this);
            svr.Bind();
            svr.Listen();

            svr.AcceptStart();

            label32.Text = "SERVER OPEN";

          //  button10.Visible = true;
            while (true)
            {
                svr.Update();
                Thread.Sleep(300);
            }

            /*
            ServerSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.IP);
            ServerSocket.Bind(new IPEndPoint(IPAddress.Any, intPortNum));
            ServerSocket.Listen(9);

            IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
            IPAddress ipAddr = host.AddressList[0];
      

            label32.Text = "SERVER OPEN";
     
            button10.Visible = true;

          

            while (bSocketEnd == false)
            {


                try
                {
                    //  Socket socClient = (Socket)ServerSocket.Accept();
                    ClientSocket = (Socket)ServerSocket.Accept();

                    ClientMessage_Process(ClientSocket);



                    byte[] Sendbuff = new byte[4096];

                    Sendbuff[0] = (byte)Int32.Parse(strClientID);
                    Sendbuff[1] = 0x00;


                    ClientCnt++;
                    strLog = string.Format("[{0}]?묒냽", strClientID);
                    //     Add_Log(strLog);
                    // label8.Text = strLog;

                    if (strClientID == "1")
                    {
                        client1_Visible(true);
                        //  Client1_Icon.Visible = true;
                        //   label8.Text = strLog;
                    }
                    else if (strClientID == "2")
                    {
                        client2_Visible(true);
                        // Client2_Icon.Visible = true;
                        //    label9.Text = strLog;
                    }
                    else if (strClientID == "3")
                    {
                        client3_Visible(true);
                        // Client3_Icon.Visible = true;
                        //   label10.Text = strLog;
                    }
                    else if (strClientID == "4")
                    {
                        client4_Visible(true);
                        // Client4_Icon.Visible = true;

                        //   label11.Text = strLog;
                    }
                    else if (strClientID == "0")
                    {
                        Sendbuff[0] = 0xFD;
                    }
                    else if (strClientID == "@")
                    {
                        Sendbuff[0] = 0xFE;
                    }

                    ClientSocket.BeginSend(Sendbuff, 0, byteSendMsg.Length, SocketFlags.None, new AsyncCallback(CallBack_SendMsg), ClientSocket);
                }
                catch (Exception err)
                {

                    MessageBox.Show(err.Message);
               

                }


            }*/
        }
        private void Disconnet()
        {
          
        }
        public void SendMsg(string id, string msg) //메시지를 보낸다.
        {
            lock (this)
            {

            
               
                /*
                    if (id == "1")
                    {
                        eyChild[0].ClientSend(msg);
                    }
                    else if (id == "2")
                    {
                        eyChild[1].ClientSend(msg);
                    }
                    else if (id == "3")
                    {
                        eyChild[2].ClientSend(msg);
                    }
                    else if (id == "4")
                    {
                        eyChild[3].ClientSend(msg);
                    }
              */



            }

        }
      
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {


            if (getSerialData != null)
            {
                getSerialData.Abort();
            }
            if (getSerialData2 != null)
            {
                getSerialData2.Abort();
            }
            if (getSerialData3 != null)
            {
                getSerialData3.Abort();
            }
            if (getSerialData4 != null)
            {
                getSerialData4.Abort();
            }

            if (ServerThread != null)
            {
                ServerThread.Abort();
            }

            if (serialPort1.IsOpen)
            {
                serialPort1.Close();
            }

            if (Ibuttonthread != null)
            {
                Ibuttonthread.Abort();
            }


        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort2.IsOpen)
                    serialPort2.Close();


            }
            catch
            {
                MessageBox.Show("Can't close already opened port", "Error");
                return;
            }

            serialPort2.PortName = comboBox2.SelectedItem.ToString();
            serialPort2.BaudRate = Convert.ToInt32(comboBox1.SelectedItem);

            try
            {
                serialPort2.Open();

                label14.Text = comboBox2.SelectedItem.ToString() + " Open";

                getSerialData2 = new Thread(new ThreadStart(Run_SerialThread2));
                getSerialData2.IsBackground = true;
                Thread.Sleep(100);
                getSerialData2.Start();


            }
            catch
            {
                MessageBox.Show("Can't open port", "Error");
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort2.IsOpen)
                    serialPort2.Close();
                label14.Text = comboBox2.SelectedItem.ToString() + " Close";
                
                label10.Text = "NORMAL";
                label10.ForeColor = Color.Black;

                pictureBox4_Visible(true);
                pictureBox3_Visible(false);

             

                pictureBox11_Visible(true);
                pictureBox12_Visible(false);



            }
            catch
            {
                MessageBox.Show("Can't close already opened port", "Error");
                return;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort3.IsOpen)
                    serialPort3.Close();


            }
            catch
            {
                MessageBox.Show("Can't close already opened port", "Error");
                return;
            }

            serialPort3.PortName = comboBox4.SelectedItem.ToString();
            serialPort3.BaudRate = Convert.ToInt32(comboBox3.SelectedItem);

            try
            {
                serialPort3.Open();

                label22.Text = comboBox4.SelectedItem.ToString() + " Open";

                getSerialData3 = new Thread(new ThreadStart(Run_SerialThread3));
                getSerialData3.IsBackground = true;
                Thread.Sleep(100);
                getSerialData3.Start();


            }
            catch
            {
                MessageBox.Show("Can't open port", "Error");
                return;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (serialPort3.IsOpen)
                    serialPort3.Close();
                label22.Text = comboBox4.SelectedItem.ToString() + " Close";

                label19.Text = "NORMAL";
                label19.ForeColor = Color.Black;

                pictureBox5_Visible(false);
                pictureBox6_Visible(true);

                pictureBox13_Visible(false);
                pictureBox14_Visible(true);

            }
            catch
            {
                MessageBox.Show("Can't close already opened port", "Error");
                return;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
          
        }

        private void button5_Click(object sender, EventArgs e)
        {
        }

        public void Tacho_Run(string str)
        {
           
            lock (this)
            {
               
                //  Thread.Sleep(1000);
                if (pasingcnt == 4)
                {
                    pasingcnt = 0;
                }
                pasingcnt++;

                //  Worker workerObject = new Worker();
                // Thread workerThread = new Thread(workerObject.DoWork);
                //   workerThread.Start();

                bool Last_br = false;
                bool bOpenFileData = true;
                int IndexNumber = 0;
                long[] ID_Number = new long[500];


                DateTime standardTime = new DateTime(1, 1, 1, 0, 0, 0);

                string strFname;
                long oldFdPosition = 0, newFdPosition = 0;
                int bReadData = 0;
                byte[] Endbr = new byte[381];
                byte[][] bArrEmergencyData = new byte[16][];

                strFname = str;
                FileStream fs = new FileStream(strFname, FileMode.Open, FileAccess.ReadWrite);
                BinaryReader br = new BinaryReader(fs);
                BinaryWriter bw = new BinaryWriter(fs);

                if (fs.Length == 0 || (fs.Length < 384 ))
                {
                    // workerObject.RequestStop();
                    return;
                }
                ////  Worker workerObject = new Worker();
                //  Thread workerThread = new Thread(workerObject.DoWork);
                //  workerThread.Start();


                int data_start = 0;
                bool Last_data_ff = false;
                /////////////////////////////////////////////////   12. 03.21  :   sd에서 데이터 추출할경우 파일끝이 0xf5로 끝나는 경우 파일끝의 
                //   0xf5-> 0xf7으로 변환후 파싱을 시작한다.

                if (fs.Length == 0)
                {
                    return;
                }

                fs.Position = fs.Length;

                fs.Position -= 1;

                byte endfiletemp = 0;
                endfiletemp = br.ReadByte();

                if (endfiletemp == 0xf5)
                {
                    fs.Position -= 1;
                    bw.Write(0xf7);
                }

                fs.Position = 0;

                /////////////////////////////////////////////////

                try
                {

                    NewFA = false;

                    // Fill Struct from Ram data
                    int yy, mon, dd, hh, mm, ss;

                    byte[] temp = new byte[8];
                    TachoRamData stTachoRamData = new TachoRamData();
                    TachoDataCode stTachoDataCode = new TachoDataCode();

                    #region Initialized struct value
                    {
                        stTachoDataCode.moneyTblTacho = 0;
                        stTachoDataCode.salesKmTblTacho = 0;
                        stTachoDataCode.driveDistanceTblTacho = 0;
                        stTachoDataCode.overrunTime = new DateTime(1, 1, 1, 0, 0, 0);
                        stTachoDataCode.emerBreakCnt = 0;
                        stTachoDataCode.driveBasicCnt = 0;
                        stTachoDataCode.driveAfterCnt = 0;
                        stTachoDataCode.premiumBasicCnt = 0;
                        stTachoDataCode.premiumAfterCnt = 0;
                        stTachoDataCode.doorOpenCnt = 0;
                        stTachoDataCode.SalesTotalMoney = 0;

                        stTachoDataCode.driveCount = 0;
                        stTachoDataCode.yymmdd = new DateTime(2010, 1, 1, 0, 0, 0);
                        stTachoDataCode.beforeTime = new DateTime(2010, 1, 1, 0, 0, 0);
                        stTachoDataCode.afterTime = new DateTime(2010, 1, 1, 0, 0, 0);
                        stTachoDataCode.salesKm = 0;
                        stTachoDataCode.money = 0;
                        stTachoDataCode.empty = 0;
                        stTachoDataCode.emptyTime = 0;
                        stTachoDataCode.notuse = false;
                        stTachoDataCode.add = false;
                        stTachoDataCode.key = false;
                        stTachoDataCode.emerBreak = false;
                        stTachoDataCode.emerTime = new DateTime(2010, 1, 1, 0, 0, 0);
                        stTachoDataCode.emptyStartTime = new DateTime(2010, 1, 1, 0, 0, 0);
                        stTachoDataCode.emerSpeed = 0;

                        // '10. 7.19 추가
                        stTachoDataCode.celldriveBasicCnt = 0;
                        stTachoDataCode.celldriveAfterCnt = 0;
                        stTachoDataCode.cellpremiumBasicCnt = 0;
                        stTachoDataCode.cellpremiumAfterCnt = 0;
                        stTachoDataCode.cellsalesCnt = 0;
                        stTachoDataCode.cellsalesTime = 0;
                        stTachoDataCode.cellcarEmptyTime = 0;
                        stTachoDataCode.cellkeyUseCnt = 0;
                        stTachoDataCode.cellemerBreakCnt = 0;
                        stTachoDataCode.celloverrunTime = new DateTime(2010, 1, 1, 0, 0, 0);
                        // '10. 7.19 추가

                        stTachoDataCode.MKdoor = false;
                        stTachoDataCode.TotalDriveDistanceSaved = 0;
                        stTachoDataCode.TotalDriveDistance = 0;
                        stTachoDataCode.speed = 0;
                        stTachoDataCode.distance = 0;
                        stTachoDataCode.sales = false;
                        stTachoDataCode.engine = false;

                        stTachoDataCode.salesCnt = 0;
                        stTachoDataCode.carEmptyTime = 0;
                        stTachoDataCode.keyUseCnt = 0;
                        stTachoDataCode.salesTime = 0;
                    }
                    #endregion Initialized struct value


                    if (data_start == 0)
                    {
                        temp[0] = br.ReadByte();
                        if (temp[0] != 0x00)
                        {
                            SDcard_chk = true;
                        }
                        else
                        {


                            fs.Position = 0x00;
                        }
                        data_start = 1;
                    }

                    // 할증 기본 거리 (0x00)
                    temp = br.ReadBytes(2);
                    stTachoRamData.PremiumBasicDistance = BcdToDecimalByLsb(temp, 2);
                    // 할증 이후 거리 (0x02)
                    temp = br.ReadBytes(2);
                    stTachoRamData.PremiumAfterDistance = BcdToDecimalByLsb(temp, 2);
                    // 주행 기본 거리 (0x04)
                    temp = br.ReadBytes(2);
                    stTachoRamData.DriveBasicDistance = BcdToDecimalByLsb(temp, 2);
                    // 주행 이후 거리 (0x06)
                    temp = br.ReadBytes(2);
                    stTachoRamData.DriveAfterDistance = BcdToDecimalByLsb(temp, 2);
                    // 할증 기본 요금 (0x08)
                    temp = br.ReadBytes(2);
                    stTachoRamData.PremiumBasicMoney = BcdToDecimalByLsb(temp, 2);
                    // 할증 이후 요금 (0x0A)
                    temp = br.ReadBytes(2);
                    stTachoRamData.PremiumAfterMoney = BcdToDecimalByLsb(temp, 2);
                    // 주행 기본 요금 (0x0C)
                    temp = br.ReadBytes(2);
                    stTachoRamData.DriveBasicMoney = BcdToDecimalByLsb(temp, 2);
                    // 주행 이후 요금 (0x0E)
                    temp = br.ReadBytes(2);
                    stTachoRamData.DriveAfterMoney = BcdToDecimalByLsb(temp, 2);
                    // 호출 요금 (0x10)
                    temp = br.ReadBytes(2);
                    stTachoRamData.CallMoney = BcdToDecimalByLsb(temp, 2);
                    // 화물 요금 (0x12)
                    temp = br.ReadBytes(2);
                    stTachoRamData.FreightMoney = BcdToDecimalByLsb(temp, 2);
                    // 인할 요금 (0x14)
                    temp = br.ReadBytes(2);
                    stTachoRamData.DiscountMoney = BcdToDecimalByLsb(temp, 2);

                    stTachoRamData.PremiumBasicMoney = (int)(stTachoRamData.DriveBasicMoney * 0.2) + stTachoRamData.DriveBasicMoney;

                    stTachoRamData.PremiumAfterMoney = (int)(stTachoRamData.DriveAfterMoney * 0.2) + stTachoRamData.DriveAfterMoney;


                    byte[] jump = br.ReadBytes(3);
                    temp = br.ReadBytes(2);

                    if (temp[0] != 0x00 && temp[1] != 0x00)
                    {
                        NewFA = true;
                        //    MessageBox.Show("test");
                    }

                    jump = br.ReadBytes(4);

                    // byte[] jump = br.ReadBytes(9);
                    // Off Size  (0x1f)
                    temp = br.ReadBytes(2);

                    stTachoRamData.OffSize = temp[1];
                    stTachoRamData.OffSize = (stTachoRamData.OffSize << 8) + temp[0];

                    temp = br.ReadBytes(2);	// 기사번호  (0x21)
                    stTachoRamData.DriverNumber = ((BcdToDecimal(temp[1]) * 100) + BcdToDecimal(temp[0]));
                    jump = br.ReadBytes(12); // 건너뛰기

                    temp[0] = br.ReadByte();  // 1  버젼 확인 하기
                    stTachoRamData.TVersion = temp[0];

                    if (stTachoRamData.TVersion == 1)
                    {
                        e7cnt = 13;
                    }
                    else
                    {
                        e7cnt = 10;
                    }

                    temp[1] = br.ReadByte();  // f9 확인 하기
                    f9cnt = temp[1];

                    if (f9cnt == 1)
                    {
                        e7cnt = 13;
                    }

                    jump = br.ReadBytes(99); // 건너뛰기

                    // 건너뛰기
                    //	byte[] jump = br.ReadBytes(126);

                    // 총 주행 거리 (0x94)
                    temp = br.ReadBytes(4);
                    stTachoRamData.TotalDriveDistance = BcdToDecimalByLsb(temp, 4);

                    // 건너뛰기
                    jump = br.ReadBytes(21);

                    // 총 영업 거리 (0xAC)
                    temp = br.ReadBytes(4);
                    stTachoRamData.TotalTradeDistance = BcdToDecimalByLsb(temp, 4);

                    // 건너뛰기
                    jump = br.ReadBytes(10);

                    // 당일 수입금 (0xBB)
                    temp = br.ReadBytes(3);


                    /*	if (temp[0] == 0x00 && temp[1] == 0x00 && temp[2] == 0x00)  // 11.06.21
                        {

                            while (fs.Position < fs.Length)
                            {

                                byte bRead = br.ReadByte();
                                if (bRead == 0xFD || bRead == 0xF7 || bRead == 0xFF)
                                {

                                    meterzero = true;
                                    newFdPosition = fs.Position;
                                    break;

                                }

                            }
                            continue;
                        }*/

                    stTachoRamData.TodayIncomeMoney = BcdToDecimalByLsb(temp, 3);
                    // 1펄스당 거리 (0xBE)
                    temp = br.ReadBytes(2);
                    stTachoRamData.DistanceBy1Pulse = BcdToDecimalByLsb(temp, 2);

                    // 건너뛰기
                    jump = br.ReadBytes(21);

                    // 당일 총 주행 거리 (0xD5)
                    temp = br.ReadBytes(4);
                    stTachoRamData.TodayTotalDriveDistance = BcdToDecimalByLsb(temp, 4);

                    // 건너뛰기
                    jump = br.ReadBytes(59);

                    // 당일 총 영업 거리 (0x114)
                    temp = br.ReadBytes(4);
                    stTachoRamData.TodayTotalTradeDistance = BcdToDecimalByLsb(temp, 4);
                    // 입고 시각 (0x118)
                    temp[0] = br.ReadByte();
                    ss = BcdToDecimalByLsb(temp, 1);
                    temp[0] = br.ReadByte();
                    mm = BcdToDecimalByLsb(temp, 1);
                    temp[0] = br.ReadByte();


                    if (temp[0] > 0x40)      // 13.5.2 12시간제와 24시간제 체크하기
                    {
                        if ((temp[0] & 0x40) == 0x40) // PM
                        {
                            temp[0] &= 0x3F;
                            hh = 12 + BcdToDecimalByLsb(temp, 1);
                            if (hh == 24) hh = 12;
                        }
                        else   // AM
                        {
                            temp[0] &= 0x3F;
                            hh = BcdToDecimalByLsb(temp, 1);
                            if (hh == 12) hh = 0;
                        }

                    }
                    else
                    {
                        hh = BcdToDecimal(temp[0]);

                    }
                    temp[0] = br.ReadByte();
                    dd = BcdToDecimalByLsb(temp, 1);
                    temp[0] = br.ReadByte();
                    mon = BcdToDecimalByLsb(temp, 1);
                    temp[0] = br.ReadByte();
                    yy = BcdToDecimalByLsb(temp, 1);

                    // Data가 채워지지 않은 경우에 대비
                    if ((mon == 0) || (dd == 0) || (mon > 12) || (dd > 31) || (hh > 24) || (mm > 60) || (ss > 60))
                    {
                        return;
                    }

                    stTachoRamData.InWarehouseTime = new DateTime(2000 + yy, mon, dd, hh, mm, ss);

                    // 건너뛰기
                    jump = br.ReadBytes(26);

                    // 실 수입금 (0x138)
                    temp = br.ReadBytes(3);
                    stTachoRamData.RealIncomeMoney = BcdToDecimalByLsb(temp, 3);

                    // 건너뛰기
                    jump = br.ReadBytes(2);

                    // 기사 번호 / 연료량 (0x13D)
                    temp = br.ReadBytes(3);
                    //	stTachoRamData.DriverNumber = ((int)(temp[2] >> 4) * 100) + ((int)(temp[2] & 0x0F) * 10) + ((int)(temp[1] >> 4));
                    stTachoRamData.Fuel = ((double)(temp[1] & 0x0F)) + ((double)(temp[0] >> 4) / 10) + ((double)(temp[0] & 0x0F) / 100);

                    // 건너뛰기
                    jump = br.ReadBytes(10);

                    // 차량 번호 (0x14A)
                    temp = br.ReadBytes(6);
                    stTachoRamData.CarNumber = String.Format("{0:X2}-{1:X2}{2:X2}", temp[5], temp[4], temp[3]);

                    switch (temp[1])
                    {
                        case 0x01: CarArea = "서울"; break;
                        case 0x02: CarArea = "인천"; break;
                        case 0x03: CarArea = "대전"; break;
                        case 0x04: CarArea = "광주"; break;
                        case 0x05: CarArea = "대구"; break;
                        case 0x06: CarArea = "울산"; break;
                        case 0x07: CarArea = "부산"; break;
                        case 0x08: CarArea = "경기"; break;
                        case 0x09: CarArea = "강원"; break;
                        case 0x10: CarArea = "충북"; break;
                        case 0x11: CarArea = "충남"; break;
                        case 0x12: CarArea = "전북"; break;
                        case 0x13: CarArea = "전남"; break;
                        case 0x14: CarArea = "경북"; break;
                        case 0x15: CarArea = "경남"; break;

                        case 0xa: CarArea = "충북"; break;
                        case 0xb: CarArea = "충남"; break;
                        case 0xc: CarArea = "전북"; break;
                        case 0xd: CarArea = "전남"; break;
                        case 0xe: CarArea = "경북"; break;
                        case 0xf: CarArea = "경남"; break;

                        case 0x16: CarArea = "제주"; break;
                        default: CarArea = "없음"; break;
                    }
                    int nn = temp[0];

                    string carRegistNumSign = " 가나다라마바사아자차카타파하거너더러머버서어저처커터퍼허고노도로모보소오조초코토포호구누두루무부수우주추쿠투푸후그느드르므브스으즈츠크트프흐기니디리미비시이지치키티피히";

                    if (nn > carRegistNumSign.Length)
                    {
                        nn = 0;
                    }
                    else
                    {
                        CarSign = carRegistNumSign[nn];
                    }
                    CarSign = carRegistNumSign[nn];

                   // formData.CarArea = CarArea;
                   // formData.CarSign = CarSign;

                    // 건너뛰기
                    jump = br.ReadBytes(32);

                    // 초 데이터 포인터 (0x170)
                    byte chsum = 0x00;
                    stTachoRamData._start = br.ReadUInt16();
                    chsum += (byte)(stTachoRamData._start >> 8);
                    chsum += (byte)(stTachoRamData._start & 0x00FF);
                    stTachoRamData._size = br.ReadUInt16();
                    chsum += (byte)(stTachoRamData._size >> 8);
                    chsum += (byte)(stTachoRamData._size & 0x00FF);
                    stTachoRamData._pointer = br.ReadUInt16();
                    chsum += (byte)(stTachoRamData._pointer >> 8);
                    chsum += (byte)(stTachoRamData._pointer & 0x00FF);
                    stTachoRamData._overflag = br.ReadByte();
                    chsum += stTachoRamData._overflag;
                    byte readchsum = br.ReadByte();

                    if (readchsum != chsum)
                    {
                        //	this.Text = "어?";
                    }

                    // 건너뛰기
                    jump = br.ReadBytes(2);

                    // 버전 번호 (0x17A)
                    temp = br.ReadBytes(3);
                    stTachoRamData.VersionNumber = BcdToDecimalByLsb(temp, 1);
                    // 타고 저장 시간(초) (0x17D)
                    temp = br.ReadBytes(3);
                    stTachoRamData.TachoSavedTime = (int)temp[0] + 1;

                    // Fill Struct from Emergency data
                    for (int k = 0; k < 16; k++)
                        bArrEmergencyData[k] = br.ReadBytes(16);

                    //if (stTachoRamData._size > 0)                     // 전주 수정택시  
                    //	fs.Position += stTachoRamData._size;

                    fs.Position -= 2;

                    temp = br.ReadBytes(2);
                    if (temp[0] == 0x7e && temp[1] == 0x1a)
                    {

                        //	if (temp[1] != 0 && temp[0] != 0)
                        //{

                        stTachoRamData.Nandsize = temp[1];
                        stTachoRamData.Nandsize = (stTachoRamData.Nandsize << 8) + temp[0];
                        //	}
                    }
                    else
                    {
                        //temp = br.ReadBytes(2);
                        //stTachoRamData.Nandsize = temp[1];
                        //stTachoRamData.Nandsize = (stTachoRamData.Nandsize << 8) + temp[0];
                        stTachoRamData.Nandsize = 0x1a7e;
                    }


                    // 여기서부터가 실질적인 Tacho data 분석 루틴
                    oldFdPosition = newFdPosition = fs.Position;
                    byte[] Tacho_time_byte = new byte[6];
                    int[] Tacho_time_int = new int[3];

                    string[] Tacho_time = new string[3];
                    bool F7Mark = false;



                    DateTime[] DayDBName = new DateTime[40];
                    string[] DayDbName_str = new string[40];
                    string YearDir = "20" + Tacho_time[0];
                    string tachoday = "타코";
                    long test = 0;
                    int Day_cnt = 0;
                    byte[] TimeTemp = new byte[6];
                    bool timerror = false;


                    while (fs.Position < fs.Length)   // 
                    {

                        byte bRead = br.ReadByte();
                        newFdPosition++;
                        ////////////////////////////////////////////////////////////////////////////////
                        if (f9cnt != 3)
                        {

                            if (bRead == 0xFD)
                            {


                                fs.Position -= 11;
                                Tacho_time_byte[0] = br.ReadByte();
                                Tacho_time_byte[1] = br.ReadByte();
                                Tacho_time_byte[2] = br.ReadByte();
                                Tacho_time_byte[3] = br.ReadByte();
                                Tacho_time_byte[4] = br.ReadByte();
                                Tacho_time_byte[5] = br.ReadByte();

                                yy = BcdToDecimal(Tacho_time_byte[0]);
                                mon = BcdToDecimal(Tacho_time_byte[1]);
                                dd = BcdToDecimal(Tacho_time_byte[2]);



                                if ((Tacho_time_byte[3] & 0x40) == 0x40) // PM
                                {
                                    Tacho_time_byte[3] &= 0x3F;
                                    hh = 12 + BcdToDecimal(Tacho_time_byte[3]);
                                    if (hh == 24) hh = 12;
                                }
                                else   // AM
                                {
                                    Tacho_time_byte[3] &= 0x3F;
                                    hh = BcdToDecimal(Tacho_time_byte[3]);
                                    if (hh == 12) hh = 0;
                                }

                                mm = BcdToDecimal(Tacho_time_byte[4]);
                                ss = BcdToDecimal(Tacho_time_byte[5]);



                                stTachoRamData.OutWarehouseTime = new DateTime(2000 + yy, mon, dd, hh, mm, ss);

                                /*  Tacho_time_int[0] = BcdToDecimal(Tacho_time_byte[0]);
                                  Tacho_time_int[1] = BcdToDecimal(Tacho_time_byte[1]);
                                  Tacho_time_int[2] = BcdToDecimal(Tacho_time_byte[2]);

                                  Tacho_time[0] = Tacho_time_int[0].ToString();
                                  Tacho_time[1] = Tacho_time_int[1].ToString();
                                  Tacho_time[2] = Tacho_time_int[2].ToString();*/


                                break;

                            }
                        }
                        ////////////////////////////////////////////////////////////////////////////////


                    }

                    double TotalDistance = (double)stTachoRamData.TodayTotalDriveDistance / 1000;
                    double TotalSalesDist = (double)stTachoRamData.TodayTotalTradeDistance / 1000;
                    //  Thread.Sleep(3000);


                         listviewID = listView1.Items.Count;


                         if (listviewID == 1000)
                         {
                           //  Excel_print();
                             listView1.Items.Clear();
                             listviewID = 0;

                         }

                      /*   if (listviewID != 0)
                         {
                             listView1.Items[listviewID - 1].Remove();
                            listviewID = listView1.Items.Count;
                         }*/

                         listviewID++;
                         ListViewItem a = new ListViewItem(listviewID.ToString());              // ID
                        a.SubItems.Add(stTachoRamData.CarNumber);
                        a.SubItems.Add(stTachoRamData.OutWarehouseTime.ToString("yyyy-MM-dd tt HH:mm"));         // 출고시간
                        a.SubItems.Add(stTachoRamData.InWarehouseTime.ToString("yyyy-MM-dd tt HH:mm"));          // 입고시간


                        string mmm = string.Format("{0:C}", stTachoRamData.TodayIncomeMoney);
                        recvTotal_Money += stTachoRamData.TodayIncomeMoney;
                        a.SubItems.Add(mmm);       // 미터금

                        string dist = string.Format("{0:N} Km", TotalDistance);                  // 주행거리
                        recvTotal_Dist += TotalDistance;
                        a.SubItems.Add(dist);

                        dist = string.Format("{0:N} Km", TotalSalesDist);                  // 영업거리
                        reevTotal_SalesDist += TotalSalesDist;
                        a.SubItems.Add(dist);


                       

                      //  a.SubItems.Add(mmm);          // 실입금 
                        //   recvTotal_Money += MeterMoney[i];


                        listView1.Items.Add(a);


                     

                          if (listView1.Items.Count != 0)
                            {
                               listView1.Items[listView1.Items.Count - 1].EnsureVisible();   // 스크롤 이동
                            }

                           

                    
                    string strtemp = "";


                        strtemp = "A" + listviewID.ToString();         // id
               //    strtemp = "A" + " ";

                    strtemp += "B" + stTachoRamData.CarNumber;         // CarNumber

                    strtemp += "C" + stTachoRamData.OutWarehouseTime.ToString("yyyy-MM-dd tt HH:mm");         // OutTime


                    strtemp += "D" + stTachoRamData.InWarehouseTime.ToString("yyyy-MM-dd tt HH:mm");         // InTime


                    strtemp += "E" + string.Format("{0:N} Km", TotalDistance);         // 주행거리

                    strtemp += "F" + string.Format("{0:N} Km", TotalSalesDist);         // 영업거리


                    strtemp += "G" + string.Format("{0:C}", stTachoRamData.TodayIncomeMoney);         //미터 수입
                    strtemp += "H";


                    svr.TachoStr = strtemp;
                    //int IDtemp = 0;
                   
                 
                    
                /*    if (IDList.Count > 1)
                    {

                        for (int q = 0; q < IDList.Count; q++)
                        {

                            if (q < IDList.Count)
                            {
                                SendMsg(IDList[q], strtemp);
                            }


                        }

                    }
                    else
                    {
                        if (IDList.Count == 1)
                        {

                            SendMsg(IDList[0], strtemp);

                        }
                    }
                    */

                }
                catch (System.Exception e)
                {
                 //   MessageBox.Show(e.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    string path = Application.StartupPath + "\\ErrorLog.jie";
                    using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                    {
                        sw.WriteLine("Tacho_Run :[" + DateTime.Now.ToString() + "] " + e.StackTrace);
                    }
                }

                fs.Close();
                br.Close();
                bw.Close();
            }

        }
        public void Excel_print()
        {

            string filePath = "c:\\Tacho_Report.xlsx";
            Excel.ApplicationClass excel = new Excel.ApplicationClass();

            int colIndex = 0;
            int rowIndex = 3;
            Excel.Application excelApp = null;
            excel.UserName = formname + ".xlsx";
         

            object missingType = Type.Missing;
            object fileName = formname + ".xlsx";
           
          
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            //  excel.Application.Workbooks.Add(true);
            Excel.Range oRng = null;
            Excel.Workbook excelBook = excel.Workbooks.Add(missingType);
          
      

         //   excelBook = excelApp.Workbooks.Open((string)fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
          //    Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing);



        //    excelBook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
      //  Excel.XlSaveAsAccessMode.xlShared, Excel.XlSaveConflictResolution.xlLocalSessionChanges,
      //  Type.Missing, Type.Missing, Type.Missing, Type.Missing);

       
            Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.Add(missingType, missingType, missingType, missingType);
            excelWorksheet.Name = formname;
            ((Excel.Range)excelWorksheet.get_Range("A:A", Missing.Value)).ColumnWidth = 8.33;
            ((Excel.Range)excelWorksheet.get_Range("B:B", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("C:C", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("D:D", Missing.Value)).ColumnWidth = 25.33;
            ((Excel.Range)excelWorksheet.get_Range("E:E", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("F:F", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("G:G", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("H:H", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("I:I", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("J:J", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("K:K", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("L:L", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("M:M", Missing.Value)).ColumnWidth = 14.33;
            ((Excel.Range)excelWorksheet.get_Range("N:N", Missing.Value)).ColumnWidth = 18.33;
            ((Excel.Range)excelWorksheet.get_Range("O:O", Missing.Value)).ColumnWidth = 18.33;
            ((Excel.Range)excelWorksheet.get_Range("P:P", Missing.Value)).ColumnWidth = 14.33;
        

            // BackColor 설정
             ((Excel.Range)excelWorksheet.get_Range("A3", "A3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("B3", "B3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("C3", "C3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("D3", "D3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("E3", "E3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("F3", "F3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("G3", "G3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("H3", "H3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("I3", "I3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("J3", "J3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("K3", "K3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("L3", "L3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("M3", "M3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("N3", "N3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("O3", "O3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);
             ((Excel.Range)excelWorksheet.get_Range("P3", "P3")).Interior.Color = ColorTranslator.ToOle(Color.LightGray);



             ((Excel.Range)excelWorksheet.get_Range("A:A, B:B, C:C, D:D, E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,N:N,O:O,P:P", Missing.Value)).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            oRng = excel.get_Range("A1", "N1"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.Value2 =  " SHIFT REPORT ";  //문구 삽입
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Size = 16;
            oRng.Font.Bold = true;  //볼드
            //  oRng.Font.Color = 0xfffffff;    //폰트 컬러




            oRng = excel.get_Range("A3", "A3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //  oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;

            ///////////////////////////////////////////////////////

            oRng = excel.get_Range("B3", "B3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            ///////////////////////////////////////////////////////

            oRng = excel.get_Range("C3", "C3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            ///////////////////////////////////////////////////////

            oRng = excel.get_Range("D3", "D3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            ///////////////////////////////////////////////////////

            oRng = excel.get_Range("E3", "E3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            ///////////////////////////////////////////////////////

            oRng = excel.get_Range("F3", "F3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            ///////////////////////////////////////////////////////

            oRng = excel.get_Range("G3", "G3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            ///////////////////////////////////////////////////////

            oRng = excel.get_Range("H3", "H3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            ///////////////////////////////////////////////////////

            oRng = excel.get_Range("I3", "I3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;



            oRng = excel.get_Range("J3", "J3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("K3", "K3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("L3", "L3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("M3", "M3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("N3", "N3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("O3", "O3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;


            oRng = excel.get_Range("P3", "P3"); //해당 범위의 셀 획득
            oRng.MergeCells = true; //머지
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;   //가운데 정렬
            oRng.Font.Bold = true;  //볼드
            //    oRng.Font.Color = -16776961;    //폰트 컬러
            //테두리
            oRng.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            oRng.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;



            for (int i = 0; i < this.listView1.Columns.Count; i++)
            {
                if (iButtonMode == false)
                {
                    if (THAILAND_Set == true)
                    {
                        if (i == 3 || i == 8 || i == 9 || i == 14 || i == 15 || i == 12 || i == 13 || i == 2)
                        {
                            continue;
                        }
                    }
                    else
                    {
                        if (i == 3 || i == 8 || i == 9 || i == 14 || i == 15 )
                        {
                            continue;
                        }
                    }
                }
                colIndex++;
                excel.Cells[3, colIndex] = this.listView1.Columns[i].Text;

            }

            for (int i = 0; i < this.listView1.Items.Count; i++)
            {
                rowIndex++;
                colIndex = 0;
                for (int j = 0; j < this.listView1.Items[i].SubItems.Count; j++)
                {
                    if (iButtonMode == false)
                    {
                        if (THAILAND_Set == true)
                        {

                            if (j == 3 || j == 8 || j == 9 || j == 14 || j == 15 || j == 12 || j == 13 || j == 2)
                            {
                                continue;
                            }
                        }
                        else
                        {
                            /*
                            if (j == 3 || j == 8 || j == 9 || j == 14 || j == 15 )
                            {
                                continue;
                            }
                             */
                            if (j == 3 || j == 8 || j == 9 || j == 14 || j == 15 || j == 12 || j == 13 || j == 2)
                            {
                                continue;
                            }
                        }
                    }
                    colIndex++;

                    excel.Cells[rowIndex, colIndex] = this.listView1.Items[i].SubItems[j].Text;
                    //     excel.Cells.AutoOutline();

                }
            }
            string Acell = "A";
            string Jcell = "P";
            Acell += (this.listView1.Items.Count+3).ToString();
            Jcell += (this.listView1.Items.Count+3).ToString();
            ((Excel.Range)excelWorksheet.get_Range(Acell, Jcell)).Interior.Color = ColorTranslator.ToOle(Color.LightGray);

            System.IO.Directory.CreateDirectory("c:\\Tacho2\\EXCEL");
            string savefile = "c:\\Tacho2\\EXCEL\\" + formname + ".xlsx";

            savefile = "c:\\Tacho2\\EXCEL\\Report_20" + formname[0] + formname[1] + "-" + formname[2] + formname[3] + "-" + formname[4] + formname[5];
           
           // excelBook.SaveAs(savefile, Excel.XlFileFormat.xlExcel7);
         
            if (System.IO.File.Exists(savefile+".xlsx"))  // information 같은 파일의이름이 존재 함
            {
         
            //    MessageBox.Show("같은 파일 존재");
             //   return;
                int num = 1;
                bool check = true;
                savefile += "_" + num.ToString();
          

                do{
                 
                            if (!System.IO.File.Exists(savefile + ".xlsx"))  // information 같은 파일의이름이 존재 함
                            {
                                check = false;
                            }
                            num++;
                            if (check == true)
                            {
                                savefile = savefile.Remove(savefile.Length - 1);
                                savefile += num.ToString();
                            }
                  

                }while (check);
              
               
            }
       
            excel.Visible = true;
            excelBook.SaveAs(savefile); 

        /*    if (excelApp != null) // 작업후 프로세스가 남는 경우를 방지하기 위해서...
            {
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                pProcess[0].Kill();
            }*/

         


         //   Marshal.ReleaseComObject(excelWorksheet);
         //   Marshal.ReleaseComObject(excelBook);
         //   Marshal.ReleaseComObject(excelApp);
           // excel.Visible = true;
            
            /*
            string filePath = "c:\\test.xlsx";
            Excel.ApplicationClass excel = new Excel.ApplicationClass();
            int colIndex = 0;
            int rowIndex = 1;
            excel.Application.Workbooks.Add(true);

            for (int i = 0; i < listView1.Columns.Count; i++)
            {
                colIndex++;
                excel.Cells[1, colIndex] = listView1.Columns[i].Text;
            }
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                rowIndex++;
                colIndex = 0;
                for (int j = 0; j < listView1.Items[i].SubItems.Count; j++)
                {
                    colIndex++;
                    excel.Cells[rowIndex, colIndex] = listView1.Items[i].SubItems[j].Text;
                    //     excel.Cells.AutoOutline();

                }
            }
            excel.Visible = true;*/


            //excel.Save(filePath);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Excel_print();
           /* byte[] Data = new byte[2];

            Data[0] = 0x55;
            Data[1] = 0x15;

            serialPort1.DiscardOutBuffer();
            serialPort1.DiscardInBuffer();
            serialPort1.Write(Data, 0, Data.Length);*/
            
        }

        private void button8_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            listviewID = 0;
        }
        public void TransData1()
        {
             try
            {

               


                string TmfPath = TACHO2_path + "TMF\\";  // tmf 경로 

                string TransDataPath = System.IO.Path.Combine(TACHO2_path + "TMF", "TransData");  // 판독기 파일 경로 


                string[] OrignalFiles = Directory.GetFiles(TransDataPath, "*.tmf");  // string file 목록  원본 파일 

                DirectoryInfo dir = new DirectoryInfo(TransDataPath);
                FileInfo[] filesName = dir.GetFiles();   // 파일 가져오기

                string BackupPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "BACKUP");

                for (int i = 0; i < filesName.Length; i++)
                {
                    string NowhhmmssTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                        DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);  // 현재 시간 

                    string TransData = filesName[i].ToString();
                    TransData = TransData.Replace(".TMF", "");

                    System.IO.File.Copy(OrignalFiles[i], TmfPath + TransData + "_" + NowhhmmssTime + ".TMF");

                
                    // Create the subfolder
                    System.IO.Directory.CreateDirectory(BackupPath);

                    System.IO.Directory.Move(OrignalFiles[i], BackupPath + "\\" + TransData + "_" + NowhhmmssTime + ".TMF");
                }

             ///////////////// auto 폴더 backup 폴더 이동/////////////////////////////////////////////////////////////////////////////
                string AutoDataPath = System.IO.Path.Combine(TACHO2_path + "TMF", "Auto");  // 판독기 파일 경로 
                System.IO.Directory.CreateDirectory(AutoDataPath);
                System.IO.Directory.CreateDirectory(BackupPath);
                string NowhhmmssTime1 = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}", (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day,
                                                       DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);  // 현재 시간 

                string BackupPath1 = System.IO.Path.Combine(TACHO2_path + "\\TMF\\BACKUP", "Auto_" + NowhhmmssTime1);
                System.IO.Directory.Move(AutoDataPath, BackupPath1);

               
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            

        }

        public void TransData2()
        {
            try
            {




                string TmfPath = TACHO2_path + "TMF_1\\";  // tmf 경로 

                string TransDataPath = System.IO.Path.Combine(TACHO2_path + "TMF_1", "TransData");  // 판독기 파일 경로 
                System.IO.Directory.CreateDirectory(TransDataPath);

                string[] OrignalFiles = Directory.GetFiles(TransDataPath, "*.tmf");  // string file 목록  원본 파일 

                DirectoryInfo dir = new DirectoryInfo(TransDataPath);
                FileInfo[] filesName = dir.GetFiles();   // 파일 가져오기

                string BackupPath = System.IO.Path.Combine(TACHO2_path + "\\TMF_1", "BACKUP");

                for (int i = 0; i < filesName.Length; i++)
                {
                    string NowhhmmssTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                        DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);  // 현재 시간 

                    string TransData = filesName[i].ToString();
                    TransData = TransData.Replace(".TMF", "");

                    System.IO.File.Copy(OrignalFiles[i], TmfPath + TransData + "_" + NowhhmmssTime + ".TMF");

                   
                    // Create the subfolder
                    System.IO.Directory.CreateDirectory(BackupPath);

                    System.IO.Directory.Move(OrignalFiles[i], BackupPath + "\\" + TransData + "_" + NowhhmmssTime + ".TMF");
                }

                ///////////////// auto 폴더 backup 폴더 이동/////////////////////////////////////////////////////////////////////////////
                string AutoDataPath = System.IO.Path.Combine(TACHO2_path + "TMF_1", "Auto");  // 판독기 파일 경로 
                System.IO.Directory.CreateDirectory(AutoDataPath);
                System.IO.Directory.CreateDirectory(BackupPath);
                string NowhhmmssTime1 = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}", (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day,
                                                      DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);  // 현재 시간 

                string BackupPath1 = System.IO.Path.Combine(TACHO2_path + "\\TMF_1\\BACKUP", "Auto_" + NowhhmmssTime1);
                System.IO.Directory.Move(AutoDataPath, BackupPath1);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }


        public void TransData3()
        {
            try
            {




                string TmfPath = TACHO2_path + "TMF_2\\";  // tmf 경로 

                string TransDataPath = System.IO.Path.Combine(TACHO2_path + "TMF_2", "TransData");  // 판독기 파일 경로 
                System.IO.Directory.CreateDirectory(TransDataPath);

                string[] OrignalFiles = Directory.GetFiles(TransDataPath, "*.tmf");  // string file 목록  원본 파일 

                DirectoryInfo dir = new DirectoryInfo(TransDataPath);
                FileInfo[] filesName = dir.GetFiles();   // 파일 가져오기


                string BackupPath = System.IO.Path.Combine(TACHO2_path + "\\TMF_2", "BACKUP");

                for (int i = 0; i < filesName.Length; i++)
                {
                    string NowhhmmssTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                        DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);  // 현재 시간 

                    string TransData = filesName[i].ToString();
                    TransData = TransData.Replace(".TMF", "");

                    System.IO.File.Copy(OrignalFiles[i], TmfPath + TransData + "_" + NowhhmmssTime + ".TMF");
     
                    // Create the subfolder
                    System.IO.Directory.CreateDirectory(BackupPath);

                    System.IO.Directory.Move(OrignalFiles[i], BackupPath + "\\" + TransData + "_" + NowhhmmssTime + ".TMF");
                }

                ///////////////// auto 폴더 backup 폴더 이동/////////////////////////////////////////////////////////////////////////////
                string AutoDataPath = System.IO.Path.Combine(TACHO2_path + "TMF_2", "Auto");  // 판독기 파일 경로 
                System.IO.Directory.CreateDirectory(AutoDataPath);

                
                // Create the subfolder
                System.IO.Directory.CreateDirectory(BackupPath);

                string NowhhmmssTime1 = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}", (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day,
                                                    DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);  // 현재 시간 

                string BackupPath1 = System.IO.Path.Combine(TACHO2_path + "\\TMF_2\\BACKUP", "Auto_" + NowhhmmssTime1);
                System.IO.Directory.Move(AutoDataPath, BackupPath1);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }


        public void TransData4()
        {
            try
            {




                string TmfPath = TACHO2_path + "TMF_3\\";  // tmf 경로 

                string TransDataPath = System.IO.Path.Combine(TACHO2_path + "TMF_3", "TransData");  // 판독기 파일 경로 
                System.IO.Directory.CreateDirectory(TransDataPath);

                string[] OrignalFiles = Directory.GetFiles(TransDataPath, "*.tmf");  // string file 목록  원본 파일 

                DirectoryInfo dir = new DirectoryInfo(TransDataPath);
                FileInfo[] filesName = dir.GetFiles();   // 파일 가져오기

                string BackupPath = System.IO.Path.Combine(TACHO2_path + "\\TMF_3", "BACKUP");

                for (int i = 0; i < filesName.Length; i++)
                {
                    string NowhhmmssTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                        DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);  // 현재 시간 

                    string TransData = filesName[i].ToString();
                    TransData = TransData.Replace(".TMF", "");

                    System.IO.File.Copy(OrignalFiles[i], TmfPath + TransData + "_" + NowhhmmssTime + ".TMF");

                
                    // Create the subfolder
                    System.IO.Directory.CreateDirectory(BackupPath);

                    System.IO.Directory.Move(OrignalFiles[i], BackupPath + "\\" + TransData + "_" + NowhhmmssTime + ".TMF");
                }

                ///////////////// auto 폴더 backup 폴더 이동/////////////////////////////////////////////////////////////////////////////
                string AutoDataPath = System.IO.Path.Combine(TACHO2_path + "TMF_3", "Auto");  // 판독기 파일 경로 
                System.IO.Directory.CreateDirectory(AutoDataPath);
                System.IO.Directory.CreateDirectory(BackupPath);
                string NowhhmmssTime1 = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}", (DateTime.Now.Year-2000), DateTime.Now.Month, DateTime.Now.Day,
                                                      DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);  // 현재 시간 

                string BackupPath1 = System.IO.Path.Combine(TACHO2_path + "\\TMF_3\\BACKUP", "Auto_" + NowhhmmssTime1);
                System.IO.Directory.Move(AutoDataPath, BackupPath1);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.Enabled = false;

            TransData1();
            TransData2();
            TransData3();
            TransData4();

                listView1.Items.Clear();
                listviewID = 0;

                this.Enabled = true;
                   

        }
        public void TestSend()
        {
            testid++;
            string strtemp = "";
            strtemp = "A" + testid.ToString();         // id

            strtemp += "B" + "11-1234";         // CarNumber

            DateTime Outtime = new DateTime(2014, 1, 1, 1, 1, 1);
            strtemp += "C" + Outtime.ToString("yyyy-MM-dd tt HH:mm");         // OutTime



            DateTime Intime = new DateTime(2014, 1, 1, 2, 2, 2);
            strtemp += "D" + Intime.ToString("yyyy-MM-dd tt HH:mm");         // InTime

            double TotalDistance = 100;
            strtemp += "E" + string.Format("{0:N} Km", TotalDistance);         // 주행거리

            double TotalSalesDist = 1000;
            strtemp += "F" + string.Format("{0:N} Km", TotalSalesDist);         // 영업거리

            int TodayIncomeMoney = 220000;
            strtemp += "G" + string.Format("{0:C}", TodayIncomeMoney);         //미터 수입

            strtemp += "H";



            if (IDList.Count > 1)
            {
                for (int q = 0; q < IDList.Count; q++)
                {


                    SendMsg(IDList[q], strtemp);
                }

            }
            else
            {
                if (IDList.Count == 1)
                {
                    SendMsg(IDList[0], strtemp);
                }
            }
            
        }

        private void spTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            TimeCount++;
          
            TimeCnt++;

            TimeLabel.Text = string.Format("{0:D2}:{1:D2}:{2:D2}", DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);

            CurTime = string.Format("{0:D2}:{1:D2}:{2:D2}", DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);



          /*  if (TimeCount == 10)
            {
                TimeCount = 0;
                if (CashierMode == false)
                {
                    if (treeView1.Nodes.Count != 0)
                    {
                        string path = "C:\\tacho2\\";
                        TimeCount = 0;
                        Treeview_Refresh();
                        mdbfilename = path + "TACHO" + "\\" + treeView1.Nodes[0].Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes[0].Text + ".mdb";

                        DB_ReadData(0, 1);
                    }
                }
            }*/


            if (pictureBox17.Visible == true)
            {
                LoadingTimeOut++;

                if (LoadingTimeOut == 20)
                {
                    LoadingTimeOut = 0;
                    pictureBox17.Visible = false;
                    Client_label.Text = "";

                }
            }
            if (TimeCnt == 30)
            {
                TimeCnt = 0;
                
                if (NetworkInterface.GetIsNetworkAvailable())  // 백업된 데이터를 서버로 전송하자.~
                {
                    string path = TACHO2_path + "TMF\\BACKUP";

                  

                   

                    DirectoryInfo di = new DirectoryInfo(path);

                    if (di.Exists == false)
                    {

                        di.Create();

                    }
                    DirectoryInfo dir = new DirectoryInfo(path);

                    FileInfo[] files = dir.GetFiles();

                    for (int i = 0; i < files.Length; i++)
                    {
                        string filename = "";
                        if (files[i].Extension == ".AMF")
                        {
                            filename = path + "\\" + files[i].ToString();

                            ManagerAmfparsing(filename);
                            Thread.Sleep(300);
                            string BackupPath = System.IO.Path.Combine(TACHO2_path + "\\TMF\\BACKUP", "Retried_Files");

                            // Create the subfolder
                            System.IO.Directory.CreateDirectory(BackupPath);

                            System.IO.File.Move(filename, BackupPath + "\\" + files[i].ToString());

                            
                  
                        }
                    }
                    isBackup = false; // 다 전송후 
                    //////////////////////   paidTime 서버에 데이터에 업데이트하자/////////////////////////////////

                }

                /*
                TimeCnt =0;
                string TmpFile = Application.StartupPath + "\\test.AMF";
                AMF_Data(TmpFile);

                if (CashierMode == true)   // 네트 워크 체크 기능이 필요하다. 오류시 백업모드로 변환하자. 
                {
                    if (NetworkInterface.GetIsNetworkAvailable())
                    {
                        ManagerAmfparsing(TmpFile);
                        //  MessageBox.Show("네트워크 사용가능");
              
         
                       

                    }
                    else
                    {

                    

                    }

                }*/

              /*  if (ViewerMode == true)
                {
                    TimeCnt = 0;
                    Sharethread = new Thread(new ThreadStart(ShareRun));
                    Sharethread.IsBackground = true;
                    Thread.Sleep(100);
                    Sharethread.Start();
                }*/

            }
          
          /*  if (TimeCnt == 5)
            {
                TimeCnt = 0;
                if (svr != null)
                {
                         Client_label.Text =  "Client : "+ svr.ClientCnt.ToString();


              
                   if (svr.ClientCnt == 0)
                   {
                       client1_Visible(false);
                       client2_Visible(false);
                       client3_Visible(false);
                       client4_Visible(false);
                   }
                   else if (svr.ClientCnt == 1)
                   {

                       client1_Visible(true);
                       client2_Visible(false);
                       client3_Visible(false);
                       client4_Visible(false);
                   }
                   else if (svr.ClientCnt == 2)
                   {

                       client1_Visible(true);
                       client2_Visible(true);
                       client3_Visible(false);
                       client4_Visible(false);
                   }
                   else if (svr.ClientCnt == 3)
                   {

                       client1_Visible(true);
                       client2_Visible(true);
                       client3_Visible(true);
                       client4_Visible(false);
                   }
                   else if (svr.ClientCnt >= 4)
                   {

                       client1_Visible(true);
                       client2_Visible(true);
                       client3_Visible(true);
                       client4_Visible(true);
                   }
               }

             
             //   CheckSend();
            }*/
       

         //   TestSend();
            /*
            if (ClientTimeoutcnt1 == 2)
            {
                ClientTimeoutcnt1 = 0;
                if (eyChild[0] != null)
                {
                    if (eyChild[0].ClientCheck == true)
                    {
                        eyChild[0].ClientCheck = false;
                        IDList.Remove("1");
                        if (!que.Contains(1))
                        {
                            client1_Visible(false);
                            que.Enqueue(1);        // ID 반납
                         //   ClientList.Remove(strClientID);

                            eyChild[0].Disconnet();
                            eyChild[0].SendMassge();
                         

                        }

                    }
                }
            }
           if (ClientTimeoutcnt2 == 2)
            {
                ClientTimeoutcnt2 = 0;
                if (eyChild[1] != null)
                {

                    if (eyChild[1].ClientCheck == true)
                    {
                        eyChild[1].ClientCheck = false;
                        IDList.Remove("2");
                        if (!que.Contains(2))
                        {
                            //  form1.Client2_Icon.Visible = false;
                            client2_Visible(false);
                            que.Enqueue(2);        // ID 반납
                          //  ClientList.Remove(strClientID);
                            eyChild[1].Disconnet();
                        }
                    }
                }
            }
            if (ClientTimeoutcnt3 == 2)
            {
                ClientTimeoutcnt3 = 0;
                if (eyChild[2] != null)
                {

                    if (eyChild[2].ClientCheck == true)
                    {
                        eyChild[2].ClientCheck = false;
                        IDList.Remove("3");
                        if (!que.Contains(3))
                        {
                            // form1.Client3_Icon.Visible = false;
                            client3_Visible(false);
                            que.Enqueue(3);        // ID 반납
                        //    ClientList.Remove(strClientID);
                            eyChild[2].Disconnet();
                        }
                    }
                }
            }
            if (ClientTimeoutcnt4 == 2)
            {
                ClientTimeoutcnt4 = 0;
                if (eyChild[3] != null)
                {

                    if (eyChild[3].ClientCheck == true)
                    {
                        eyChild[3].ClientCheck = false;
                        IDList.Remove("4");
                        if (!que.Contains(4))
                        {
                            // form1.Client4_Icon.Visible = false;
                            client4_Visible(false);
                            que.Enqueue(4);        // ID 반납
                          //  ClientList.Remove(strClientID);
                            eyChild[3].Disconnet();
                        }
                    }
                }
            }

             */  
                     

            if (TIME1_ENABLE == true)
            {
                if (Time1_str == CurTime)
                {
                  //  Excel_print();
                    TransData1();
                    TransData2();
                    TransData3();
                    TransData4();

                    listView1.Items.Clear();
                    listviewID = 0;
                }
            }
           
        }

        private void button15_Click()
        {
            throw new NotImplementedException();
        }
        int testid = 0;
    
        public void CheckSend()
        {
            /*
            string strtemp = "Z";
            if (IDList.Count > 1)
            {
                TimeCnt = 0;
                for (int q = 0; q < IDList.Count; q++)
                {
                    if (eyChild[0] != null)
                    {
                        if (IDList[q] == "1")
                        {
                            if (eyChild[0].ClientCheck == false)
                            {
                                eyChild[0].ClientCheck = true;
                                ClientTimeoutcnt1 = 0;
                            }
                            else
                            {
                                ClientTimeoutcnt1++;
                            }
                        }
                    }
                    if (eyChild[1] != null)
                    {
                        if (IDList[q] == "2")
                        {
                            if (eyChild[1].ClientCheck == false)
                            {
                                eyChild[1].ClientCheck = true;
                                ClientTimeoutcnt2 = 0;
                            }
                            else
                            {
                                ClientTimeoutcnt2++;
                            }
                        }
                    }
                    if (eyChild[2] != null)
                    {
                        if (IDList[q] == "3")
                        {
                            if (eyChild[2].ClientCheck == false)
                            {
                                eyChild[2].ClientCheck = true;
                                ClientTimeoutcnt3 = 0;

                            }
                            else
                            {
                                ClientTimeoutcnt3++;
                            }
                        }
                    }
                    if (eyChild[3] != null)
                    {
                        if (IDList[q] == "4")
                        {
                            if (eyChild[3].ClientCheck == false)
                            {
                                eyChild[3].ClientCheck = true;
                                ClientTimeoutcnt4 = 0;
                            }
                            else
                            {
                                ClientTimeoutcnt4++;
                            }
                        }
                    }

                    SendMsg(IDList[q], strtemp);

                }

            }
            else
            {
                TimeCnt = 0;
                if (IDList.Count == 1)
                {
                    if (IDList[0] == "1")
                    {
                        if (eyChild[0].ClientCheck == false)
                        {
                            eyChild[0].ClientCheck = true;
                            ClientTimeoutcnt1 = 0;
                        }
                        else
                        {
                            ClientTimeoutcnt1++;
                        }
                    }
                    else if (IDList[0] == "2")
                    {
                        if (eyChild[1].ClientCheck == false)
                        {
                            eyChild[1].ClientCheck = true;
                            ClientTimeoutcnt2 = 0;
                        }
                        else
                        {
                            ClientTimeoutcnt2++;
                        }

                       // eyChild[1].ClientCheck = true;
                    }
                    else if (IDList[0] == "3")
                    {
                        if (eyChild[2].ClientCheck == false)
                        {
                            eyChild[2].ClientCheck = true;
                            ClientTimeoutcnt3 = 0;
                        }
                        else
                        {
                            ClientTimeoutcnt3++;
                        }
                       // eyChild[2].ClientCheck = true;
                    }
                    else if (IDList[0] == "4")
                    {
                        if (eyChild[3].ClientCheck == false)
                        {
                            eyChild[3].ClientCheck = true;
                            ClientTimeoutcnt4 = 0;
                        }
                        else
                        {
                            ClientTimeoutcnt4++;
                        }
                       // eyChild[3].ClientCheck = true;
                    }

                    SendMsg(IDList[0], strtemp);

                
                }
            }
             * */
        }
        private void button10_Click(object sender, EventArgs e)
        {
            /*
            byte[] Data = new byte[2];

            Data[0] = 0x55;
            Data[1] = 0xA5;

            serialPort1.DiscardOutBuffer();
            serialPort1.DiscardInBuffer();
                serialPort1.Write(Data, 0, Data.Length);
            */
            /*
             testid++;
            string strtemp = "";
            strtemp = "A" + testid.ToString();         // id

            strtemp += "B" + "11-1234";         // CarNumber

            DateTime Outtime = new DateTime(2014, 1, 1, 1, 1, 1);
            strtemp += "C" + Outtime.ToString("yyyy-MM-dd tt HH:mm");         // OutTime



            DateTime Intime = new DateTime(2014, 1, 1, 2, 2, 2);
            strtemp += "D" + Intime.ToString("yyyy-MM-dd tt HH:mm");         // InTime

            double TotalDistance = 100;
            strtemp += "E" + string.Format("{0:N} Km", TotalDistance);         // 주행거리

            double TotalSalesDist = 1000;
            strtemp += "F" + string.Format("{0:N} Km", TotalSalesDist);         // 영업거리

            int TodayIncomeMoney = 220000;
            strtemp += "G" + string.Format("{0:C}", TodayIncomeMoney);         //미터 수입

            strtemp += "H";
            svr.TachoSend1 = true;
            svr.TachoStr = strtemp;
            /*
            /*
            if (ServerOpen == true)
            {

                ServerThread.Abort();

                if (ServerSocket != null)
                {
                    ServerSocket.Close();
                   
                }

          

                ServerOpen = false;
                label32.Text = "Server Close";
            
            }
            else
            {
                
                ServerThread = new Thread(new ThreadStart(Thread_Run));
                ServerThread.IsBackground = true;
                ServerThread.Start();
            }
             */
              
            
            /*
            if (ServerOpen == true)
            {
             
                ServerThread.Abort();
         
                if (ServerSocket != null)
                {
                    ServerSocket.Close();
                }

             //   ServerThread.Abort();
            
                ServerOpen = false;
                label32.Text = "Server Close";
                button10.Enabled = false;
            }
           
              ServerOpen = true;
           

             

            if (ServerOpen == true)
            {
                button10.Enabled = true;
                ServerThread = new Thread(new ThreadStart(Thread_Run));
                ServerThread.IsBackground = true;
                ServerThread.Start();
            }
            */
            
            /*
            if (ClientList.Count > 1)
            {
               
                for (int q = 0; q < IDList.Count; q++)
                {
                    if (IDList[q] == "1")
                    {
                        eyChild[0].Disconnet();
                    }
                    if (IDList[q] == "2")
                    {
                        eyChild[1].Disconnet();
                    }
                    if (IDList[q] == "3")
                    {
                        eyChild[2].Disconnet();
                    }
                    if (IDList[q] == "4")
                    {
                        eyChild[3].Disconnet();
                    }

                   

                }

            }
            else
            {
              
                if (IDList.Count == 1)
                {
                    if (IDList[0] == "1")
                    {
                        eyChild[0].Disconnet();
                    }
                    else if (IDList[0] == "2")
                    {
                        eyChild[1].Disconnet();
                    }
                    else if (IDList[0] == "3")
                    {
                        eyChild[2].Disconnet();
                    }
                    else if (IDList[0] == "4")
                    {
                        eyChild[3].Disconnet();
                    }
                }
            }

            */
         //   CheckSend();


         //   TestSend();

       
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public void BackupAmfparsing(string FileName)
        {
            DateTime DbTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day); //현재 시간 

            string FolderName = "TACHO\\" + DbTime.Year.ToString() + "\\" + DbTime.Month.ToString();


            string newPath = TACHO2_path + "TMF\\BACKUP";

         //   newPath = @"\\" + ShareIP + "\\tacho2\\" + FolderName;




            // Create the subfolder
            System.IO.Directory.CreateDirectory(newPath);

            string DBstring = String.Format(@"{0:yyMMdd}.mdb", DbTime);
            string DayDbName_str = "BackupDB.mdb";

            string filename = "TachoSample.mdb";
            OleDbConnection conn;

            // mdbfilename = newPath + "\\" + DayDbName_str;
            if (System.IO.File.Exists(newPath + "\\" + DayDbName_str))  // 같은 파일의이름이 존재 함
            {



                DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + newPath + "\\" + DayDbName_str;
                conn = new OleDbConnection(@DBstring);

                //      conn.Open();

            }
            else
            {

                string Filesource = System.IO.Path.Combine(Application.StartupPath, filename);
                System.IO.File.Copy(Filesource, newPath + "\\" + DayDbName_str);   // 일별 만큼 파일 생성, 년도 폴더에 생성한다
                DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + newPath + "\\" + DayDbName_str;
                conn = new OleDbConnection(@DBstring);
                // conn.Open();



            }


            string strFname = FileName;
            FileStream fs = new FileStream(strFname, FileMode.Open, FileAccess.ReadWrite);
            BinaryReader br = new BinaryReader(fs);
            BinaryWriter bw = new BinaryWriter(fs);
            try
            {


                if (fs.Length == 0)
                {
                    // workerObject.RequestStop();

                    return;
                }

                byte CheckSum = 0;

                /*************  Check Sum ***********************/

                /*  while (fs.Position < fs.Length-1)
                  {
                      CheckSum += br.ReadByte();
                  }
                  fs.Position = 0x00;*/
                /**************************************************/

                byte[] temp = new byte[9];
                ////////  상세 영업을 저장하자 !
                int DataCnt = (int)(fs.Length - 256 - 2);

                DateTime OutTime = new DateTime(2014, 1, 1, 1, 1, 1);
                DateTime InTime = new DateTime(2014, 1, 1, 1, 1, 1);
                byte[] Sales_Detail = new byte[DataCnt];
                if (DataCnt % 64 == 0)
                {

                    fs.Position = 0x100;
                    Sales_Detail = br.ReadBytes(DataCnt);
                }
                ////////////////////////////// dotsetting
                fs.Position = 0x10;
                temp = br.ReadBytes(1);
                DotValue = temp[0];

                ///////////////////////////  comsetting
                fs.Position = 0x12;
                temp = br.ReadBytes(1);

                if (temp[0] == 0x03)
                {
                    ibutonCheck = true;
                }
                /////////////////////////////////

                fs.Position = 0x50;     // Preview Grand Total Money
                temp = br.ReadBytes(6);
                double Prevmoney = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000) +
                                             (BcdToDecimal(temp[3]) * 1000000) + (BcdToDecimal(temp[4]) * 100000000);


                if (DotValue == 0x04)
                {
                    Prevmoney = Prevmoney / 100;
                }
                else if (DotValue == 0x08)
                {
                    Prevmoney = Prevmoney / 1000;
                }



                fs.Position = 0x67;
                temp = br.ReadBytes(7);

                try
                {
                    InTime = new DateTime(BcdToDecimal(temp[5]) + 2000, BcdToDecimal(temp[4]), BcdToDecimal(temp[3]), BcdToDecimal(temp[2]),
                                    BcdToDecimal(temp[1]), BcdToDecimal(temp[0]));
                }
                catch
                {
                    //  string NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}",
                    //                (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                    InTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                    OutTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                }
                /////////////////////
                string DriverName = "";
                try
                {
                    fs.Position = 0x70;
                    byte[] Drivername_byte = new byte[32];
                    Drivername_byte = br.ReadBytes(32);

                    DriverName = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}" +
                                                                   "{8:C2}{9:C2}{10:C2}{11:C2}{12:C2}{13:C2}{14:C2}{15:C2}" +
                                                                   "{16:C2}{17:C2}{18:C2}{19:C2}{20:C2}{21:C2}{22:C2}{23:C2}" +
                                                                   "{24:C2}{25:C2}{26:C2}{27:C2}{28:C2}{29:C2}{30:C2}{31:C2}"
                                           , Convert.ToChar(Drivername_byte[0]), Convert.ToChar(Drivername_byte[1]), Convert.ToChar(Drivername_byte[2])
                                           , Convert.ToChar(Drivername_byte[3]), Convert.ToChar(Drivername_byte[4]), Convert.ToChar(Drivername_byte[5])
                                           , Convert.ToChar(Drivername_byte[6]), Convert.ToChar(Drivername_byte[7]), Convert.ToChar(Drivername_byte[8])
                                           , Convert.ToChar(Drivername_byte[9]), Convert.ToChar(Drivername_byte[10]), Convert.ToChar(Drivername_byte[11])
                                           , Convert.ToChar(Drivername_byte[12]), Convert.ToChar(Drivername_byte[13]), Convert.ToChar(Drivername_byte[14])
                                           , Convert.ToChar(Drivername_byte[15]), Convert.ToChar(Drivername_byte[16]), Convert.ToChar(Drivername_byte[17])
                                           , Convert.ToChar(Drivername_byte[18]), Convert.ToChar(Drivername_byte[19]), Convert.ToChar(Drivername_byte[20])
                                           , Convert.ToChar(Drivername_byte[21]), Convert.ToChar(Drivername_byte[22]), Convert.ToChar(Drivername_byte[23])
                                           , Convert.ToChar(Drivername_byte[24]), Convert.ToChar(Drivername_byte[25]), Convert.ToChar(Drivername_byte[26])
                                           , Convert.ToChar(Drivername_byte[27]), Convert.ToChar(Drivername_byte[28]), Convert.ToChar(Drivername_byte[29])
                                           , Convert.ToChar(Drivername_byte[30]), Convert.ToChar(Drivername_byte[31]));
                }
                catch (Exception ex)
                {
                    DriverName = "";
                }

                ////////////////////////////
                fs.Position = 0xB4;
                temp = br.ReadBytes(9);

                string CarrNum = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}{8:C2}"
                            , Convert.ToChar(temp[8]), Convert.ToChar(temp[7]), Convert.ToChar(temp[6])
                            , Convert.ToChar(temp[5]), Convert.ToChar(temp[4]), Convert.ToChar(temp[3])
                            , Convert.ToChar(temp[2]), Convert.ToChar(temp[1]), Convert.ToChar(temp[0]));


                fs.Position = 0xc0;
                temp = br.ReadBytes(9);

                string DriverID = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}{8:C2}"
                            , Convert.ToChar(temp[8]), Convert.ToChar(temp[7]), Convert.ToChar(temp[6])
                            , Convert.ToChar(temp[5]), Convert.ToChar(temp[4]), Convert.ToChar(temp[3])
                            , Convert.ToChar(temp[2]), Convert.ToChar(temp[1]), Convert.ToChar(temp[0]));


                fs.Position = 0xF0;
                temp = br.ReadBytes(5);

                string Model = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}"
                              , Convert.ToChar(temp[0]), Convert.ToChar(temp[1]), Convert.ToChar(temp[2])
                              , Convert.ToChar(temp[3]), Convert.ToChar(temp[4]));


                fs.Position = 0x100;

                double Intmoney = 0;
                double call = 0;
                double lugg = 0;
                double ap = 0;
                double extra = 0;
                double toll = 0;



                int SalesCount = 0;

                double TotalMoney = 0;
                double TotalDist = 0;
                double TotalSalesDist = 0;
                bool FirstTacho = false;
                int SalesLength = 0;
                while (fs.Position != (fs.Length - 2))
                {


                    SalesLength++;
                    temp = br.ReadBytes(3);  // 횟수
                    // SalesCount = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000);

                    double Dist = 0;
                    double SalesDist = 0;
                    temp = br.ReadBytes(6);  //  Outtime

                    if (temp[0] != 0xFF)
                    {
                        // ListViewItem a = new ListViewItem(SalesCount.ToString());


                        if (temp[1] == 0x00)
                        {
                            temp[1] = 0x01;
                        }



                        if (temp[0] == 0x00)
                        {
                            temp[0] = 0x14;
                            temp[1] = 0x1;
                            temp[2] = 0x1;
                        }
                        if (FirstTacho == false)
                        {
                            FirstTacho = true;
                            OutTime = new DateTime(BcdToDecimal(temp[0]) + 2000, BcdToDecimal(temp[1]), BcdToDecimal(temp[2]), BcdToDecimal(temp[3]),
                               BcdToDecimal(temp[4]), BcdToDecimal(temp[5]));
                        }

                        temp = br.ReadBytes(6);  //  Intime

                        if (temp[0] == 0x00)
                        {
                            temp[0] = 0x14;
                            temp[1] = 0x1;
                            temp[2] = 0x1;
                        }
                        if (temp[1] == 0x00)
                        {
                            temp[1] = 0x01;
                        }
                        if (ibutonCheck == true)
                        {
                            InTime = new DateTime(BcdToDecimal(temp[0]) + 2000, BcdToDecimal(temp[1]), BcdToDecimal(temp[2]), BcdToDecimal(temp[3]),
                            BcdToDecimal(temp[4]), BcdToDecimal(temp[5]));

                        }
                        //  InTime = new DateTime(BcdToDecimal(temp[0]) + 2000, BcdToDecimal(temp[1]), BcdToDecimal(temp[2]), BcdToDecimal(temp[3]),
                        //     BcdToDecimal(temp[4]), BcdToDecimal(temp[5]));

                        br.ReadBytes(33);

                        temp = br.ReadBytes(3);  // 토탈 거리 
                        Dist = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000);
                        Dist = Dist / 1000;
                        TotalDist += Dist;


                        br.ReadBytes(4);  // 속도  + 센서 


                        temp = br.ReadBytes(5);
                        Intmoney = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000) +
                                         (BcdToDecimal(temp[3]) * 1000000) + (BcdToDecimal(temp[4]) * 100000000);

                        if (DotValue == 0x04)
                        {
                            Intmoney = Intmoney / 100;
                        }
                        else if (DotValue == 0x08)
                        {
                            Intmoney = Intmoney / 1000;
                        }

                        TotalMoney += Intmoney;

                        temp = br.ReadBytes(4);// Dist
                        SalesDist = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000);
                        SalesDist = SalesDist / 1000;
                        TotalSalesDist += SalesDist;



                        // total 64
                        string strTemp = string.Format("{0:N} Km", Dist);

                        //  a.SubItems.Add(CarrNum);
                        //  a.SubItems.Add(OutTime.ToString());
                        //   a.SubItems.Add(InTime.ToString());
                        //  a.SubItems.Add(Intmoney.ToString());
                        //  a.SubItems.Add(strTemp);
                        //  listView1.Items.Add(a);
                    }
                    else
                    {
                        br.ReadBytes(55);
                    }

                    //  DateTime[] DayDBName = new DateTime[40];

                }

                if (TotalMoney < 0 || Prevmoney < 0)
                {
                    IbuttonReadCheck = true;
                    return;
                }


                SalesLength = SalesLength * 64;

                try
                {
                    conn.Open();
                    OleDbCommand commTblTacho;

                    // Fill DB - TblTacho
                    string queryTblTacho = "Insert into TblTacho (TaxiID,DriverNo,DriverName,OutTime, InTime, Income, SalesDist,TotalDist,Sales_Detail,PreviousMoney,SalesLength,DotValue,CashierID"
                                                 + ") values(?,?,?,?,?,?,?,?,?,?,?,?,?)";
                    commTblTacho = new OleDbCommand(queryTblTacho, conn);


                    commTblTacho.Parameters.Add("TaxiID", OleDbType.Char).Value = CarrNum;
                    commTblTacho.Parameters.Add("DriverNo", OleDbType.Char).Value = DriverID;
                    commTblTacho.Parameters.Add("DriverName", OleDbType.Char).Value = DriverName;
                    commTblTacho.Parameters.Add("OutTime", OleDbType.Date).Value = OutTime;
                    commTblTacho.Parameters.Add("InTime", OleDbType.Date).Value = InTime;
                    commTblTacho.Parameters.Add("Income", OleDbType.Double).Value = TotalMoney;  // 11.06.27 추가
                    commTblTacho.Parameters.Add("SalesDist", OleDbType.Double).Value = TotalSalesDist;
                    commTblTacho.Parameters.Add("TotalDist", OleDbType.Double).Value = TotalDist;
                    commTblTacho.Parameters.Add("Sales_Detail", OleDbType.Binary).Value = Sales_Detail;
                    commTblTacho.Parameters.Add("PreviousMoney", OleDbType.Currency).Value = Prevmoney; //previous money
                    commTblTacho.Parameters.Add("SalesLength", OleDbType.Decimal).Value = SalesLength;   // SalesLength/64 = 갯수
                    commTblTacho.Parameters.Add("DotValue", OleDbType.Decimal).Value = DotValue;  // dot 표시 하기 위하여 
                    commTblTacho.Parameters.Add("CashierID", OleDbType.Char).Value = CashierID;
                    commTblTacho.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    conn.Close();
                    fs.Close();
                    br.Close();
                    bw.Close();
                    IbuttonReadCheck = true;
                    Thread.Sleep(2000);
                    AMF_Data(FileName);
                    //     MessageBox.Show(ex.Message);
                    return;

                }
                finally
                {
                    conn.Close();
                    fs.Close();
                    br.Close();
                    bw.Close();

                }

                //     total.Money = TotalMoney;
                //     total.Distance = TotalDist;
                // 



            }
            catch (Exception ex)
            {
                conn.Close();
                fs.Close();
                br.Close();
                bw.Close();
                //   MessageBox.Show(ex.Message);
                IbuttonReadCheck = true;

                return;

            }

            //   Treeview_Refresh();
            //    DB_ReadData(0, 1);
            //   listView1.Items[listView1.Items.Count - 2].Selected = true;
        }

        public void ManagerAmfparsing(string FileName)
        {
            DateTime DbTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day); //현재 시간 

            string FolderName = "TACHO\\" + DbTime.Year.ToString() + "\\" + DbTime.Month.ToString();


            string newPath = TACHO2_path + FolderName;
         
                newPath = @"\\" + ShareIP + "\\tacho2\\" + FolderName;
              
                string strFname = FileName;
                FileStream fs = new FileStream(strFname, FileMode.Open, FileAccess.ReadWrite);
                BinaryReader br = new BinaryReader(fs);
                BinaryWriter bw = new BinaryWriter(fs);
                OleDbConnection conn= null;
                try
                {


                    // Create the subfolder
                    System.IO.Directory.CreateDirectory(newPath);

                    string DBstring = String.Format(@"{0:yyMMdd}.mdb", DbTime);
                    string DayDbName_str = String.Format(@"{0:yyMMdd}.mdb", DbTime);

                    string filename = "TachoSample.mdb";


                    // mdbfilename = newPath + "\\" + DayDbName_str;
                    if (System.IO.File.Exists(newPath + "\\" + DayDbName_str))  // 같은 파일의이름이 존재 함
                    {



                        DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + newPath + "\\" + DayDbName_str;
                        conn = new OleDbConnection(@DBstring);

                        //      conn.Open();

                    }
                    else
                    {

                        string Filesource = System.IO.Path.Combine(Application.StartupPath, filename);
                        System.IO.File.Copy(Filesource, newPath + "\\" + DayDbName_str);   // 일별 만큼 파일 생성, 년도 폴더에 생성한다
                        DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + newPath + "\\" + DayDbName_str;
                        conn = new OleDbConnection(@DBstring);
                        // conn.Open();



                    }



                    // try
                    // {


                    if (fs.Length == 0)
                    {
                        // workerObject.RequestStop();

                        return;
                    }

                    byte CheckSum = 0;

                    /*************  Check Sum ***********************/

                    /*  while (fs.Position < fs.Length-1)
                      {
                          CheckSum += br.ReadByte();
                      }
                      fs.Position = 0x00;*/
                    /**************************************************/

                    byte[] temp = new byte[9];
                    ////////  상세 영업을 저장하자 !
                    int DataCnt = (int)(fs.Length - 256 - 2);

                    DateTime OutTime = new DateTime(2014, 1, 1, 1, 1, 1);
                    DateTime InTime = new DateTime(2014, 1, 1, 1, 1, 1);
                    byte[] Sales_Detail = new byte[DataCnt];
                    if (DataCnt % 64 == 0)
                    {

                        fs.Position = 0x100;
                        Sales_Detail = br.ReadBytes(DataCnt);
                    }
                    ////////////////////////////// dot setting
                    fs.Position = 0x10;
                    temp = br.ReadBytes(1);
                    DotValue = temp[0];

                    ///////////////////////////  com setting
                    fs.Position = 0x12;
                    temp = br.ReadBytes(1);

                    if (temp[0] == 0x03)
                    {
                        ibutonCheck = true;
                    }
                    /////////////////////////////////

                    fs.Position = 0x50;     // Preview Grand Total Money
                    temp = br.ReadBytes(6);
                    double Prevmoney = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000) +
                                                 (BcdToDecimal(temp[3]) * 1000000) + (BcdToDecimal(temp[4]) * 100000000);


                    if (DotValue == 0x04)
                    {
                        Prevmoney = Prevmoney / 100;
                    }
                    else if (DotValue == 0x08)
                    {
                        Prevmoney = Prevmoney / 1000;
                    }



                    fs.Position = 0x67;
                    temp = br.ReadBytes(7);

                    try
                    {
                        InTime = new DateTime(BcdToDecimal(temp[5]) + 2000, BcdToDecimal(temp[4]), BcdToDecimal(temp[3]), BcdToDecimal(temp[2]),
                                        BcdToDecimal(temp[1]), BcdToDecimal(temp[0]));
                    }
                    catch
                    {
                        //  string NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}",
                        //                (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                        InTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                        OutTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                    }
                    /////////////////////
                    string DriverName = "";
                    try
                    {
                        fs.Position = 0x70;
                        byte[] Drivername_byte = new byte[32];
                        Drivername_byte = br.ReadBytes(32);

                        DriverName = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}" +
                                                                       "{8:C2}{9:C2}{10:C2}{11:C2}{12:C2}{13:C2}{14:C2}{15:C2}" +
                                                                       "{16:C2}{17:C2}{18:C2}{19:C2}{20:C2}{21:C2}{22:C2}{23:C2}" +
                                                                       "{24:C2}{25:C2}{26:C2}{27:C2}{28:C2}{29:C2}{30:C2}{31:C2}"
                                               , Convert.ToChar(Drivername_byte[0]), Convert.ToChar(Drivername_byte[1]), Convert.ToChar(Drivername_byte[2])
                                               , Convert.ToChar(Drivername_byte[3]), Convert.ToChar(Drivername_byte[4]), Convert.ToChar(Drivername_byte[5])
                                               , Convert.ToChar(Drivername_byte[6]), Convert.ToChar(Drivername_byte[7]), Convert.ToChar(Drivername_byte[8])
                                               , Convert.ToChar(Drivername_byte[9]), Convert.ToChar(Drivername_byte[10]), Convert.ToChar(Drivername_byte[11])
                                               , Convert.ToChar(Drivername_byte[12]), Convert.ToChar(Drivername_byte[13]), Convert.ToChar(Drivername_byte[14])
                                               , Convert.ToChar(Drivername_byte[15]), Convert.ToChar(Drivername_byte[16]), Convert.ToChar(Drivername_byte[17])
                                               , Convert.ToChar(Drivername_byte[18]), Convert.ToChar(Drivername_byte[19]), Convert.ToChar(Drivername_byte[20])
                                               , Convert.ToChar(Drivername_byte[21]), Convert.ToChar(Drivername_byte[22]), Convert.ToChar(Drivername_byte[23])
                                               , Convert.ToChar(Drivername_byte[24]), Convert.ToChar(Drivername_byte[25]), Convert.ToChar(Drivername_byte[26])
                                               , Convert.ToChar(Drivername_byte[27]), Convert.ToChar(Drivername_byte[28]), Convert.ToChar(Drivername_byte[29])
                                               , Convert.ToChar(Drivername_byte[30]), Convert.ToChar(Drivername_byte[31]));
                    }
                    catch (Exception ex)
                    {
                        DriverName = "";
                    }

                    ////////////////////////////
                    fs.Position = 0xB4;
                    temp = br.ReadBytes(9);

                    string CarrNum = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}{8:C2}"
                                , Convert.ToChar(temp[8]), Convert.ToChar(temp[7]), Convert.ToChar(temp[6])
                                , Convert.ToChar(temp[5]), Convert.ToChar(temp[4]), Convert.ToChar(temp[3])
                                , Convert.ToChar(temp[2]), Convert.ToChar(temp[1]), Convert.ToChar(temp[0]));


                    fs.Position = 0xc0;
                    temp = br.ReadBytes(9);

                    string DriverID = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}{8:C2}"
                                , Convert.ToChar(temp[8]), Convert.ToChar(temp[7]), Convert.ToChar(temp[6])
                                , Convert.ToChar(temp[5]), Convert.ToChar(temp[4]), Convert.ToChar(temp[3])
                                , Convert.ToChar(temp[2]), Convert.ToChar(temp[1]), Convert.ToChar(temp[0]));


                    fs.Position = 0xF0;
                    temp = br.ReadBytes(5);

                    string Model = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}"
                                  , Convert.ToChar(temp[0]), Convert.ToChar(temp[1]), Convert.ToChar(temp[2])
                                  , Convert.ToChar(temp[3]), Convert.ToChar(temp[4]));


                    fs.Position = 0x100;

                    double Intmoney = 0;
                    double call = 0;
                    double lugg = 0;
                    double ap = 0;
                    double extra = 0;
                    double toll = 0;

                    int SalesCount = 0;

                    double TotalMoney = 0;
                    double TotalDist = 0;
                    double TotalSalesDist = 0;
                    bool FirstTacho = false;
                    int SalesLength = 0;
                    while (fs.Position != (fs.Length - 2))
                    {


                        SalesLength++;
                        temp = br.ReadBytes(3);  // 횟수
                        // SalesCount = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000);

                        double Dist = 0;
                        double SalesDist = 0;
                        temp = br.ReadBytes(6);  //  Outtime

                        if (temp[0] != 0xFF)
                        {
                            // ListViewItem a = new ListViewItem(SalesCount.ToString());


                            if (temp[1] == 0x00)
                            {
                                temp[1] = 0x01;
                            }



                            if (temp[0] == 0x00)
                            {
                                temp[0] = 0x14;
                                temp[1] = 0x1;
                                temp[2] = 0x1;
                            }
                            if (FirstTacho == false)
                            {
                                FirstTacho = true;
                                OutTime = new DateTime(BcdToDecimal(temp[0]) + 2000, BcdToDecimal(temp[1]), BcdToDecimal(temp[2]), BcdToDecimal(temp[3]),
                                   BcdToDecimal(temp[4]), BcdToDecimal(temp[5]));
                            }

                            temp = br.ReadBytes(6);  //  Intime

                            if (temp[0] == 0x00)
                            {
                                temp[0] = 0x14;
                                temp[1] = 0x1;
                                temp[2] = 0x1;
                            }
                            if (temp[1] == 0x00)
                            {
                                temp[1] = 0x01;
                            }
                            if (ibutonCheck == true)
                            {
                                InTime = new DateTime(BcdToDecimal(temp[0]) + 2000, BcdToDecimal(temp[1]), BcdToDecimal(temp[2]), BcdToDecimal(temp[3]),
                                BcdToDecimal(temp[4]), BcdToDecimal(temp[5]));

                            }
                            //  InTime = new DateTime(BcdToDecimal(temp[0]) + 2000, BcdToDecimal(temp[1]), BcdToDecimal(temp[2]), BcdToDecimal(temp[3]),
                            //     BcdToDecimal(temp[4]), BcdToDecimal(temp[5]));

                            br.ReadBytes(33);

                            temp = br.ReadBytes(3);  // 토탈 거리 
                            Dist = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000);
                            Dist = Dist / 1000;
                            TotalDist += Dist;


                            br.ReadBytes(4);  // 속도  + 센서 


                            temp = br.ReadBytes(5);
                            Intmoney = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000) +
                                             (BcdToDecimal(temp[3]) * 1000000) + (BcdToDecimal(temp[4]) * 100000000);

                            if (DotValue == 0x04)
                            {
                                Intmoney = Intmoney / 100;
                            }
                            else if (DotValue == 0x08)
                            {
                                Intmoney = Intmoney / 1000;
                            }

                            TotalMoney += Intmoney;

                            temp = br.ReadBytes(4);// Dist
                            SalesDist = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000);
                            SalesDist = SalesDist / 1000;
                            TotalSalesDist += SalesDist;



                            // total 64
                            string strTemp = string.Format("{0:N} Km", Dist);

                            //  a.SubItems.Add(CarrNum);
                            //  a.SubItems.Add(OutTime.ToString());
                            //   a.SubItems.Add(InTime.ToString());
                            //  a.SubItems.Add(Intmoney.ToString());
                            //  a.SubItems.Add(strTemp);
                            //  listView1.Items.Add(a);
                        }
                        else
                        {
                            br.ReadBytes(55);
                        }

                        //  DateTime[] DayDBName = new DateTime[40];

                    }

                    if (TotalMoney < 0 || Prevmoney < 0)
                    {
                        IbuttonReadCheck = true;
                        return;
                    }


                    SalesLength = SalesLength * 64;

                    if (SalesLength == 0)
                    {
                        InTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                        OutTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                    }
                    //  try
                    //   {
                    conn.Open();
                    OleDbCommand commTblTacho;

                    // Fill DB - TblTacho
                    string queryTblTacho = "Insert into TblTacho (TaxiID,DriverNo,DriverName,OutTime, InTime, Income, SalesDist,TotalDist,Sales_Detail,PreviousMoney,SalesLength,DotValue,CashierID"
                                                 + ") values(?,?,?,?,?,?,?,?,?,?,?,?,?)";
                    commTblTacho = new OleDbCommand(queryTblTacho, conn);


                    commTblTacho.Parameters.Add("TaxiID", OleDbType.Char).Value = CarrNum;
                    commTblTacho.Parameters.Add("DriverNo", OleDbType.Char).Value = DriverID;
                    commTblTacho.Parameters.Add("DriverName", OleDbType.Char).Value = DriverName;
                    commTblTacho.Parameters.Add("OutTime", OleDbType.Date).Value = OutTime;
                    commTblTacho.Parameters.Add("InTime", OleDbType.Date).Value = InTime;
                    commTblTacho.Parameters.Add("Income", OleDbType.Double).Value = TotalMoney;  // 11.06.27 추가
                    commTblTacho.Parameters.Add("SalesDist", OleDbType.Double).Value = TotalSalesDist;
                    commTblTacho.Parameters.Add("TotalDist", OleDbType.Double).Value = TotalDist;
                    commTblTacho.Parameters.Add("Sales_Detail", OleDbType.Binary).Value = Sales_Detail;
                    commTblTacho.Parameters.Add("PreviousMoney", OleDbType.Currency).Value = Prevmoney; //previous money
                    commTblTacho.Parameters.Add("SalesLength", OleDbType.Decimal).Value = SalesLength;   // SalesLength/64 = 갯수
                    commTblTacho.Parameters.Add("DotValue", OleDbType.Decimal).Value = DotValue;  // dot 표시 하기 위하여 
                    commTblTacho.Parameters.Add("CashierID", OleDbType.Char).Value = CashierID;
                    commTblTacho.ExecuteNonQuery();
                    //  }
                    /*   catch (Exception ex)
                       {
                           conn.Close();
                           fs.Close();
                           br.Close();
                           bw.Close();
                           IbuttonReadCheck = true;
                         
                         
                           //     MessageBox.Show(ex.Message);
                           return;

                       }
                       finally
                       {
                           conn.Close();
                           fs.Close();
                           br.Close();
                           bw.Close();

                       }*/

                    //     total.Money = TotalMoney;
                    //     total.Distance = TotalDist;
                    // 



                    //    }
                    /*   catch (Exception ex)
                       {
                           conn.Close();
                           fs.Close();
                           br.Close();
                           bw.Close();
                           //   MessageBox.Show(ex.Message);
                           IbuttonReadCheck = true;

                           return;

                       }*/
                }
                catch (Exception err)
                {
                    // 백업하자 !!!
                    //  conn.Close();

                    if (conn != null)
                    {
                        conn.Close();
                    }
                    fs.Close();
                    br.Close();
                    bw.Close();
                    if (!NetworkInterface.GetIsNetworkAvailable())  // 네트 워크 오류 발생 ~~
                    {
                        /***************AMF file 백업기능 ********************/

                        FileStream rs = null;
                        FileStream ws = null;

                        isBackup = true;  // 백업발생 체크 

                        string BackupPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "BACKUP");
                        // Create the subfolder
                        System.IO.Directory.CreateDirectory(BackupPath);

                        string path = Application.StartupPath + "\\ErrorLog.jie";
                        using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                        {
                            sw.WriteLine("[" + DateTime.Now.ToString() + "] " + err.StackTrace);
                        }
                        MessageBox.Show(err.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Thread.Sleep(1000);
                        ManagerAmfparsing(FileName);

                        //     string BackupFile = TACHO2_path + "\\TMF\\BACKUP\\" + FileName;

                        //      System.IO.File.Copy(FileName, FileName, true);

                        //   Treeview_Refresh();
                        //    DB_ReadData(0, 1);
                        //   listView1.Items[listView1.Items.Count - 2].Selected = true;
                    }
                    else
                    {   // 네트 워크 정상이라ㅣ면 다시 시도 !!
                        string path = Application.StartupPath + "\\ErrorLog.jie";
                        using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.Default))
                        {
                            sw.WriteLine("[" + DateTime.Now.ToString() + "] " + err.StackTrace);
                        }
                        MessageBox.Show(err.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Thread.Sleep(1000);
                        ManagerAmfparsing(FileName);
                    }
                    IbuttonReadCheck = true;
                    return;

                }
                finally
                {
                    if (conn != null)
                    {
                        conn.Close();
                    }
                    fs.Close();
                    br.Close();
                    bw.Close();
                }
       }




        public void AMF_Data(string FileName)
        {
          //  List<byte> Sales_Detail_List = new List<byte>();

         

            DateTime DbTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day); //현재 시간 

            string FolderName = "TACHO\\" + DbTime.Year.ToString() +"\\"+DbTime.Month.ToString();


            string newPath = TACHO2_path + FolderName;
            if (ViewerMode == true)
            {
                newPath = @"\\" + ShareIP + "\\tacho2\\" + FolderName;
            }
            else
            {
                newPath = newPath = TACHO2_path + FolderName;
            }

          
            // Create the subfolder
            System.IO.Directory.CreateDirectory(newPath);

            string DBstring = String.Format(@"{0:yyMMdd}.mdb", DbTime);
            string DayDbName_str = String.Format(@"{0:yyMMdd}.mdb", DbTime);

            string filename = "TachoSample.mdb";
            OleDbConnection conn;
            MdbName = DayDbName_str;
            mdbfilename = newPath + "\\" + DayDbName_str;
            if (System.IO.File.Exists(newPath + "\\" + DayDbName_str))  // 같은 파일의이름이 존재 함
            {



                DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + newPath + "\\" + DayDbName_str;
                conn = new OleDbConnection(@DBstring);

          //      conn.Open();

            }
            else
            {

                string Filesource = System.IO.Path.Combine(Application.StartupPath, filename);
                System.IO.File.Copy(Filesource, newPath + "\\" + DayDbName_str);   // 일별 만큼 파일 생성, 년도 폴더에 생성한다
                DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + newPath + "\\" + DayDbName_str;
                conn = new OleDbConnection(@DBstring);
               // conn.Open();

           

            }


            string strFname = FileName;
            FileStream fs = new FileStream(strFname, FileMode.Open, FileAccess.ReadWrite);
            BinaryReader br = new BinaryReader(fs);
            BinaryWriter bw = new BinaryWriter(fs);
            try
            {
               

                if (fs.Length == 0)
                {
                    // workerObject.RequestStop();

                    return;
                }

                byte CheckSum = 0;

                /*************  Check Sum ***********************/

              /*  while (fs.Position < fs.Length-1)
                {
                    CheckSum += br.ReadByte();
                }
                fs.Position = 0x00;*/
               /**************************************************/

                   byte[] temp = new byte[9];
                ////////  상세 영업을 저장하자 !
                int DataCnt = (int)(fs.Length - 256-2);

                 DateTime OutTime = new DateTime(2014, 1, 1, 1, 1, 1);
                DateTime InTime = new DateTime(2014, 1, 1, 1, 1, 1);
                byte[] Sales_Detail = new byte[DataCnt];
                if (DataCnt % 64 == 0)
                {
                   
                    fs.Position = 0x100;
                    Sales_Detail = br.ReadBytes(DataCnt);
                }
                ////////////////////////////// dotsetting
                fs.Position = 0x10;
                temp = br.ReadBytes(1);
                DotValue = temp[0];

                ///////////////////////////  comsetting
                fs.Position = 0x12;
                temp = br.ReadBytes(1);

                if (temp[0] == 0x03)
                {
                    ibutonCheck = true;
                }
                /////////////////////////////////

                fs.Position = 0x50;     // Preview Grand Total Money
                temp = br.ReadBytes(6);
                double Prevmoney = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000) +
                                             (BcdToDecimal(temp[3]) * 1000000) + (BcdToDecimal(temp[4]) * 100000000);


                if (DotValue == 0x04)
                {
                    Prevmoney = Prevmoney / 100;
                }
                else if (DotValue == 0x08)
                {
                    Prevmoney = Prevmoney / 1000;
                }



                 fs.Position = 0x67;
                  temp = br.ReadBytes(7);

                  try
                  {
                      InTime = new DateTime(BcdToDecimal(temp[5]) + 2000, BcdToDecimal(temp[4]), BcdToDecimal(temp[3]), BcdToDecimal(temp[2]),
                                      BcdToDecimal(temp[1]), BcdToDecimal(temp[0]));
                  }
                  catch
                  {
                                //  string NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}",
                                 //                (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                                  InTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                                  OutTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                  }
                /////////////////////
                  string DriverName = "";
                  try
                  {
                      fs.Position = 0x70;
                      byte[] Drivername_byte = new byte[32];
                      Drivername_byte = br.ReadBytes(32);

                       DriverName = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}" +
                                                                      "{8:C2}{9:C2}{10:C2}{11:C2}{12:C2}{13:C2}{14:C2}{15:C2}" +
                                                                      "{16:C2}{17:C2}{18:C2}{19:C2}{20:C2}{21:C2}{22:C2}{23:C2}" +
                                                                      "{24:C2}{25:C2}{26:C2}{27:C2}{28:C2}{29:C2}{30:C2}{31:C2}"
                                              , Convert.ToChar(Drivername_byte[0]), Convert.ToChar(Drivername_byte[1]), Convert.ToChar(Drivername_byte[2])
                                              , Convert.ToChar(Drivername_byte[3]), Convert.ToChar(Drivername_byte[4]), Convert.ToChar(Drivername_byte[5])
                                              , Convert.ToChar(Drivername_byte[6]), Convert.ToChar(Drivername_byte[7]), Convert.ToChar(Drivername_byte[8])
                                              , Convert.ToChar(Drivername_byte[9]), Convert.ToChar(Drivername_byte[10]), Convert.ToChar(Drivername_byte[11])
                                              , Convert.ToChar(Drivername_byte[12]), Convert.ToChar(Drivername_byte[13]), Convert.ToChar(Drivername_byte[14])
                                              , Convert.ToChar(Drivername_byte[15]), Convert.ToChar(Drivername_byte[16]), Convert.ToChar(Drivername_byte[17])
                                              , Convert.ToChar(Drivername_byte[18]), Convert.ToChar(Drivername_byte[19]), Convert.ToChar(Drivername_byte[20])
                                              , Convert.ToChar(Drivername_byte[21]), Convert.ToChar(Drivername_byte[22]), Convert.ToChar(Drivername_byte[23])
                                              , Convert.ToChar(Drivername_byte[24]), Convert.ToChar(Drivername_byte[25]), Convert.ToChar(Drivername_byte[26])
                                              , Convert.ToChar(Drivername_byte[27]), Convert.ToChar(Drivername_byte[28]), Convert.ToChar(Drivername_byte[29])
                                              , Convert.ToChar(Drivername_byte[30]), Convert.ToChar(Drivername_byte[31]));

                       if (iButtonMode == false)
                       {
                           DriverName = "";
                       }

                       Driver_Name_temp = DriverName;
                      
                  }
                  catch (Exception ex)
                  {
                      DriverName = "";
                  }

             ////////////////////////////
                fs.Position = 0xB4;
                temp = br.ReadBytes(9);

                string CarrNum = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}{8:C2}"
                            , Convert.ToChar(temp[8]), Convert.ToChar(temp[7]), Convert.ToChar(temp[6])
                            , Convert.ToChar(temp[5]), Convert.ToChar(temp[4]), Convert.ToChar(temp[3])
                            , Convert.ToChar(temp[2]), Convert.ToChar(temp[1]), Convert.ToChar(temp[0]));


                fs.Position = 0xc0;
                temp = br.ReadBytes(9);

                string DriverID = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}{8:C2}"
                            , Convert.ToChar(temp[8]), Convert.ToChar(temp[7]), Convert.ToChar(temp[6])
                            , Convert.ToChar(temp[5]), Convert.ToChar(temp[4]), Convert.ToChar(temp[3])
                            , Convert.ToChar(temp[2]), Convert.ToChar(temp[1]), Convert.ToChar(temp[0]));
                Driver_ID_temp = DriverID;

                fs.Position = 0xF0;
                temp = br.ReadBytes(5);

                string Model = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}"
                              , Convert.ToChar(temp[0]), Convert.ToChar(temp[1]), Convert.ToChar(temp[2])
                              , Convert.ToChar(temp[3]), Convert.ToChar(temp[4]));


                fs.Position = 0x100;
               
                double Intmoney = 0;
                double call = 0;
                double lugg = 0;
                double ap = 0;
                double extra = 0;
                double toll = 0;



                int SalesCount = 0;

                double TotalMoney = 0;
                double TotalDist = 0;

                double TotalSalesDist = 0;
                double TotalVacant = 0;

                bool FirstTacho = false;
                int SalesLength = 0;
                while (fs.Position != (fs.Length - 2))
                {


                    SalesLength++;
                    temp = br.ReadBytes(3);  // 횟수
                   // SalesCount = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000);

                    double Dist=0;
                    double SalesDist = 0;
                    temp = br.ReadBytes(6);  //  Outtime
                    
                    if (temp[0] != 0xFF)
                    {
                       // ListViewItem a = new ListViewItem(SalesCount.ToString());
                       
                         
                            if (temp[1] == 0x00)
                            {
                                temp[1] = 0x01;
                            }



                            if (temp[0] == 0x00)
                            {
                                temp[0] = 0x14;
                                temp[1] = 0x1;
                                temp[2] = 0x1;
                            }
                            if (FirstTacho == false)
                            {
                                FirstTacho = true;

                                try
                                {
                                    OutTime = new DateTime(BcdToDecimal(temp[0]) + 2000, BcdToDecimal(temp[1]), BcdToDecimal(temp[2]), BcdToDecimal(temp[3]),
                                  BcdToDecimal(temp[4]), BcdToDecimal(temp[5]));
                                }
                                catch
                                {
                                  
                                  
                                    OutTime = new DateTime(2011, 1, 1, 1, 1, 1);
                                }
                               
                            }

                            temp = br.ReadBytes(6);  //  Intime

                            if (temp[0] == 0x00)
                            {
                                temp[0] = 0x14;
                                temp[1] = 0x1;
                                temp[2] = 0x1;
                            }
                            if (temp[1] == 0x00)
                            {
                                temp[1] = 0x01;
                            }
                            if (ibutonCheck == true)
                            {
                                try
                                {
                                    InTime = new DateTime(BcdToDecimal(temp[0]) + 2000, BcdToDecimal(temp[1]), BcdToDecimal(temp[2]), BcdToDecimal(temp[3]),
                                    BcdToDecimal(temp[4]), BcdToDecimal(temp[5]));
                                }
                                catch
                                {


                                    InTime = new DateTime(2011, 1, 1, 1, 1, 1);
                                }

                            }
                          //  InTime = new DateTime(BcdToDecimal(temp[0]) + 2000, BcdToDecimal(temp[1]), BcdToDecimal(temp[2]), BcdToDecimal(temp[3]),
                           //     BcdToDecimal(temp[4]), BcdToDecimal(temp[5]));

                             temp =  br.ReadBytes(5);  //Call
                             temp = br.ReadBytes(5);  //lugg
                            temp = br.ReadBytes(5);  //A/P
                              ap = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000) +
                                                  (BcdToDecimal(temp[3]) * 1000000) + (BcdToDecimal(temp[4]) * 100000000);
                             
                              if (DotValue == 0x04)
                              {
                                  ap = ap / 100;
                              }
                              else if (DotValue == 0x08)
                              {
                                  ap = ap / 1000;
                              }
                              AP_temp = string.Format("{0:N2}", ap);  // ap total
                              ///////////////////////////////////////////////////////
                              temp = br.ReadBytes(5);  //Extra

                              ///////////////////////////////////////////////////////
                              temp = br.ReadBytes(5);  //Toll
                              toll = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000) +
                                                     (BcdToDecimal(temp[3]) * 1000000) + (BcdToDecimal(temp[4]) * 100000000);

                              if (DotValue == 0x04)
                              {
                                  toll = toll / 100;
                              }
                              else if (DotValue == 0x08)
                              {
                                  toll = toll / 1000;
                              }
                              Toll_temp = string.Format("{0:N2}", toll);  // toll total
                              ///////////////////////////////////////////////////////


                              br.ReadBytes(1); // call cnt
                              br.ReadBytes(1); //lugg cnt
                              br.ReadBytes(1); // A/p cnt
                              br.ReadBytes(1); //Extra cnt
                              br.ReadBytes(1); // toll cnt
                          
                            



                            temp = br.ReadBytes(3);  // 빈차 거리 
                            Dist = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000);
                            Dist = Dist / 1000;
                            TotalVacant += Dist;
                            Vacant_dist_temp = string.Format("{0:N2}", TotalVacant);  // 영업거리
                            Dist = 0;

                            temp = br.ReadBytes(3);  // 토탈 거리 
                            Dist = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000);
                            Dist = Dist / 1000;
                            TotalDist += Dist;
                            Total_dist_temp = string.Format("{0:N2}", TotalDist);  // 토탈거리

                            br.ReadBytes(4);  // 속도  + 센서 

                           
                            temp = br.ReadBytes(5);
                            Intmoney = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000) +
                                             (BcdToDecimal(temp[3]) * 1000000) + (BcdToDecimal(temp[4]) * 100000000);

                            if (DotValue == 0x04)
                            {
                                Intmoney = Intmoney / 100;
                            }
                            else if (DotValue == 0x08)
                            {
                                Intmoney = Intmoney / 1000;
                            }

                            TotalMoney += Intmoney;
                            Total_income_temp = string.Format("{0:N2}", TotalMoney);  // 영업
                            

                            temp = br.ReadBytes(4);// Dist
                            SalesDist = BcdToDecimal(temp[0]) + (BcdToDecimal(temp[1]) * 100) + (BcdToDecimal(temp[2]) * 10000);
                            SalesDist = SalesDist / 1000;
                            TotalSalesDist += SalesDist;

                            Hired_dist_temp = string.Format("{0:N2}", TotalSalesDist);  // 영업거리



                            // total 64
                            string strTemp = string.Format("{0:N} Km", Dist);
                        
                      //  a.SubItems.Add(CarrNum);
                      //  a.SubItems.Add(OutTime.ToString());
                     //   a.SubItems.Add(InTime.ToString());
                      //  a.SubItems.Add(Intmoney.ToString());
                      //  a.SubItems.Add(strTemp);
                      //  listView1.Items.Add(a);
                    }
                    else
                    {
                        br.ReadBytes(55);
                    }

                  //  DateTime[] DayDBName = new DateTime[40];
                  
                }

                if (TotalMoney < 0 || Prevmoney<0)
                {
                    IbuttonReadCheck = true;
                    return ;
                }
               

                SalesLength = SalesLength * 64;

                if (SalesLength == 0)
                {
                    InTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                    OutTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);


                    TotalDist = (BcdToDecimal(Daily_Distance[0])) + (BcdToDecimal(Daily_Distance[1]) * 100) + (BcdToDecimal(Daily_Distance[2]) * 10000);
                       
                    TotalDist = TotalDist / 1000;
                  
                }
               // -----------------------------------------------------영수증 모드
                if (GetTime_Receipt == true)
                {
                    GetTime_Receipt = false;
                    Money_Check = TotalMoney;
   
                    return;
                }
                //*****************************************************  중복 체크를 하자!!!
                bool check = false;
                check = DuplicateDataCheck(CarrNum, OutTime, InTime);

                if (check == true)
                {
                    fs.Close();
                    br.Close();
                    bw.Close();
                    //   MessageBox.Show(ex.Message);
                    IbuttonReadCheck = true;

                    return;
                }
                  
                 
                //************************************************************


                CashierID = CashierID.Trim();

                try
                {
                    conn.Open();
                    OleDbCommand commTblTacho;
                   
                    // Fill DB - TblTacho
                    string queryTblTacho = "Insert into TblTacho (TaxiID,DriverNo,DriverName,OutTime, InTime, Income, SalesDist,TotalDist,Sales_Detail,PreviousMoney,SalesLength,DotValue,CashierID"
                                              + ") values(?,?,?,?,?,?,?,?,?,?,?,?,?)";
                    commTblTacho = new OleDbCommand(queryTblTacho, conn);


                    commTblTacho.Parameters.Add("TaxiID", OleDbType.Char).Value = CarrNum;
                    commTblTacho.Parameters.Add("DriverNo", OleDbType.Char).Value = DriverID;
                    commTblTacho.Parameters.Add("DriverName", OleDbType.Char).Value = DriverName;
                    commTblTacho.Parameters.Add("OutTime", OleDbType.Date).Value = OutTime;
                    commTblTacho.Parameters.Add("InTime", OleDbType.Date).Value = InTime;
                    commTblTacho.Parameters.Add("Income", OleDbType.Double).Value = TotalMoney;  // 11.06.27 추가
                    commTblTacho.Parameters.Add("SalesDist", OleDbType.Double).Value = TotalSalesDist;
                    commTblTacho.Parameters.Add("TotalDist", OleDbType.Double).Value = TotalDist;
                    commTblTacho.Parameters.Add("Sales_Detail", OleDbType.Binary).Value = Sales_Detail;
                    commTblTacho.Parameters.Add("PreviousMoney", OleDbType.Currency).Value = Prevmoney; //previous money
                    commTblTacho.Parameters.Add("SalesLength", OleDbType.Decimal).Value = SalesLength;   // SalesLength/64 = 갯수
                    commTblTacho.Parameters.Add("DotValue", OleDbType.Decimal).Value = DotValue;  // dot 표시 하기 위하여 
                    commTblTacho.Parameters.Add("CashierID", OleDbType.Char).Value = CashierID;
                    commTblTacho.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    conn.Close();
                    fs.Close();
                    br.Close();
                    bw.Close();
                    IbuttonReadCheck = true;
                    Thread.Sleep(2000);
                    AMF_Data(FileName);
               //     MessageBox.Show(ex.Message);
                    return;

                }
                finally
                {
                    conn.Close();
                    fs.Close();
                    br.Close();
                    bw.Close();

                }

                total.Money = TotalMoney;
                total.Distance = TotalDist;
                // 


               
            }
            catch (Exception ex)
            {
                conn.Close();
                fs.Close();
                br.Close();
                bw.Close();
             //   MessageBox.Show(ex.Message);
                IbuttonReadCheck = true;
              
                return;
           
            }
        
            Treeview_Refresh();

             DB_ReadData(0, 1);

             listView1.Items[listView1.Items.Count - 2].Selected = true;
         

        }

        public bool DuplicateDataCheck(string CarrNum, DateTime OutTime, DateTime InTime)
        {
            bool is_duplicate =false;
            int cnt = 0;
            string NameDB = "";
            //	SetAlternatingRowColors(listView1, a, b);

            try
            {

                //	selectedColumnIndex = 4;   // 출고 시간별 정렬
                //	selectedOrder = 1;
                //	FillData(selectedColumnIndex, selectedOrder);




                FileInfo file = new FileInfo(mdbfilename);

                string DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + mdbfilename;

                OleDbConnection conn = new OleDbConnection(@DBstring);
                string queryDelTblTacho;

                // Fill Data


                conn.Open();

                string queryTblTacho = "SELECT * FROM TblTacho ORDER BY ID DESC";
                OleDbCommand commTblTacho = new OleDbCommand(queryTblTacho, conn);
                OleDbDataReader srRead = commTblTacho.ExecuteReader();



                string queryTblTacho1 = "SELECT * FROM TblTacho ORDER BY ID DESC";
                OleDbCommand commTblTacho1 = new OleDbCommand(queryTblTacho1, conn);
                OleDbDataReader srRead1 = commTblTacho1.ExecuteReader();
                cnt = 0;
                while (srRead1.Read())   // db 데이터 읽어 오기 아이디 , 차량 번호 ,출고 시간 ,입고시간
                {
                    cnt++;
                }
              
                
            //    cnt = listView1.Items.Count - 1;
                string[] dbId = new string[cnt];
                DateTime[] dbOuttime = new DateTime[cnt];
                DateTime[] dbIntime = new DateTime[cnt];
                string[] dbCarnumber = new string[cnt];

                int[] intdbId = new int[cnt * 3];
                int[] RepeatId = new int[cnt * 3];
           
                int ii = 0;


                while (srRead.Read())   // db 데이터 읽어 오기 아이디 , 차량 번호 ,출고 시간 ,입고시간
                {
                    intdbId[ii] = srRead.GetInt32(0);
                    dbCarnumber[ii] = srRead.GetString(1);
                    dbOuttime[ii] = srRead.GetDateTime(4);
                    dbIntime[ii] = srRead.GetDateTime(5);
                    ii++;

                }

                int k = 0;
                int[] temp = new int[cnt * 3];
                bool pass = false;


                for (int i = 0; i < cnt; i++)
                {




                    if (dbCarnumber[i] == CarrNum && dbOuttime[i] == OutTime
                        && dbIntime[i] == InTime)  // 차번호 입고 출고 시간 중복 검사
                        {

                            is_duplicate = true;
                         

                        }

                   

                }
               
         
                conn.Close();


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

                //   conn.Close();
            }

            return is_duplicate;
        }

        string strFname_buff;
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null)
                {
                    for (int i = 0; i < files.Length; i++)
                    {
                        string str = files[i].ToLower();

                        if (str.EndsWith(".AMF") || str.EndsWith(".amf"))
                        {
                            string dirName = Path.GetFullPath(files[i]);
                            strFname_buff = dirName;
                            //  DragDrop_Check = true;
                            string strFname = "";
                            AMF_Data(strFname_buff);
                        }
                    }
                }
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            byte[] Data = new byte[2];

            Data[0] = 0xCC;
            Data[1] = 0xDD;

           

            richTextCnt = 0;
           
            linecnt = 0;
            serialPort1.Write(Data, 0, Data.Length);
            rcvList.Clear();
        }

        private void buttonDBDel_Click(object sender, EventArgs e)
        {

        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Text.Length == 6 || e.Node.Text.Length == 8)
            {
                //listView1.Clear();

                string path = "";

                if (ViewerMode == true)
                {
                    path = @"\\" + ShareIP + "\\tacho2\\" + "TACHO" + "\\" + e.Node.Parent.Parent.Text + "\\" + e.Node.Parent.Text;
                }
                else
                {
                    path = TACHO2_path + "TACHO"+ "\\" + e.Node.Parent.Parent.Text + "\\" + e.Node.Parent.Text;
                }
              //   path = TACHO2_path + e.Node.Parent.Parent.Parent.Text + "\\" + e.Node.Parent.Parent.Text + "\\" + e.Node.Parent.Text;
                string[] files = Directory.GetFiles(path, "*.mdb");


                for (int i = 0; i < files.Length; i++)
                {
                    FileInfo file = new FileInfo(files[i]);

                    files[i] = file.Name;
                    if (files[i] == e.Node.Text + ".mdb")
                    {

                        mdbfilename = path + "\\" + files[i];
                      
                        DB_ReadData(0,1);

                     

                      
                    }

                }


            }
        }
        public void DB_ReadData(int ColumnIdx, int Order)
        {
          
           int dotval = 0;
          
           FileInfo file = new FileInfo(mdbfilename);

           string DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + mdbfilename;

           OleDbConnection conn = new OleDbConnection(@DBstring);
           conn.Open();

                formname = file.Name;
                char[] trimChars = { '.', 'm', 'd', 'b' };

                formname = formname.TrimEnd(trimChars);

               // this.Text = yymmdd;


                string queryRead = "SELECT * FROM TblTacho";

              
              

                if (Order == -1)    // 정렬 없음
                {
                    queryRead += " ORDER BY ID";
                }
                else if (Order == 0) // 내림 차순
                {
                    switch (ColumnIdx)
                    {
                        case 0: queryRead += " ORDER BY ID DESC"; break;
                        case 1: queryRead += " ORDER BY TAxiID DESC"; break;
                        case 2: queryRead += " ORDER BY DriverNo DESC"; break;
                        case 3: queryRead += " ORDER BY DriverName DESC"; break;
                        case 4: queryRead += " ORDER BY OutTime DESC"; break;
                        case 5: queryRead += " ORDER BY OutTime DESC"; break;
                        case 6: queryRead += " ORDER BY InTime DESC"; break;
                        case 7: queryRead += " ORDER BY InTime DESC"; break;
                        case 8: queryRead += " ORDER BY PaidTime DESC"; break;
                        case 9: queryRead += " ORDER BY PaidTime DESC"; break;
                        case 10: queryRead += " ORDER BY Income DESC"; break;
                        case 11: queryRead += " ORDER BY TotalDist DESC"; break;
                        case 12: queryRead += " ORDER BY Income DESC"; break;
                        case 13: queryRead += " ORDER BY PreviousMoney DESC"; break;
                        case 14: queryRead += " ORDER BY ID DESC"; break;
                        case 15: queryRead += " ORDER BY ID DESC"; break;
                        case 16: queryRead += " ORDER BY ID DESC"; break;
                        default: queryRead += " ORDER BY ID"; break;
                    }
                }
                else
                {
                    switch (ColumnIdx)
                    {
                        case 0: queryRead += " ORDER BY ID "; break;
                        case 1: queryRead += " ORDER BY TAxiID "; break;
                        case 2: queryRead += " ORDER BY DriverNo "; break;
                        case 3: queryRead += " ORDER BY DriverName "; break;
                        case 4: queryRead += " ORDER BY OutTime "; break;
                        case 5: queryRead += " ORDER BY OutTime "; break;
                        case 6: queryRead += " ORDER BY InTime "; break;
                        case 7: queryRead += " ORDER BY InTime "; break;
                        case 8: queryRead += " ORDER BY PaidTime "; break;
                        case 9: queryRead += " ORDER BY PaidTime "; break;
                        case 10: queryRead += " ORDER BY Income "; break;
                        case 11: queryRead += " ORDER BY TotalDist "; break;
                        case 12: queryRead += " ORDER BY Income "; break;
                        case 13: queryRead += " ORDER BY PreviousMoney "; break;
                        case 14: queryRead += " ORDER BY ID "; break;
                        case 15: queryRead += " ORDER BY ID "; break;
                        case 16: queryRead += " ORDER BY ID "; break;
                        default: queryRead += " ORDER BY ID"; break;
                    }
                }


                OleDbCommand commRead = new OleDbCommand(queryRead, conn);
                OleDbDataReader srRead = commRead.ExecuteReader();



                if (listView1.Items.Count > 0)
                    listView1.Items.Clear();

                listView1.View = View.Details;
                listView1.GridLines = true;                   //   리스트 뷰 라인생성
                listView1.FullRowSelect = true;               // 라인 선택 */
                if (THAILAND_Set == true)
                {
                    this.listView1.Columns[2].Text = "";
                    this.listView1.Columns[3].Text = "";
                    this.listView1.Columns[12].Text = "";
                    this.listView1.Columns[13].Text = "";

                    this.listView1.Columns[1].Width = 150;  // Taxi id
                    this.listView1.Columns[2].Width = 0;    // driver id
                    this.listView1.Columns[3].Width = 0;    // driver name
                    this.listView1.Columns[4].Width = 140;  // trip start date
                    this.listView1.Columns[5].Width = 140;  // trip start time
                    this.listView1.Columns[6].Width = 140;  // trip end date
                    this.listView1.Columns[7].Width = 140;  // trip end time
                    this.listView1.Columns[10].Width = 140; // income
                    this.listView1.Columns[11].Width = 140; // distance
                    this.listView1.Columns[12].Width = 0;
                    this.listView1.Columns[13].Width = 0;
                }
                        this.listView1.Columns[2].Text = "";
                        this.listView1.Columns[3].Text = "";
                        this.listView1.Columns[12].Text = "";
                        this.listView1.Columns[13].Text = "";

                        this.listView1.Columns[1].Width = 150;  // Taxi id
                        this.listView1.Columns[2].Width = 0;    // driver id
                        this.listView1.Columns[3].Width = 0;    // driver name
                        this.listView1.Columns[4].Width = 140;  // trip start date
                        this.listView1.Columns[5].Width = 140;  // trip start time
                        this.listView1.Columns[6].Width = 140;  // trip end date
                        this.listView1.Columns[7].Width = 140;  // trip end time
                        this.listView1.Columns[10].Width = 140; // income
                        this.listView1.Columns[11].Width = 140; // distance
                        this.listView1.Columns[12].Width = 0;
                        this.listView1.Columns[13].Width = 0;

              /*  if (THAILAND_Set == true)
                {
                    this.listView1.Columns[0].Text = "ID";
                    this.listView1.Columns[1].Text = "Car Plate";
                    this.listView1.Columns[2].Text = "Start Date";
                    this.listView1.Columns[3].Text = "Start Time";
                   
                    this.listView1.Columns[4].Text = "End Date";
                    this.listView1.Columns[5].Text = "End Time";
                    this.listView1.Columns[6].Text = "";
                    this.listView1.Columns[7].Text = "";
                    this.listView1.Columns[8].Text = "Total Income";
                    this.listView1.Columns[9].Text = "";
                    this.listView1.Columns[10].Text = "Income";  // Sales Total 
                    this.listView1.Columns[11].Text = "Distance";
                    this.listView1.Columns[12].Text = "";
                    this.listView1.Columns[13].Text = "";
                
                }*/


                double TotalMoney = 0;
                double TotalDist = 0;

                while (srRead.Read())
                {

                    string tooltip_str = srRead.GetInt32(0).ToString();

                    int isOpenedDBNum = 0;
                    DateTime stTime = new DateTime(1, 1, 1, 0, 0, 0);
                    DateTime srTime = new DateTime(1, 1, 1, 0, 0, 0);
                    TimeSpan stSpan = new TimeSpan(0, 0, 0, 0);

                    uint tMoney = 0, tB = 0, tDB = 0, tDA = 0, tAB = 0, tAA = 0;
                    double tDistS = 0, tDistD = 0;
                    TimeSpan tOverD = new TimeSpan();

                    bool bIsMatched = false;

                    int nReadDBCnt = 0;         // 읽어들인 데이터 갯수 파악용

                    ListViewItem a = new ListViewItem(tooltip_str);       // ID


                  //  a.ToolTipText = " 영업 상세정보 \n ID 더블클릭!";

                    string strCarNo = srRead.GetString(1);  // 차량번호



                    a.SubItems.Add(strCarNo);                                              
                    string strDriverNo = "";
                    if (THAILAND_Set == true)
                    {
                        a.SubItems.Add("");
                        a.SubItems.Add("");
                    }
                    else
                    {
                        if (srRead.IsDBNull(2) == false)                                            // 기사번호
                        {
                            strDriverNo = srRead.GetString(2);
                            a.SubItems.Add(strDriverNo);

                        }
                        else
                        {
                            a.SubItems.Add(strDriverNo);
                        }



                        string strDriverName = "";
                        if (srRead.IsDBNull(3) == false)                                            // 기사이름
                        {
                            strDriverName = srRead.GetString(3);
                            a.SubItems.Add(strDriverName);

                        }
                        else
                        {
                            a.SubItems.Add(strDriverName);
                        }
                    }
                      

                    //a.SubItems.Add(srRead.GetString(2));                                

                   // string strDriverNo = srRead.GetString(2);


                  //  a.SubItems.Add(strDriverNo);                                           

                   

                    a.SubItems.Add(srRead.GetDateTime(4).ToString("yyyy-MM-dd"));                       // 출고 날짜 
                    a.SubItems.Add(srRead.GetDateTime(4).ToString(" HH:mm:ss"));                       // 출고 시간  

                //    a.SubItems.Add(srRead.GetDateTime(5).ToString("yyyy-MM-dd tt HH:mm:ss"));

                    a.SubItems.Add(srRead.GetDateTime(5).ToString("yyyy-MM-dd"));                       // 입고 날짜 
                    a.SubItems.Add(srRead.GetDateTime(5).ToString(" HH:mm:ss"));                       // 입고 시간  

                    if (srRead.IsDBNull(19) == false)                                            // paidTime
                    {
                        a.SubItems.Add(srRead.GetDateTime(19).ToString("yyyy-MM-dd"));                       // paidTime  날짜 
                        a.SubItems.Add(srRead.GetDateTime(19).ToString(" HH:mm"));                       // paidTime  시간  
                    }
                    else
                    {
                        a.SubItems.Add("");
                        a.SubItems.Add("");
                    }

                    double dotsatus = srRead.GetDouble(10);



                    double uuu = (double)srRead.GetDouble(6);

                    string money;
                    if (dotsatus != 4 && dotsatus != 2 && dotsatus !=8)
                    {
                        TotalMoney +=uuu;
                        int hhh = (int)uuu;
                        money = string.Format("{0:D}", (int)hhh);
                        a.SubItems.Add(money);
                    }
                    else
                    {
                        if (dotsatus ==8)
                        {
                             TotalMoney += uuu;
                             money = string.Format("{0:F3}", uuu);
                            a.SubItems.Add(money);
                        }
                        else
                        {
                            TotalMoney += uuu;
                            money = string.Format("{0:F2}", uuu);
                            a.SubItems.Add(money);
                        }
                    }// 미터수입
                    double ddd = 0;
                    if (srRead.IsDBNull(9) == false)                                            // 기사번호
                    {

                         ddd = srRead.GetDouble(9); // 영업거리 
                    }
                    

                    TotalDist += ddd;
                    string dist = string.Format("{0:N3} Km", ddd);
                    a.SubItems.Add(dist);


                    double GtandTotlaMoney = 0;
                    money = "0";
                    string GandMoenyStr = "0";
                    if (srRead.IsDBNull(7) == false)
                    {

                        GtandTotlaMoney = srRead.GetDouble(7);// preview
                        GtandTotlaMoney += srRead.GetDouble(6); ;  // Grnad Total

                        if (dotsatus != 4 && dotsatus != 2 && dotsatus !=8)
                        {
                            money = string.Format("{0:D}", (int)srRead.GetDouble(7));  // preview total
                            GandMoenyStr = string.Format("{0:D}",(int) GtandTotlaMoney);
                        }
                        else
                        {
                            if (dotsatus == 8)
                            {
                                money = string.Format("{0:F3}", srRead.GetDouble(7));
                                GandMoenyStr = string.Format("{0:F3}", GtandTotlaMoney);
                            }
                            else
                            {
                                money = string.Format("{0:F2}", srRead.GetDouble(7));
                                GandMoenyStr = string.Format("{0:F2}", GtandTotlaMoney);
                            }
                        }
                     

                     
                        a.SubItems.Add(GandMoenyStr);
                       a.SubItems.Add(money); 
                    }
                    else
                    {
                        a.SubItems.Add(GandMoenyStr);
                        a.SubItems.Add(money);
                    }

                    if (srRead.IsDBNull(11) == false)
                    {


                        if (dotsatus != 4 && dotsatus != 2 && dotsatus != 8)
                        {
                            money = string.Format("{0:D}", (int)srRead.GetDouble(11));  // 

                        }
                        else
                        {
                            if (dotsatus == 8)
                            {
                                money = string.Format("{0:F3}", srRead.GetDouble(11));

                            }
                            else
                            {
                                money = string.Format("{0:F2}", srRead.GetDouble(11));

                            }
                        }

                        a.SubItems.Add(money);

                    }
                    else
                    {
                        a.SubItems.Add("");

                    }

                    if (srRead.IsDBNull(20) == false)
                    {
                        string Cashier = srRead.GetString(20);
                        a.SubItems.Add(Cashier);

                    }
                    else
                    {
                        a.SubItems.Add("");

                    }
 
                  /*  if (srRead.IsDBNull(3) == false)
                    {
                        strDriverName = srRead.GetString(3);
                        a.SubItems.Add(strDriverName);

                    }
                    else
                    {
                        a.SubItems.Add(strDriverName);
                    }*/

                    // 기사이름
                    //	    a.SubItems.Add("");                                                      // 기사이름

                    //	string fuel = string.Format("{0:N} L", srRead.GetDouble(10));
                    //	a.SubItems.Add(fuel);      

                 //   money = string.Format("{0:C}", srRead.GetDecimal(7));  // 실입금액
                 //   a.SubItems.Add(money);    

                    /*
                    ddd = srRead.GetDouble(9);

                    dist = string.Format("{0:N} Km", ddd);
                    a.SubItems.Add(dist);                                                   // 주행거리

                    //	string fuel = string.Format("{0:N} L", srRead.GetDouble(10));
                    //	a.SubItems.Add(fuel);                                                   // 연료량

                    srTime = srRead.GetDateTime(11);
                    stTime = new DateTime(srTime.Year, srTime.Month, srTime.Day, 0, 0, 0);
                    stSpan = srTime - stTime;
                    tOverD += stSpan;
                    a.SubItems.Add(stSpan.ToString());                                      // 과속시간

                    uuu = (uint)srRead.GetInt32(12);
                    tB += uuu;
                    a.SubItems.Add(uuu.ToString());                                         // 급제동
                    uuu = (uint)srRead.GetInt32(13);
                    tDB += uuu;
                    a.SubItems.Add(uuu.ToString());                                         // 주행기본
                    uuu = (uint)srRead.GetInt32(14);
                    tDA += uuu;
                    a.SubItems.Add(uuu.ToString());                                         // 주행이후
                    uuu = (uint)srRead.GetInt32(15);
                    tAB += uuu;
                    a.SubItems.Add(uuu.ToString());                                         // 할증기본
                    uuu = (uint)srRead.GetInt32(16);
                    tAA += uuu;
                    a.SubItems.Add(uuu.ToString());                                         // 할증이후
                    //a.SubItems.Add(srRead.GetInt32(17).ToString());                         // 문개폐
                    isOpenedDBNum = srRead.GetInt32(18);


                    string fuel = string.Format("{0:N} L", srRead.GetDouble(10));
                    a.SubItems.Add(fuel);                                                   // 연료량
                    */

                    //	string fuel = "";

                    /*	if (srRead.IsDBNull(22) == false)
                        {
                            fuel = srRead.GetString(22);
                            a.SubItems.Add(fuel);
                        }
                        else
                        {
                            a.SubItems.Add(fuel);
                        } */

                    listView1.Items.Add(a);
                }
                total.Money = TotalMoney;
                total.Distance = TotalDist;

              int nTotalItems = listView1.Items.Count;

              if (nTotalItems != 0)
              {
                  ListViewItem b = new ListViewItem("SUM");
                  b.SubItems.Add("");                             ///   공백 다섯칸 만들기
                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  int valtemp = (int)total.Money;
                  double vlatemp1 = (double)valtemp;

                  string temp = "";
                  if ((total.Money % vlatemp1) != 0)
                  {

                      temp = string.Format("{0:F2}", total.Money);
                      b.SubItems.Add(temp);
                  }
                  else
                  {
                      temp = string.Format("{0:D}",(int) total.Money);
                      b.SubItems.Add(temp);
                  }

                  temp = string.Format("{0:N3} Km", total.Distance);
                  b.SubItems.Add(temp);
                  b.BackColor = System.Drawing.Color.LightGray;

                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  b.SubItems.Add("");
                  listView1.Items.Add(b);
              }


                conn.Close();


                if (iButtonMode == true)
                {
                    ListViewBackColor();
                }


                FillList(this.m_list, listView1);
        }
     
        public void ListViewBackColor()
        {
            Color AliceBlue = Color.FromArgb(240, 248, 255);
            Color LightPink = Color.FromArgb(255, 182, 193);

            for (int i = 0; i < listView1.Items.Count - 1; i++)
            {
                if (listView1.Items[i].SubItems[9].Text == "")
                {
                    listView1.Items[i].BackColor = Color.Wheat;
                }
                //listView1.Items[j].BackColor = LightPink;
            }
        }

        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (selectedColumnIndex == e.Column)
            {
                if (selectedOrder == 1)
                    selectedOrder = 0;
                else
                    selectedOrder = 1;
            }
            else
            {
                selectedColumnIndex = e.Column;

                selectedOrder = 1;
            }
            //	opendlg = false;
         //   columnclick = true;
            DB_ReadData(selectedColumnIndex, selectedOrder);
        }
        public void PrintPageSetup_FormData()
        {
            m_list.PageSetup();
        }
        private void buttonPrint_Click(object sender, EventArgs e)
        {
          
            PrintPreview_FormData();
        }
        public void PrintPreview_FormData()
        {
           
            m_list.PageSetup();
            m_list.Title = "SHIFT REPORT\n";


            //m_list.FitToPage = m_cbFitToPage.Checked;
            //	m_list.PageSetup();

            m_list.FitToPage = true;
            m_list.PrintPreview();

        }
        public void FillList(ListView list, ListView table)
        {
            list.SuspendLayout();

            // Clear list
            list.Items.Clear();
            list.Columns.Clear();

            // Columns
            int nCol = 0;
          

            int a = 0;

            int col_cnt = 0;

            if (iButtonMode == false)
            {
                col_cnt = 16;
                try
                {
                    for (int i = 0; i < col_cnt; i++)
                    {


                        ColumnHeader[] col = new ColumnHeader[col_cnt];
                        ColumnHeader ch = new ColumnHeader();

                        col[i] = table.Columns[i];

                        if (i == 3 || i==8 || i==9 || i==14 || i==15)
                        {
                            continue;
                        }
                        ch.Text = col[i].Text;
                        ch.TextAlign = HorizontalAlignment.Right;
                        switch (nCol)
                        {
                            case 0: ch.Width = 40; break;       // id   

                            case 1: ch.Width = 80; break;       // 차량번호

                            case 2: ch.Width = 80; break;       // 기사 번호

                            case 3: ch.Width = 90; break;
                            case 4:
                                ch.TextAlign = HorizontalAlignment.Left;    // 출고 날짜 
                                ch.Width = 100;
                                break;
                            case 5: ch.Width = 90; break;              // 츨고 타임
                            case 6: ch.Width = 90; break;               // 입고 날짜

                            case 7: ch.Width = 90; break;               // 입고 시간

                            case 8: ch.Width = 130; break;               // 

                            case 9: ch.Width = 130; break;

                            case 10: ch.Width = 150; break;

                            case 11: ch.Width = 100; break;
                            case 12: ch.Width = 100; break;
                            case 13: ch.Width = 100; break;
                            case 14: ch.Width = 100; break;
                            case 15: ch.Width = 100; break;
                            //    case 11:
                            //   case 12:
                            //   case 13:
                            //   case 14:
                            //   case 15: ch.Width = 40; break;
                            //	case 16:
                            //	case 17: 
                            default:
                                ch.Width = 0;

                                break;
                        }
                        list.Columns.Add(ch);
                        nCol++;
                    }


                    for (int n = 0; n < table.Items.Count; n++)
                    {
                        ListViewItem item = new ListViewItem();
                        //item.Text = row[0].ToString();

                        item.Text = table.Items[n].Text;

                        for (int i = 1; i < table.Columns.Count; i++)
                        {
                            if (i == 3 || i == 8 || i == 9 || i == 14 || i == 15)
                            {
                                continue;
                            }
                            item.SubItems.Add(table.Items[n].SubItems[i].Text);
                        }
                        list.Height = 100;
                        list.Items.Add(item);
                    }



                    list.ResumeLayout();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                col_cnt = 16;
                try
                {
                    for (int i = 0; i < col_cnt; i++)
                    {


                        ColumnHeader[] col = new ColumnHeader[col_cnt];
                        ColumnHeader ch = new ColumnHeader();

                        col[i] = table.Columns[i];
                        ch.Text = col[i].Text;
                        ch.TextAlign = HorizontalAlignment.Right;
                        switch (nCol)
                        {
                            case 0: ch.Width = 40; break;       // id   

                            case 1: ch.Width = 80; break;       // 차량번호

                            case 2: ch.Width = 80; break;       // 기사 번호

                            case 3: ch.Width = 170; break;
                            case 4:
                                ch.TextAlign = HorizontalAlignment.Left;    // 출고 날짜 
                                ch.Width = 100;
                                break;
                            case 5: ch.Width = 100; break;              // 츨고 타임
                            case 6: ch.Width = 100; break;               //입고 날짜

                            case 7: ch.Width = 100; break;               // 입고 시간

                            case 8: ch.Width = 100; break;               // 

                            case 9: ch.Width = 100; break;

                            case 10: ch.Width = 100; break;

                            case 11: ch.Width = 100; break;
                            case 12: ch.Width = 100; break;
                            case 13: ch.Width = 100; break;
                            case 14: ch.Width = 100; break;
                            case 15: ch.Width = 100; break;
                            //    case 11:
                            //   case 12:
                            //   case 13:
                            //   case 14:
                            //   case 15: ch.Width = 40; break;
                            //	case 16:
                            //	case 17: 
                            default:
                                ch.Width = 0;

                                break;
                        }
                        list.Columns.Add(ch);
                        nCol++;
                    }

                  
                        for (int n = 0; n < table.Items.Count; n++)
                        {
                            ListViewItem item = new ListViewItem();
                            //item.Text = row[0].ToString();

                            item.Text = table.Items[n].Text;

                            for (int i = 1; i < table.Columns.Count; i++)
                            {

                                item.SubItems.Add(table.Items[n].SubItems[i].Text);
                            }
                            list.Height = 100;
                            list.Items.Add(item);
                        }
                  


                    list.ResumeLayout();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }


          
        }

        private void buttonDBDel_Click_1(object sender, EventArgs e)
        {
            if (UserLogin == true || ViewerMode ==true)
            {
                MessageBox.Show("You don't have permission.");
                return;
            }
            if (mdbfilename == "" || listView1.SelectedItems.Count ==0)
            {
                return;
            }

            if (MessageBox.Show("You want to delete the selected data?", "Delete", MessageBoxButtons.YesNoCancel) == DialogResult.Yes)
            {
                List<int> liID = new List<int>();
                int cnt = 0;

                cnt = listView1.SelectedItems.Count;

                string Selmetermoney = "";  // 미터 수입 
                string SalesDistance = ""; // 영업거리
                string TotalDistance = "";
                string tBreaek = "";		// 급제동
                string DriveBasic = "";		//주행기본
                string tDA = "";			//주행이후
                string tAB = "";			//할증기본
                string tAA = "";					// 할증이후

                for (int i = 0; i < listView1.SelectedItems.Count; i++)
                {
                    if (listView1.SelectedItems[i].SubItems[0].Text != "SUM")
                    {
                        liID.Add(Convert.ToInt32(listView1.SelectedItems[i].SubItems[0].Text));


                        Selmetermoney = listView1.SelectedItems[i].SubItems[4].Text;  // 미터 수입 추출
                        SalesDistance = listView1.SelectedItems[i].SubItems[5].Text; // 영업거리
                        //  TotalDistance = listView1.SelectedItems[i].SubItems[9].Text;  // 주행거리
                        //  tBreaek = listView1.SelectedItems[i].SubItems[11].Text;  // 급제동
                        //   DriveBasic = listView1.SelectedItems[i].SubItems[12].Text;  //주행기본
                        //   tDA = listView1.SelectedItems[i].SubItems[13].Text;
                        //   tAB = listView1.SelectedItems[i].SubItems[14].Text;
                        //   tAA = listView1.SelectedItems[i].SubItems[15].Text;

                        Selmetermoney = Selmetermoney.Replace(",", "");
                        Selmetermoney = Selmetermoney.Replace("₩", "");
                        Selmetermoney = Selmetermoney.Replace(@"₩", "");
                        Selmetermoney = Selmetermoney.Replace("\\", "");
                        Selmetermoney = Selmetermoney.Replace(".", "");



                        SalesDistance = SalesDistance.Replace(",", "");
                        SalesDistance = SalesDistance.Replace(".", "");
                        SalesDistance = SalesDistance.Replace("Km", "");
                    }

                  //  TotalDistance = TotalDistance.Replace(",", "");
                 //   TotalDistance = TotalDistance.Replace(".", "");
                  //  TotalDistance = TotalDistance.Replace("Km", "");




                    /*	
                            tB    // 급제동
							
                            tDB   // 주행기본
							
                            tDA   // 주행이후
							
                            tAB   // 할증기본
							
                            tAA ;// 할증이후 */


                  /*  total.tMoney -= Int32.Parse(Selmetermoney);  // 미터 수입 
                    total.tDistS -= (double)Int32.Parse(SalesDistance) / 100;  // 영업거리
                    total.tDistD -= (double)Int32.Parse(TotalDistance) / 100;  // 주행거리
                    total.tB -= Int32.Parse(tBreaek);					// 급제동
                    total.tDB -= Int32.Parse(DriveBasic);             // 주행기본
                    total.tDA -= Int32.Parse(tDA);					//주행이후
                    total.tAB -= Int32.Parse(tAB);					//할증기본
                    total.tAA -= Int32.Parse(tAA);					//할증이후*/
                }


                for (int i = 0; i < listView1.SelectedItems.Count; i++)
                {
                    //	liID.Add(Convert.ToInt32(listView1.SelectedItems[i].SubItems[0].Text));
                    //	listView1.SelectedItems[i].Remove();

                }


                //  liID.Sort();
                //	Thread DellThread = new Thread(DellWork);
                //	DellThread.Start();


                if (DeleteDB(liID))
                {
                    DB_ReadData(0, 1);
                    MessageBox.Show("Successfully deleted!", "Result");

                    if (listView1.Items.Count == 1)
                    {
                        //	MessageBox.Show("listView1.Items[0].Remove();");
                        listView1.Items[0].Remove();
                    }
                }
                //FillData(0, 1);
            }
        }
        public bool DeleteDB(List<int> liID)
        {
            bool bIsDeleteSuccess = false;

            
            
            string NameDB = "";
            string formname = "";
            FileInfo file = new FileInfo(mdbfilename);

            string DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + mdbfilename;

            OleDbConnection conn1 = new OleDbConnection(@DBstring);
            string queryDelTblTacho;
            try
            {
                conn1.Open();


                for (int i = 0; i < liID.Count; i++)
                {
                    queryDelTblTacho = "DELETE * FROM TblTacho WHERE ID=" + liID[i].ToString();
                    OleDbCommand commDelTblTacho = new OleDbCommand(queryDelTblTacho, conn1);
                    commDelTblTacho.ExecuteNonQuery();


                }

                bIsDeleteSuccess = true;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);


                bIsDeleteSuccess = false;
                // Fill Data		
            }
            finally
            {
                conn1.Close();

            }

            return bIsDeleteSuccess;
        }

        private void button9_Click_1(object sender, EventArgs e)
        {



            SerialSend = true;

            if (serialPort1.IsOpen)
            {

                byte[] Data = new byte[2];

                Data[0] = 0x55;
                Data[1] = 0xA6;

                serialPort1.DiscardOutBuffer();
                serialPort1.DiscardInBuffer();
                serialPort1.Write(Data, 0, Data.Length);
            }
        }

        private void 전체삭제AToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (UserLogin == true || ViewerMode ==true)
            {
                MessageBox.Show("You don't have permission.");
                return;
            }
            string DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + mdbfilename;

            OleDbConnection conn1 = new OleDbConnection(@DBstring);


            string NameDB = "";


            if (MessageBox.Show("Delete all?", "Delete", MessageBoxButtons.YesNoCancel) == DialogResult.Yes)
            {
               // workerThread.Start();

                try
                {
                    // DB 전체 데이터 삭제

                    OleDbConnection conn = new OleDbConnection(@DBstring);
                    conn.Open();

                    string queryRead = "select * from TblTacho ORDER BY ID";
                    OleDbCommand commRead = new OleDbCommand(queryRead, conn);
                    OleDbDataReader srRead = commRead.ExecuteReader();

                    List<int> delID = new List<int>();

                    while (srRead.Read())
                    {
                        delID.Add(srRead.GetInt32(0));
                    }

                    conn.Close();

                    if (delID.Count != 0)
                    {
                        OleDbConnection conDB = new OleDbConnection(@DBstring);
                        conDB.Open();

                        foreach (int d in delID)
                        {
                            string queryDelTblTacho = "DELETE FROM TblTacho where ID = " + d.ToString();
                            OleDbCommand comDelDelTblTacho = new OleDbCommand(queryDelTblTacho, conDB);
                            comDelDelTblTacho.ExecuteNonQuery();


                        }

                        conDB.Close();
                        DB_ReadData(0, 1);
                       
                        MessageBox.Show("Sucessfully Remove All!", "Result");
                    }
                    else
                    {
                        MessageBox.Show("Empty Data!", "결과");
                    }
                }
                catch (Exception ex)
                {
                    /*  string path = Application.StartupPath + "\\ErrorLog.jie";
                      using (StreamWriter sw = new StreamWriter(path, true, System.Text.Encoding.UTF8))
                      {
                          sw.WriteLine("[" + DateTime.Now.ToString() + "] " + ex.Message);
                      }*/
                }
                finally
                {
                    total.Distance = 0;
                    total.Money = 0;
                 


                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
           // Repaet_btn = true;
            if (ViewerMode == true)
            {
                MessageBox.Show("You don't have permission.");
                return;
            }
            Thread LoadingThread = new Thread(LoadingWork);
            LoadingThread.Start();
        }

        public void LoadingWork()    // LoadThread Funtion
        {
            bool _shouldStop;

            //	tmf_data(strFname);

            Repeat_Data();
            _shouldStop = true;


        }
        public T[] GEtDistinctValues<T>(T[] array)   // 배열 중복 검사 제거
        {
            List<T> tmp = new List<T>();

            for (int i = 0; i < array.Length; i++)
            {
                if (tmp.Contains(array[i]))
                    continue;
                tmp.Add(array[i]);
            }
            return tmp.ToArray();
        }
        public void Repeat_Data()
        {


           // Worker workerObject = new Worker();
           //// Thread workerThread = new Thread(workerObject.DoWork);
            //workerThread.Start();
           // OleDbConnection conn;
            Color AliceBlue = Color.FromArgb(240, 248, 255);
            Color LightPink = Color.FromArgb(255, 182, 193);

            int cnt = 0;
            string NameDB = "";
            //	SetAlternatingRowColors(listView1, a, b);

            try
            {

                //	selectedColumnIndex = 4;   // 출고 시간별 정렬
                //	selectedOrder = 1;
                //	FillData(selectedColumnIndex, selectedOrder);

               

               
                FileInfo file = new FileInfo(mdbfilename);

                string DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + mdbfilename;

                OleDbConnection conn = new OleDbConnection(@DBstring);
                string queryDelTblTacho;

                // Fill Data

              
                conn.Open();

                string queryTblTacho = "SELECT * FROM TblTacho ORDER BY ID DESC";
                OleDbCommand commTblTacho = new OleDbCommand(queryTblTacho, conn);
                OleDbDataReader srRead = commTblTacho.ExecuteReader();


                cnt = listView1.Items.Count - 1;
                string[] dbId = new string[cnt];
                DateTime[] dbOuttime = new DateTime[cnt];
                DateTime[] dbIntime = new DateTime[cnt];
                string[] dbCarnumber = new string[cnt];

                int[] intdbId = new int[cnt * 3];
                int[] RepeatId = new int[cnt * 3];
                bool RepeatCheck = false;
                int ii = 0;


                while (srRead.Read())   // db 데이터 읽어 오기 아이디 , 차량 번호 ,출고 시간 ,입고시간
                {
                    intdbId[ii] = srRead.GetInt32(0);
                    dbCarnumber[ii] = srRead.GetString(1);
                    dbOuttime[ii] = srRead.GetDateTime(4);
                    dbIntime[ii] = srRead.GetDateTime(5);
                    ii++;

                }

                int k = 0;
                int[] temp = new int[cnt * 3];
                bool pass = false;


                for (int i = 0; i < cnt; i++)
                {


                    for (int j = 0; j < cnt; j++)
                    {

                        for (int u = 0; u < temp.Length; u++)  // 지운 아이디의 인덱스 루틴패스시킴
                        {
                            if (i != 0)
                            {
                                if (i == temp[u])
                                    pass = true;
                            }

                        }


                        if (i == j || pass == true)  // 자기 자신 건너뛰고 지운데이터 인덱스 건너 뛰기
                        {
                            pass = false;
                            continue;

                        }


                        if (dbCarnumber[i] == dbCarnumber[j] && dbOuttime[i] == dbOuttime[j])  // 차번호 입고 출고 시간 중복 검사
                        {
                            RepeatId[k] = intdbId[j];
                            temp[k] = j;
                            RepeatCheck = true;


                            k++;

                        }

                    }

                }
                RepeatId = GEtDistinctValues<int>(RepeatId);  // 배열 중복 제거 하기

                //	string aA = listView1.Items[0].SubItems[0].Text;  // 리스트 뷰 id를 텍스트로 얻어오기	

                int[] listviewindex = new int[cnt];

                // 현재의 데이터를 리스트뷰를 읽어 파확한다	
                if (cnt != -2)
                {

                    for (int i = 0; i < RepeatId.Length; i++)  // 중복데이터 라인 색 칠하기
                    {

                        for (int j = 0; j < cnt; j++)
                        {
                            dbId[j] = listView1.Items[j].SubItems[0].Text;   // listview Id -> string
                            intdbId[j] = Int32.Parse(dbId[j]);              //  listview id -> int

                            if (RepeatId[i] == intdbId[j])
                            {

                                listView1.Items[j].BackColor = LightPink;

                            }


                        }

                    }
                    //	workerObject.RequestStop();

                    if (RepeatCheck == true)
                    {
                        string datacnt = (RepeatId.Length - 1).ToString();
                        DialogResult result = MessageBox.Show(datacnt + " data was duplicated." + " \nDiscard data?", "Replicated Data", MessageBoxButtons.YesNoCancel);
                        //	DialogResult result = MessageBox.Show("중복된 데이터를 삭제 하시겠습니까?", "중복 체크", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            for (int i = 0; i < RepeatId.Length; i++)
                            {



                                 queryDelTblTacho = "DELETE FROM TblTacho where ID = " + RepeatId[i].ToString();
                                OleDbCommand comDelDelTblTacho = new OleDbCommand(queryDelTblTacho, conn);
                                comDelDelTblTacho.ExecuteNonQuery();



                            }

                            selectedColumnIndex = 1;   // 출고 시간별 정렬
                            selectedOrder = 1;
                            DB_ReadData(0, 1);
                            //FillData(selectedColumnIndex, selectedOrder);

                        }
                    }
                    else
                    {
                        
                    }


                }
                conn.Close();


            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
              
             //   conn.Close();
            }
            finally
            {
               // workerObject.RequestStop();
              //  conn.Close();



                //		selectedColumnIndex = 4;   // 출고 시간별 정렬
                //		selectedOrder = 1;
                //		FillData(selectedColumnIndex, selectedOrder);


                
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            testid++;
            string strtemp = "";
            strtemp = "A" + testid.ToString();         // id

            strtemp += "B" + "차량번호";         // CarNumber

            DateTime Outtime = new DateTime(2014, 1, 1, 1, 1, 1);
            strtemp += "C" + Outtime.ToString("yyyy-MM-dd tt HH:mm");         // OutTime



            DateTime Intime = new DateTime(2014, 1, 1, 2, 2, 2);
            strtemp += "D" + Intime.ToString("yyyy-MM-dd tt HH:mm");         // InTime

            double TotalDistance = 100;
            strtemp += "E" + string.Format("{0:N} Km", TotalDistance);         // 주행거리

            double TotalSalesDist = 1000;
            strtemp += "F" + string.Format("{0:N} Km", TotalSalesDist);         // 영업거리

            int TodayIncomeMoney = 220000;
            strtemp += "G" + string.Format("{0:F}", TodayIncomeMoney);         //미터 수입

            strtemp += "H";


          /*  if (textBox1.Text == "") return;

          //  byteSendMsg = Encoding.Default.GetBytes(textBox1.Text);
            strtemp = string.Format("[CLIENT] : {0}", textBox1.Text);
        
         */

            strtemp = textBox1.Text;
        
            svr.TachoSend1 = true;
            svr.TachoStr = strtemp;
          //  svr.TachoStr = textBox1.Text;
          
        }
     
        public void Make_AMf_File(byte[] mStreamBuffer)
        {
            try
            {
                int datacnt = 0;
                // bool fdcheck = false;

                for (int i = 0; i < mStreamBuffer.Length; i++)
                {
                   
                    
                        if (mStreamBuffer[i] == 0xfd && i > 256)
                        {
                            //   fdcheck = true;
                            datacnt++;
                            break;
                        }
                        else
                        {
                            datacnt++;
                        }
                    
                }
                datacnt++;     // CheckSum index!!


              
                

                byte[] Amf_array = new byte[datacnt];

                for (int i = 0; i < datacnt; i++)
                {
                    Amf_array[i] = mStreamBuffer[i];
                }
                

             /*   byte[] Amf_array = new byte[mStreamBuffer.Length];

                for (int i = 0; i < mStreamBuffer.Length; i++)
                {
                    Amf_array[i] = mStreamBuffer[i];
                }*/


                //////////////////////////////// 한개 짜리 tmf 저장 /////////////////////////////////////////

                string NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}{3:D2}{4:D2}{5:D2}",
                                                           (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                string carnum = "";
                string Model = "";
                if (Amf_array.Length > 256 + 64)
                {
                    for (int i = 0; i < 9; i++)
                    {
                        if (Amf_array[180 + i] < 0x20)
                        {
                            Amf_array[180 + i] = 0x20;
                        }
                    }
                    carnum = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}{5:C2}{6:C2}{7:C2}{8:C2}"
                                             , Convert.ToChar(Amf_array[185]), Convert.ToChar(Amf_array[186]), Convert.ToChar(Amf_array[187]), Convert.ToChar(Amf_array[188])
                                             , Convert.ToChar(Amf_array[183]), Convert.ToChar(Amf_array[184])
                                          , Convert.ToChar(Amf_array[180]), Convert.ToChar(Amf_array[181]), Convert.ToChar(Amf_array[182]));

                    Model = String.Format("{0:C2}{1:C2}{2:C2}{3:C2}{4:C2}"
                                          , Convert.ToChar(Amf_array[240]), Convert.ToChar(Amf_array[241]), Convert.ToChar(Amf_array[242]), Convert.ToChar(Amf_array[243]), Convert.ToChar(Amf_array[244]));
                    //  strMN += String.Format("{0:C}", Convert.ToChar(newProInHeader.ModelName[i]));
                }

                //  string TmpFile = Application.StartupPath + "\\" + NowReceiveTime + ".TMF";


                string newPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "AMF");
                // Create the subfolder
                System.IO.Directory.CreateDirectory(newPath);

                string TmpFile = TACHO2_path + "\\TMF\\AMF\\" + Model + "_" + NowReceiveTime + "_" + carnum + ".AMF";

                Amf_path = TmpFile;

                // byte[] rcvByte = new byte[mStreamBuffer.Length];
                // rcvList.CopyTo(rcvByte);

                FileStream fs = new FileStream(TmpFile, FileMode.OpenOrCreate, FileAccess.Write);
                BinaryWriter bw = new BinaryWriter(fs);

                bw.Write(Amf_array);



                fs.Close();
                bw.Close();


                if (svr != null)
                {

                    svr.TachoSend1 = true;
                }

             //   AMF_Data(Amf_path);

                //////////////////////////////////////   2번째 저장 total tmf  기존 데이터와  이어 붙히자 !!///////////////////

                string TMFPath = System.IO.Path.Combine(TACHO2_path + "\\TMF", "TransData");
                // Create the subfolder
                System.IO.Directory.CreateDirectory(TMFPath);

                NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                        (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day);
                TmpFile = TACHO2_path + "\\TMF\\TransData\\" + NowReceiveTime + ".AMF";

                // rcvList.RemoveAt(0);
                //  rcvByte = new byte[rcvList.Count];
                //  rcvList.CopyTo(rcvByte);

                fs = new FileStream(TmpFile, FileMode.Append, FileAccess.Write);
                bw = new BinaryWriter(fs);

                bw.Write(Amf_array);
                fs.Close();
                bw.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            AMF_Data(Amf_path);

              //  svr.TachoSend1 = true;
              //  svr.TachoStr = "Good!";


            
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Open File";
            openFileDialog1.Filter = "*.amf|*.AMF";
            string AMf_File = "";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

              AMf_File  = openFileDialog1.FileName;

              AMF_Data(AMf_File);
            }
        }
      
        private void listView1_SubItemClicked(object sender, ListViewEx.SubItemEventArgs e)
        {
            try
            {
               // if (e.SubItem == 0 || e.SubItem == 1 || e.SubItem == 2 || e.SubItem == 3 || e.SubItem == 4 || e.SubItem == 5 || e.SubItem == 6 || e.SubItem == 7)
                if(e.SubItem >=0 && e.SubItem  <=15)
                {
                    if (listView1.SelectedItems.Count == 0)
                    {
                       
                        return;
                    }
                    ListView.SelectedListViewItemCollection items = listView1.SelectedItems;

                    foreach (ListViewItem item in items)
                    {
                        if ((item.SubItems[0].Text.CompareTo("SUM") == 0))
                        {
                            nOpenedindex = 0;
                            return;
                        }
                        else
                        {
                            nOpenedindex = Convert.ToInt32(item.SubItems[0].Text);
                        }
                    }
                   
                        TransactionForm transactionform = new TransactionForm(this);
                        transactionform.Read_Transaction(nOpenedindex);
                        transactionform.MdiParent = this.ParentForm;

                        //  transactionform.MdiParent = this;
                        transactionform.BringToFront();
                        LayoutMdi(MdiLayout.TileHorizontal);
                        if (TransactionForm_Run == false)
                        {
                            transactionform.Show();
                        }
                        else
                        {
                            TransactionForm_Run = false;
                        }

                        transactionform.Focus();
                 
                  
                  /*  TransactionForm  transactionform = new TransactionForm(this);
                    string formName = "TransactionForm";

                    foreach (System.Windows.Forms.Form theForm in this.MdiChildren)
                    {
                        if (formName.Equals(theForm.Name))
                        {
                            //해당form의 인스턴스가 존재하면 해당 창을 활성시킨다.
                            theForm.BringToFront();
                            theForm.Focus();
                            return;
                        }
                    }
                 
                    transactionform.MdiParent = this;
                    transactionform.Show();
                    LayoutMdi(MdiLayout.TileHorizontal);*/
                }

            }
            catch (Exception excep)
            {
                MessageBox.Show(excep.ToString());
            }
        }

        private void 끝내기XToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (UAE_Set == true)
            {
                System.Diagnostics.Process.Start("www.luxumborj.com");
             //   System.Diagnostics.Process.Start("www.taximeter.net");
            }
            else
            {
                System.Diagnostics.Process.Start("www.taximeter.net");
               
            }
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {

                byte[] Data = new byte[2];

                Data[0] = 0x55;
                Data[1] = 0x15;

                serialPort1.DiscardOutBuffer();
                serialPort1.DiscardInBuffer();
                serialPort1.Write(Data, 0, Data.Length);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            IbuttonForm ibutton = new IbuttonForm(this);

            
            if (ibutonCheck == false)
            {
                IbuttonSetting = true;
                ibutton.ShowDialog();
            }
            else
            {
                ibutonCheck = false;
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            SearchForm searchform = new SearchForm(this);
            searchform.ShowDialog();
        }

        private void listView1_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if (e.Item.SubItems[0].Text != "SUM")
            {

                Taxiid_textBox.Text = e.Item.SubItems[1].Text;
                Driverid_textBox.Text = e.Item.SubItems[2].Text;
                Drivername_textBox.Text = e.Item.SubItems[3].Text;
                SelectItem = Int32.Parse(e.Item.SubItems[0].Text);
                textBox2.Text = e.Item.SubItems[10].Text;
                textBox3.Text = e.Item.SubItems[10].Text;
            }
        }
        /*
        private void button16_Click(object sender, EventArgs e)
        {
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////

            if (label32.Text =="")
            {
                MessageBox.Show("Disconnected  to ibutton.");
                return;
            }
            // 서버에 접속하여 Paid Time을 설정해야함
            string CarNum = Taxiid_textBox.Text;
            string DriverID = Driverid_textBox.Text;
            string DriverName = Drivername_textBox.Text;
            byte[] CarByte = new byte[9];
            byte[] DriverByte = new byte[9];
            byte[] DriverNameByte = new byte[32];
          
            /////////////////////////////////////////////////////////////////
            short[] ROM = new short[8];
            //      byte[] block = new byte[322];
            byte[] TachoAddr = new byte[6]; // 타코 시작 주소와 끝 주소를 찾는다. 
            Int16 StartAddr = 0;
            Int16 EndAddr = 0;
            int n = 0;
            byte[] Address = new byte[2];
            byte[] TachoStartByte = new byte[1];
            byte[] TachoOverFlow = new byte[1];
            short ret = 0;
            short portNum = 0;
            short portType = 0;
            int hSess = 0;
            bool errorstatus = false;



            TMEX TMEXLibrary = new TMEX();
            TMEXLibrary.TMReadDefaultPort(ref portNum, ref portType);

            hSess = TMEXLibrary.TMExtendedStartSession(portNum, portType, ref sessionOptions);

           


          
            try
            {
                if (bConnectServer)
                {
                    //  strLog = string.Format("[SYSTEM] : 현재 서버와 접속중입니다.!!");
                    //    Add_Log(strLog);
                    //  return;
                }
              
                Client_label.Visible = true;
                Client_label.Text = "Connecting to the server ...";
                Loading_Visible(true);
                EY_ChatClient = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.IP);
                IPAddress ipServer = IPAddress.Parse(ServerIP);

                //  while (true)
                // {
                try
                {
                    EY_ChatClient.Connect(new IPEndPoint(ipServer, intPortNum));


                    if (EY_ChatClient.Connected)
                    {
                        // strLog = string.Format("[SYSTEM] : 서버 접속 성공!!");
                        //  Add_Log(strLog);
                        bConnectServer = true;

                        //  pictureBox1.Visible = true;

                        //  label7.Text = "서버 접속중";
                        byte CheckSum = 0;

                        byteSendMsg[0] = 0x55;
                        byteSendMsg[1] = 0x15;

                        if (Taxiid_textBox.Text == "")
                        {

                        }
                        else
                        {
                            int cnt = 8;
                            for (int i = 0; i < CarNum.Length; i++)
                            {
                                CarByte[cnt] = (byte)CarNum[i];
                                cnt--;
                            }

                        }

                        byteSendMsg[2] = CarByte[0];  // Taxi Id 
                        byteSendMsg[3] = CarByte[1];
                        byteSendMsg[4] = CarByte[2];
                        byteSendMsg[5] = CarByte[3];
                        byteSendMsg[6] = CarByte[4];
                        byteSendMsg[7] = CarByte[5];
                        byteSendMsg[8] = CarByte[6];
                        byteSendMsg[9] = CarByte[7];
                        byteSendMsg[10] = CarByte[8];



                        for (int i = 0; i < 9; i++)
                        {
                            DriverByte[i] = 0x20;
                        }
                        if (Driverid_textBox.Text == "")
                        {

                        }
                        else
                        {
                            int cnt = 8;
                            for (int i = 0; i < DriverID.Length; i++)
                            {
                                DriverByte[cnt] = (byte)DriverID[i];
                                cnt--;
                            }



                        }
                        byteSendMsg[11] = DriverByte[0];  // Driver ID
                        byteSendMsg[12] = DriverByte[1];
                        byteSendMsg[13] = DriverByte[2];
                        byteSendMsg[14] = DriverByte[3];
                        byteSendMsg[15] = DriverByte[4];
                        byteSendMsg[16] = DriverByte[5];
                        byteSendMsg[17] = DriverByte[6];
                        byteSendMsg[18] = DriverByte[7];
                        byteSendMsg[19] = DriverByte[8];

                        for (int a = 0; a < 19; a++)
                        {
                            CheckSum += byteSendMsg[a];
                        }
                        byteSendMsg[20] = CheckSum;
                        byteSendMsg[21] = 0xFD;


                        // 접속 메세지 전송
                        EY_ChatClient.BeginSend(byteSendMsg, 0, 22, SocketFlags.None, new AsyncCallback(CallBack_SendMsg), EY_ChatClient);
                        //이제 메시지를 전송받는다.
                        EY_ChatClient.BeginReceive(byteReceiveMsg, 0, byteReceiveMsg.Length, SocketFlags.None, new AsyncCallback(CallBack_ReceiveMsg), byteReceiveMsg);
                        //ReceiveStart();
                        // break;


                        Client_label.Text = "Complete access to the server !";
                     
                      
                    }
                }
                catch (Exception err)
                {
                    // strErr = string.Format("[SYSTEM] : {0}", err.Message);
                    //   Add_Log(strErr);
                    ///////////////////////////////////////////////////////////////
                    pictureBox17.Visible = false;
                    Client_label.Text = "Server connection fails!!";
                    MessageBox.Show(err.Message);
                    bConnectServer = false;
                    Client_label.Visible = false;
                }
                //  }

            }
            catch (Exception ex)
            {

            }
           
        }
        */
        private void button16_Click_1(object sender, EventArgs e)
        {
          //  GetTime_Receipt  // 1. 현재 아이버튼의 영업 start year month를 읽어와야 한다. 2. 시작 날짜를 생성한다. 
                               // 3. ini읽은후 시작 날짜부터 한달 계산 하여 Enddate를 생성한다.

         
                                                   
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // 네트워크 오류 일때 처리가 필요함 !!



            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
            if (label32.Text == "")
            {
                MessageBox.Show("Disconnected  to ibutton.");
                return;
            }
            Taxi_ID_temp = Taxi_ID_temp.Trim();
            Taxi_ID_temp = Taxi_ID_temp.PadRight(9);

            Taxiid_textBox.Text = Taxiid_textBox.Text.Trim();
            Taxiid_textBox.Text = Taxiid_textBox.Text.PadRight(9);

            if (Taxi_ID_temp != Taxiid_textBox.Text)
            {
                MessageBox.Show("The iButton information entered is incorrect.");
                return;
            }

            // 서버에 접속하여 Paid Time을 설정해야함
            string CarNum = Taxiid_textBox.Text;
            string DriverID = Driverid_textBox.Text;
            string DriverName = Drivername_textBox.Text;
            byte[] CarByte = new byte[9];
            byte[] DriverByte = new byte[9];
            byte[] DriverNameByte = new byte[32];

            /////////////////////////////////////////////////////////////////
            short[] ROM = new short[8];
            //      byte[] block = new byte[322];
            byte[] TachoAddr = new byte[6]; // 타코 시작 주소와 끝 주소를 찾는다. 
            Int16 StartAddr = 0;
            Int16 EndAddr = 0;
            int n = 0;
            byte[] Address = new byte[2];
            byte[] TachoStartByte = new byte[1];
            byte[] TachoOverFlow = new byte[1];
            short ret = 0;
            short portNum = 0;
            short portType = 0;
            int hSess = 0;
            bool errorstatus = false;


           
         //   TMEX TMEXLibrary = new TMEX();
          //  TMEXLibrary.TMReadDefaultPort(ref portNum, ref portType);

         //   hSess = TMEXLibrary.TMExtendedStartSession(portNum, portType, ref sessionOptions);

            if (!NetworkInterface.GetIsNetworkAvailable())
            {
                 // 네트워크 오류 상태에서 Collect 요청일때....


            }
            else
            {


                try
                {
                    if (bConnectServer)
                    {
                        //  strLog = string.Format("[SYSTEM] : 현재 서버와 접속중입니다.!!");
                        //    Add_Log(strLog);
                        //  return;
                    }

                    Client_label.Visible = true;
                    Client_label.Text = "Connecting to the server ...";
                    Loading_Visible(true);
                    EY_ChatClient = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.IP);
                    IPAddress ipServer = IPAddress.Parse(ServerIP);

                    //  while (true)
                    // {
                    try
                    {
                        EY_ChatClient.Connect(new IPEndPoint(ipServer, intPortNum));


                        if (EY_ChatClient.Connected)
                        {
                            // strLog = string.Format("[SYSTEM] : 서버 접속 성공!!");
                            //  Add_Log(strLog);
                            bConnectServer = true;

                            //  pictureBox1.Visible = true;

                            //  label7.Text = "서버 접속중";
                            byte CheckSum = 0;

                            byteSendMsg[0] = 0x55;
                            byteSendMsg[1] = 0x15;

                            if (Taxiid_textBox.Text == "")
                            {

                            }
                            else
                            {
                                int cnt = 8;
                                for (int i = 0; i < CarNum.Length; i++)
                                {
                                    CarByte[cnt] = (byte)CarNum[i];
                                    cnt--;
                                }

                            }

                            byteSendMsg[2] = CarByte[0];  // Taxi Id 
                            byteSendMsg[3] = CarByte[1];
                            byteSendMsg[4] = CarByte[2];
                            byteSendMsg[5] = CarByte[3];
                            byteSendMsg[6] = CarByte[4];
                            byteSendMsg[7] = CarByte[5];
                            byteSendMsg[8] = CarByte[6];
                            byteSendMsg[9] = CarByte[7];
                            byteSendMsg[10] = CarByte[8];



                            for (int i = 0; i < 9; i++)
                            {
                                DriverByte[i] = 0x20;
                            }
                            if (Driverid_textBox.Text == "")
                            {

                            }
                            else
                            {
                                int cnt = 8;
                                for (int i = 0; i < DriverID.Length; i++)
                                {
                                    DriverByte[cnt] = (byte)DriverID[i];
                                    cnt--;
                                }



                            }
                            byteSendMsg[11] = DriverByte[0];  // Driver ID
                            byteSendMsg[12] = DriverByte[1];
                            byteSendMsg[13] = DriverByte[2];
                            byteSendMsg[14] = DriverByte[3];
                            byteSendMsg[15] = DriverByte[4];
                            byteSendMsg[16] = DriverByte[5];
                            byteSendMsg[17] = DriverByte[6];
                            byteSendMsg[18] = DriverByte[7];
                            byteSendMsg[19] = DriverByte[8];

                            for (int a = 0; a < 19; a++)
                            {
                                CheckSum += byteSendMsg[a];
                            }
                            byteSendMsg[20] = CheckSum;
                            byteSendMsg[21] = 0xFD;


                            // 접속 메세지 전송
                            EY_ChatClient.BeginSend(byteSendMsg, 0, 22, SocketFlags.None, new AsyncCallback(CallBack_SendMsg), EY_ChatClient);
                            //이제 메시지를 전송받는다.
                            EY_ChatClient.BeginReceive(byteReceiveMsg, 0, byteReceiveMsg.Length, SocketFlags.None, new AsyncCallback(CallBack_ReceiveMsg), byteReceiveMsg);
                            //ReceiveStart();
                            // break;


                            Client_label.Text = "Complete access to the server !";

                              


                        }
                    }
                    catch (Exception err)
                    {
                        // strErr = string.Format("[SYSTEM] : {0}", err.Message);
                        //   Add_Log(strErr);
                        ///////////////////////////////////////////////////////////////
                        pictureBox17.Visible = false;
                        Client_label.Text = "Server connection fails!!";
                        MessageBox.Show(err.Message);
                        bConnectServer = false;
                        Client_label.Visible = false;
                        GetTime_Receipt = false;
                    }
                    //  }

                }
                catch (Exception ex)
                {
                    GetTime_Receipt = false;
                }
            }



            ////////////////////////////////////////////////////////////////////////////////////////////////////////////
        }

        private void button17_Click(object sender, EventArgs e)                       
        {
            this.Visible = false;
            LoginOK = false;
             Login = new LoginForm(this);

            DialogResult dialog = Login.ShowDialog();



        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            panelMenu.Height = listView1.Height;
            treeView1.Height = listView1.Height;
        }

        private void userRegistrationUToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(UserLogin == true || CashierMode ==true)
             {
                 MessageBox.Show("You don't have permission.");
                 return;
             }
            UserIdCreate = true;
            LoginRegistrationForm login = new LoginRegistrationForm(this);

            DialogResult dialog = login.ShowDialog();
        }

        private void administratorRegistrationAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (UserLogin == true || CashierMode == true)
            {
                MessageBox.Show("You don't have permission.");
                return;
            }
            LoginRegistrationForm login = new LoginRegistrationForm(this);

            DialogResult dialog = login.ShowDialog();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
            richTextBox1.Visible = false;
            button5.Visible = false;
            button16.Enabled = true;
         
        }
        int result = 0;

        public void ShareRun()
        {
            try
            {


              
                Loading_Visible(true);

                Treeview_Refresh();

                string path = "";
                if (ViewerMode == true)
                {
                    path = @"\\" + ShareIP + "\\tacho2\\";
                }

                //   path = TACHO2_path + e.Node.Parent.Parent.Parent.Text + "\\" + e.Node.Parent.Parent.Text + "\\" + e.Node.Parent.Text;


                mdbfilename = path + treeView1.Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes[0].Text + ".mdb";

                DB_ReadData(0, 1);  
           
                ///////////////////////////////////////////////////////////////////////////////////
            /*    Loading_Visible(true);
                int capacity = 64;
                uint resultFlags = 0;
                uint flags = 0;
                System.Text.StringBuilder sb = new System.Text.StringBuilder(capacity);
                NETRESOURCE ns = new NETRESOURCE();
                ns.dwType = 1;           // 공유 디스크
                ns.lpLocalName = null;   // 로컬 드라이브 지정하지 않음
                ns.lpRemoteName = @"\\" + ShareIP + "\\tacho2";
                ns.lpProvider = null;

             //   result = WNetUseConnection(IntPtr.Zero, ref ns, "", "Sec", flags,
             //                                  sb, ref capacity, out resultFlags);

                if (System.IO.File.Exists(@"\\" +ShareIP+"\\tacho2\\Information.mdb"))  // 같은 파일의이름이 존재 함
                {
                    int Month = DateTime.Now.Month;
                    int Year = DateTime.Now.Year;

                    // 현재 월에 있는 데이터만 복사한다. 
                    CopyFolder(@"\\" + ShareIP + "\\tacho2\\TACHO\\"  + Year.ToString() + "\\" + Month.ToString(), "c:\\tacho2\\TACHO\\" + Year.ToString() + "\\" + Month.ToString());


               //     CopyFolder(@"\\" + ShareIP + "\\tacho2\\TACHO\\" , "c:\\tacho2\\TACHO\\");
                    Treeview_Refresh();

                    string path = TACHO2_path;

                    mdbfilename = path + treeView1.Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Nodes[0].Text+"\\" + treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes[0].Text + ".mdb";

                    DB_ReadData(0, 1);
                    //  DB_ReadData(0, 1);
                    // string OrgFile = TACHO2_path + "타코\\" + yymmdd + ".mdb";  // tmf 경로 

                    //    string SekiFile =  "c:\\" + 1234 + ".mdb";  // tmf 경로 
                    // System.IO.File.Copy(@"\\192.168.0.34\tacho2\Information.mdb", SekiFile, true);  // 덮어 쓰기 설정 true !!!!!
                }
                else
                {
                    result = 0;

                }*/
                /////////////////////////////////////////////////////////////////////////////////
            }
            catch (Exception ex)
            {
                Loading_Visible(false);
            //    MessageBox.Show(ex.Message);
            }
            Loading_Visible(false);
            Sharethread.Abort();
            Thread.Sleep(1000);

           
          //  DB_ReadData(0, 1);
        }
        public void Driver_Receipt()
        {

            if (Money_Check == 0)
            {
                return;
            }
           // GetPrinterModelName();
            byte[] cmdData = new byte[255];
            InitPrinterStatus();
            SetCharCodeTable(255, 0);
          
            // TODO: Add your control notification handler code here


            //    DateTime NowTime = new DateTime(DateTime.Now.Year ,DateTime.Now.Month,DateTime.Now.Day,DateTime.Now.Hour,DateTime.Now.Minute,DateTime.Now.Second);

            //     PrintSpoolForTTF( "Dubai Taxi\r\n",1,1);

            
            SetTextStyle(2, true, 0, 1, false);
            SetCharSpace(2);


          //  Print_Text

          //  TextSaveSpool("Dubai Taxi\r\n\r\n");
            TextSaveSpool(Print_Text+"\r\n\r\n");
           
            string strDate = DateTime.Now.ToString("dd-MMM-yyyy\r\n", new CultureInfo("en-US"));
            string strTime = DateTime.Now.ToString("HH:mm:ss\r\n", new CultureInfo("en-US"));
            InitPrinterStatus();
        //    SetLineSpace(35);
            TextSaveSpool("Date                 :        " + strDate);
            //    TextSaveSpool("Time                  :          08:30:06\r\n");
            TextSaveSpool("Time                 :           " + strTime);


            SetTextStyle(2, true, 0, 0, false);
            SetCharSpace(2);
            Driver_ID_temp = Driver_ID_temp.Trim();
            Driver_ID_temp = Driver_ID_temp.PadLeft(9);
            TextSaveSpool("Driver Number    :       " + Driver_ID_temp + "\r\n");

            InitPrinterStatus();
            SetLineSpace(35);
            Driver_Name_temp = Driver_Name_temp.Trim();
            Driver_Name_temp = Driver_Name_temp.PadLeft(18);
            TextSaveSpool("Driver Name          : " + Driver_Name_temp + "\r\n");        

            SetTextStyle(2, true, 0, 0, false);
            SetCharSpace(2);
            //  string str2 = str1.PadLeft(10);   // -->  str2 = "      asdf"
            Taxi_ID_temp = Taxi_ID_temp.Trim();
            Taxi_ID_temp = Taxi_ID_temp.PadLeft(9);

            TextSaveSpool("Car Number       :       " + Taxi_ID_temp + "\r\n");
            Thread.Sleep(1000);
            InitPrinterStatus();
            SetLineSpace(35);
            Vacant_dist_temp = Vacant_dist_temp.Trim();
            Vacant_dist_temp = Vacant_dist_temp.PadLeft(6);
            TextSaveSpool("Vacant KM            :             "+Vacant_dist_temp+"\r\n");

            Hired_dist_temp = Hired_dist_temp.Trim();
            Hired_dist_temp = Hired_dist_temp.PadLeft(6);
            TextSaveSpool("Hired KM             :             "+Hired_dist_temp+"\r\n");

            SetTextStyle(2, true, 0, 0, false);
            SetCharSpace(2);
            Total_dist_temp = Total_dist_temp.Trim();
            Total_dist_temp = Total_dist_temp.PadLeft(6);
            TextSaveSpool("Total KM         :          "+ Total_dist_temp+"\r\n");
            InitPrinterStatus();
            SetLineSpace(35);
            string toll = "AED " + Toll_temp;
            toll = toll.PadLeft(12);
            TextSaveSpool("Toll Fare            :       " + toll + "\r\n");
            string ap = "AED " + AP_temp;
            ap = ap.PadLeft(12);
            TextSaveSpool("A/P Fare             :       " + ap + "\r\n");
            ///////////////////////////////////////////////////////////////////////////
            PrintSpool(true);

            Thread.Sleep(2000);

            ClearSpool();

            TextSaveSpool("Extra Fare           :           AED 0.00\r\n");

            SetTextStyle(2, true, 0, 0, false);
            SetCharSpace(2);
            Total_income_temp = Total_income_temp.Trim();
          //  Total_income_temp = Total_income_temp.PadLeft(9);
            string income = "AED " + Total_income_temp;
            income = income.PadLeft(12);
            TextSaveSpool("Shift meter      :    " + income + "\r\n");
            InitPrinterStatus();
            SetLineSpace(35);
          //  TextSaveSpool("Remarks              :                   \r\n");
            CashierID = CashierID.PadLeft(18);
            TextSaveSpool("Received by          : "+CashierID+"\r\n");

            SetTextStyle(2, true, 0, 0, false);
            SetCharSpace(2);


           // MonthlySearchData(Taxi_ID_temp, Driver_ID_temp,DateTime.Now.Year, DateTime.Now.Month); 

          //  DateTime StartDate = new DateTime(2016, 11, 11, 0, 0, 0); 
          //  DateTime EndDate = new DateTime(2016, 11, 11, 0, 0, 0); 

            MonthlySearchData_dubai(Taxi_ID_temp, Driver_ID_temp, StartDate, EndDate);
           
            string Monthlyincome = "AED " + Monthly_temp;
            Monthlyincome = Monthlyincome.PadLeft(12);
            TextSaveSpool("Monthly Total    :    "+Monthlyincome+"\r\n\r\n");
            InitPrinterStatus();
            TextSaveSpool("* Thank you !\r\n");
          
            TextSaveSpool("\r\n\r\n\r\n");

            //Reg Number
                      
            PrintSpool(true);
        

        }

        public void MonthlySearchData_dubai(string CarNO, string DirverID, DateTime StartDate, DateTime EndDate)
        {
            DateTime Starttime = new DateTime(StartDate.Year, StartDate.Month, StartDate.Day, 0, 0, 0);
            DateTime Endtime = new DateTime(EndDate.Year, EndDate.Month, EndDate.Day, 0, 0, 0);
            string Dirname = "";
            int nCnt = 1;
            total.Money = 0;
            total.Distance = 0;

            string year = Starttime.Year.ToString();
            string StartMonth = Starttime.Month.ToString();
            string StartDay = Starttime.Day.ToString();

            string EndMonth = Endtime.Month.ToString();
            string EndDay = Endtime.Day.ToString();
            DateTime mdbTime = new DateTime();
          
            double GtandTotlaMoney = 0;
            double GrandTotalRealIncome = 0;

            if (ViewerMode == true)
            {
                Dirname = @"\\" + ShareIP + "\\tacho2\\" + "\\TACHO\\" + year;  // 해당 연도 폴더를 지정한다. 
            }
            else
            {
                Dirname = TACHO2_path + "\\TACHO\\" + year;  // 해당 연도 폴더를 지정한다. 
            }
            //   Dirname = form1.TACHO2_path + "\\TACHO\\" + year;  // 해당 연도 폴더를 지정한다. 


            DirectoryInfo dirs = new DirectoryInfo(Dirname);
            DirectoryInfo[] DIRS = dirs.GetDirectories();
            int IdCount = 0;
            for (int i = 0; i < DIRS.Length; i++)
            {
                if (DIRS[i].ToString() == StartMonth)
                {
                    if (ViewerMode == true)
                    {
                        Dirname = @"\\" + ShareIP + "\\tacho2\\" + "\\TACHO\\" + year + "\\" + StartMonth;  // 해당 연도 폴더를 지정한다. 
                    }
                    else
                    {
                        Dirname = TACHO2_path + "\\TACHO\\" + year + "\\" + StartMonth;  // Month 폴더를 지정한다.
                    }
                    //     Dirname = form1.TACHO2_path + "\\TACHO\\" + year + "\\" + StartMonth;  // Month 폴더를 지정한다.
                    DirectoryInfo mdbPath = new DirectoryInfo(Dirname);
                    FileInfo[] files = mdbPath.GetFiles();   // 파일 가져온다.
                    string[] file_str = new string[files.Length];
                    char[] trimChars = { '.', 'm', 'd', 'b' };
                    int cnt = 0;
                    for (int a = 0; a < files.Length; a++)   // mdb 파일 목록을 가져온다.
                    {

                        if (files[a].Extension != ".ldb")
                        {

                            file_str[a] = files[a].ToString();
                            file_str[a] = file_str[a].TrimEnd(trimChars);

                        }
                        string DBstring = "";
                        if (file_str[a] == null)
                        {
                            continue;
                        }
                        DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + Dirname + "\\" + file_str[a] + ".mdb";

                        string queryRead = "select * from TblTacho";

                        OleDbConnection conn = new OleDbConnection(@DBstring);
                        conn.Open();
                        OleDbCommand commRead = new OleDbCommand(queryRead, conn);
                        OleDbDataReader srRead = commRead.ExecuteReader();


                        double TotalMoney = 0;
                        double TotalRealncome = 0;
                        double TotalDist = 0;


                        while (srRead.Read())
                        {

                            string tooltip_str = srRead.GetInt32(0).ToString();


                            DateTime stTime = new DateTime(1, 1, 1, 0, 0, 0);
                            DateTime srTime = new DateTime(1, 1, 1, 0, 0, 0);
                            TimeSpan stSpan = new TimeSpan(0, 0, 0, 0);





                            string strCarNo = srRead.GetString(1);  // 차량번호

                            strCarNo = strCarNo.TrimEnd();
                            mdbTime = srRead.GetDateTime(4); // begin Time
                            string strDriverNo = "";

                            if (srRead.IsDBNull(2) == false)
                            {
                                strDriverNo = srRead.GetString(2);
                                strDriverNo = strDriverNo.Trim();
                              //  CarNO_or_DirverID = CarNO_or_DirverID.Trim();

                            }

                            mdbTime = new DateTime(mdbTime.Year, mdbTime.Month, mdbTime.Day, 0, 0, 0);


                            if (mdbTime < Starttime || mdbTime > Endtime)
                                continue;


                            CarNO = CarNO.TrimStart();    
                                if (strCarNo != CarNO) continue;  // 차량 번호 걸러내기 

                                DirverID = DirverID.TrimStart();
                                if (strDriverNo != DirverID) continue; // 기사 번호 걸러내기 
                         



                            IdCount++;


                            double dotsatus = srRead.GetDouble(10);
                            double uuu = (double)srRead.GetDouble(6);
                            double RealIncome = 0;
                            if (srRead.IsDBNull(11) == false)
                            {
                                 RealIncome = (double)srRead.GetDouble(11);
                            }
                            string money;
                            if (dotsatus != 4 && dotsatus != 2 && dotsatus != 8)
                            {
                                TotalMoney += uuu;
                                TotalRealncome += RealIncome;
                                int hhh = (int)uuu;
                                money = string.Format("{0}", (int)hhh);

                            }
                            else
                            {
                                if (dotsatus == 8)
                                {
                                    TotalMoney += uuu;
                                    TotalRealncome += RealIncome;
                                    money = string.Format("{0:F3}", uuu);

                                }
                                else
                                {
                                    TotalMoney += uuu;
                                    TotalRealncome += RealIncome;
                                    money = string.Format("{0:F2}", uuu);

                                }
                            }// 미터수입
                                           
                          
                        }//while

                        GtandTotlaMoney += TotalMoney;
                        GrandTotalRealIncome += TotalRealncome; 
                        conn.Close();
                    }
               //     Monthly_temp = string.Format("{0:N2}", GtandTotlaMoney);  // toll 
                    Monthly_temp = string.Format("{0:N2}", GrandTotalRealIncome);  // toll total
                }
                else if (DIRS[i].ToString() == EndMonth)
                {
                    Dirname = TACHO2_path + "\\TACHO\\" + year + "\\" + EndMonth;  // Month 폴더를 지정한다.
                    DirectoryInfo mdbPath = new DirectoryInfo(Dirname);
                    FileInfo[] files = mdbPath.GetFiles();   // 파일 가져온다.
                    string[] file_str = new string[files.Length];
                    char[] trimChars = { '.', 'm', 'd', 'b' };
                    int cnt = 0;

                    for (int a = 0; a < files.Length; a++)   // mdb 파일 목록을 가져온다.
                    {

                        if (files[a].Extension != ".ldb")
                        {

                            file_str[a] = files[a].ToString();
                            file_str[a] = file_str[a].TrimEnd(trimChars);

                        }
                        string DBstring = "";
                        if (file_str[a] == null)
                        {
                            continue;
                        }
                        DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + Dirname + "\\" + file_str[a] + ".mdb";

                        string queryRead = "select * from TblTacho";

                        OleDbConnection conn = new OleDbConnection(@DBstring);
                        conn.Open();
                        OleDbCommand commRead = new OleDbCommand(queryRead, conn);
                        OleDbDataReader srRead = commRead.ExecuteReader();


                        double TotalMoney = 0;
                        double TotalDist = 0;
                        double TotalRealncome = 0;

                         TotalMoney = 0;
                         TotalDist = 0;



                         while (srRead.Read())
                         {

                             string tooltip_str = srRead.GetInt32(0).ToString();


                             DateTime stTime = new DateTime(1, 1, 1, 0, 0, 0);
                             DateTime srTime = new DateTime(1, 1, 1, 0, 0, 0);
                             TimeSpan stSpan = new TimeSpan(0, 0, 0, 0);





                             string strCarNo = srRead.GetString(1);  // 차량번호

                             strCarNo = strCarNo.TrimEnd();
                             mdbTime = srRead.GetDateTime(4); // begin Time
                             string strDriverNo = "";

                             if (srRead.IsDBNull(2) == false)
                             {
                                 strDriverNo = srRead.GetString(2);
                                 strDriverNo = strDriverNo.TrimEnd();

                             }
                             mdbTime = new DateTime(mdbTime.Year, mdbTime.Month, mdbTime.Day, 0, 0, 0);


                             if (mdbTime < Starttime || mdbTime > Endtime)
                                 continue;


                             CarNO = CarNO.TrimStart();
                             if (strCarNo != CarNO) continue;  // 차량 번호 걸러내기 


                             DirverID = DirverID.TrimStart();
                             if (strDriverNo != DirverID) continue; // 기사 번호 걸러내기 




                             IdCount++;




                             double dotsatus = srRead.GetDouble(10);
                             double RealIncome = (double)srRead.GetDouble(11);
                             double uuu = (double)srRead.GetDouble(6);

                             string money;
                             if (dotsatus != 4 && dotsatus != 2 && dotsatus != 8)
                             {
                                 TotalMoney += uuu;
                                 TotalRealncome += RealIncome;
                                 int hhh = (int)uuu;
                                 money = string.Format("{0}", (int)hhh);

                             }
                             else
                             {
                                 if (dotsatus == 8)
                                 {
                                     TotalMoney += uuu;
                                     TotalRealncome += RealIncome;
                                     money = string.Format("{0:F3}", uuu);

                                 }
                                 else
                                 {
                                     TotalMoney += uuu;
                                     TotalRealncome += RealIncome;
                                     money = string.Format("{0:F2}", uuu);

                                 }
                             }// 미터수입
                         }
                         GtandTotlaMoney += TotalMoney;
                         GrandTotalRealIncome += TotalRealncome; 
                         conn.Close();
                    }
                   // Monthly_temp = string.Format("{0:N2}", GtandTotlaMoney);  // total
                    Monthly_temp = string.Format("{0:N2}", GrandTotalRealIncome);  // total
                }

            }
         

           


          
         
        }
        public void MonthlySearchData(string CarNO,string DirverID, int Year,int Month)
        {

            // id ==0 차량번호 
            //      1 기사번호
            //      2 all

            //  StartDay
          //   Starttime = new DateTime(Starttime.Year, Starttime.Month, Starttime.Day, 0, 0, 0);
         //    Endtime = new DateTime(Endtime.Year, Endtime.Month, Endtime.Day, 0, 0, 0);

            string Dirname = "";
            int nCnt = 1;
            double GtandTotlaMoney = 0;
           
            DateTime mdbTime = new DateTime();

            Dirname = TACHO2_path + "\\TACHO\\" + Year;  // 해당 연도 폴더를 지정한다. 
           
            //   Dirname = form1.TACHO2_path + "\\TACHO\\" + year;  // 해당 연도 폴더를 지정한다. 


            DirectoryInfo dirs = new DirectoryInfo(Dirname);
            DirectoryInfo[] DIRS = dirs.GetDirectories();
            int IdCount = 0;
           
            for (int i = 0; i < DIRS.Length; i++)
            {
                if (DIRS[i].ToString() == Month.ToString())
                {
                   
                     Dirname = TACHO2_path + "\\TACHO\\" + Year + "\\" + Month;  // Month 폴더를 지정한다.
                  
                    //     Dirname = form1.TACHO2_path + "\\TACHO\\" + year + "\\" + StartMonth;  // Month 폴더를 지정한다.
                    DirectoryInfo mdbPath = new DirectoryInfo(Dirname);
                    FileInfo[] files = mdbPath.GetFiles();   // 파일 가져온다.
                    string[] file_str = new string[files.Length];
                    char[] trimChars = { '.', 'm', 'd', 'b' };
                    int cnt = 0;

                    for (int a = 0; a < files.Length; a++)   // mdb 파일 목록을 가져온다.
                    {

                        if (files[a].Extension != ".ldb")
                        {

                            file_str[a] = files[a].ToString();
                            file_str[a] = file_str[a].TrimEnd(trimChars);

                        }
                        string DBstring = "";
                        if (file_str[a] == null)
                        {
                            continue;
                        }
                        DBstring = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + Dirname + "\\" + file_str[a] + ".mdb";

                        string queryRead = "select * from TblTacho";

                        OleDbConnection conn = new OleDbConnection(@DBstring);
                        conn.Open();
                        OleDbCommand commRead = new OleDbCommand(queryRead, conn);
                        OleDbDataReader srRead = commRead.ExecuteReader();


                        double TotalMoney = 0;
                        double TotalDist = 0;


                        while (srRead.Read())
                        {

                            string tooltip_str = srRead.GetInt32(0).ToString();


                            DateTime stTime = new DateTime(1, 1, 1, 0, 0, 0);
                            DateTime srTime = new DateTime(1, 1, 1, 0, 0, 0);
                            TimeSpan stSpan = new TimeSpan(0, 0, 0, 0);





                            string strCarNo = srRead.GetString(1);  // 차량번호

                            strCarNo = strCarNo.TrimEnd();
                            mdbTime = srRead.GetDateTime(4); // begin Time
                            string strDriverNo = "";

                            if (srRead.IsDBNull(2) == false)
                            {
                                strDriverNo = srRead.GetString(2);
                                strDriverNo = strDriverNo.TrimEnd();

                            }


                            CarNO = CarNO.TrimStart();                          
                                if (strCarNo != CarNO) continue;  // 차량 번호 걸러내기 


                                DirverID = DirverID.TrimStart();
                                if (strDriverNo != DirverID) continue; // 기사 번호 걸러내기 
                           



                            IdCount++;
                             
                         


                            double dotsatus = srRead.GetDouble(10);



                            double uuu = (double)srRead.GetDouble(6);

                            string money;
                            if (dotsatus != 4 && dotsatus != 2 && dotsatus != 8)
                            {
                                TotalMoney += uuu;
                                int hhh = (int)uuu;
                                money = string.Format("{0}", (int)hhh);
                             
                            }
                            else
                            {
                                if (dotsatus == 8)
                                {
                                    TotalMoney += uuu;
                                    money = string.Format("{0:F3}", uuu);
                                  
                                }
                                else
                                {
                                    TotalMoney += uuu;
                                    money = string.Format("{0:F2}", uuu);
                                   
                                }
                            }// 미터수입



                         
                          
                        }
                        GtandTotlaMoney += TotalMoney;
                        conn.Close();
                    }
                    Monthly_temp = string.Format("{0:N2}", GtandTotlaMoney);  // toll total

                }
            

            }
          

           




        }
        private void button6_Click_1(object sender, EventArgs e)
        {
             if (MiniPrintCon == true)
             {
                     Thread.Sleep(1000);
                     Driver_Receipt();
             }
          /*  Loading_Visible(true);
            Treeview_Refresh();

            string path = "";
            if (ViewerMode == true)
            {
                path = @"\\" + ShareIP + "\\tacho2\\";
            }
           
            //   path = TACHO2_path + e.Node.Parent.Parent.Parent.Text + "\\" + e.Node.Parent.Parent.Text + "\\" + e.Node.Parent.Text;
         

            mdbfilename = path + treeView1.Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Nodes[0].Text + "\\" + treeView1.Nodes[0].Nodes[0].Nodes[0].Nodes[0].Text + ".mdb";

            DB_ReadData(0, 1);
            Loading_Visible(false);*/


         /*   Sharethread = new Thread(new ThreadStart(ShareRun));
            Sharethread.IsBackground = true;
            Thread.Sleep(100);
            Sharethread.Start();*/
            

         

         /*   while (true)
            {
                AMF_Data(Application.StartupPath + "\\test.AMF");
                Thread.Sleep(100);
                
            }*/

        }

        public void CopyFolder(string sourceFolder, string destFolder)
        {
            if (!Directory.Exists(destFolder))
                Directory.CreateDirectory(destFolder);

            string[] files = Directory.GetFiles(sourceFolder);
            string[] folders = Directory.GetDirectories(sourceFolder);

            foreach (string file in files)
            {
                string name = Path.GetFileName(file);
                string dest = Path.Combine(destFolder, name);

                /*
                //////////// 파일 크기 비교 //////////
                if (System.IO.File.Exists(dest))  // 같은 파일의이름이 존재 함
                {
                    System.IO.FileInfo Orgfile = new System.IO.FileInfo(@file);
                    long OSize = Orgfile.Length;

                    System.IO.FileInfo Newfile = new System.IO.FileInfo(@dest);
                    long nSize = Newfile.Length;

                    if (OSize != nSize)
                    {
                        System.IO.File.Copy(file, dest, true);
                    }
                }
                else
                {
                 
                    System.IO.File.Copy(file, dest, true);
                }*/



                System.IO.File.Copy(file, dest,true);
            }

            // foreach 안에서 재귀 함수를 통해서 폴더 복사 및 파일 복사 진행 완료  
            foreach (string folder in folders)
            {
                string name = Path.GetFileName(folder);
                string dest = Path.Combine(destFolder, name);
                CopyFolder(folder, dest);
              
            }
        }

        private void pictureBox17_VisibleChanged(object sender, EventArgs e)
        {
            if (pictureBox17.Visible == true)
            {
                button16.Enabled = false;
            }
            else
            {
                button16.Enabled = true;
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            CashierForm cashierForm = new CashierForm(this);
            cashierForm.Cashie_Total();
            cashierForm.ShowDialog();
          
        }  
     

        /*
        //Method to overwrite that manages the arrival of new storage units
            protected override void WndProc(ref Message m)
            {
              //This definitions are stored in “dbt.h” and “winuser.h”
              // There has been a change in the devices
              const int WM_DEVICECHANGE = 0x0219;
              // System detects a new device
              const int DBT_DEVICEARRIVAL = 0x8000;
              // Device removal request
              const int DBT_DEVICEQUERYREMOVE = 0x8001;
              // Device removal failed
              const int DBT_DEVICEQUERYREMOVEFAILED = 0x8002;
              // Device removal is pending
              const int DBT_DEVICEREMOVEPENDING = 0x8003;
              // The device has been succesfully removed from the system
              const int DBT_DEVICEREMOVECOMPLETE = 0x8004;
              // Logical Volume (A disk has been inserted, such a usb key or external HDD)
              const int DBT_DEVTYP_VOLUME = 0x00000002;
              switch (m.Msg)
              {
               //If system devices change…
               case WM_DEVICECHANGE:
                switch (m.WParam.ToInt32())
                {
                 //If there is a new device…
                 case DBT_DEVICEARRIVAL:
                 {
                  int devType = Marshal.ReadInt32(m.LParam, 4);
                  //…and is a Logical Volume (A storage device)
                  if (devType == DBT_DEVTYP_VOLUME)
                  {
                   DEV_BROADCAST_VOLUME vol;
                   vol = (DEV_BROADCAST_VOLUME)Marshal.PtrToStructure(
                   m.LParam, typeof(DEV_BROADCAST_VOLUME));
                  MessageBox.Show(
                    "A storage device has been inserted, unit: " +
                    UnitName(vol.dbcv_unitmask));
                  }
                 }
                break;
                case DBT_DEVICEREMOVECOMPLETE:
                 MessageBox.Show("Device removed from system.");
                break;
               }
               break;
              }
              //After the custom manager, we want to use the default system’s manager
              base.WndProc(ref m);
            }
             
            //Method to detect the unit name (”D:”, “F:”, etc)
            char UnitName(int unitmask)
            {
              char[] units ={ 'A', 'B', 'C', 'D', 'E', 'F', 'G',
                'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P',
                'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
              int i = 0;
              //Convert the mask in an array, and search
              //the index for the first occurrenc (the unit’s name)
              System.Collections.BitArray ba = new
                System.Collections.BitArray(System.BitConverter.GetBytes(unitmask));
              foreach (bool var in ba)
              {
               if (var == true)
               break;
              i++;
              }
            return units[i];
            }
        */

           


    }
}