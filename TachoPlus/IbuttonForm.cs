using System;
using System.Security.Permissions;
using Microsoft.Win32;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.IO.Ports;
using System.Net;


namespace TachoPlus
{
    public partial class IbuttonForm : Form
    {
        Form1 form1;
     //   public TMEX TMEXLibrary;
      //  public short portNum;
     //   public short portType;
     //   public int hSess;
        public int sessionOptions;
        public byte[] stateBuffer = new byte[15360];						// internal for TMX interface
     //   public short[] ROM = new short[8];
        private List<byte> FileData = new List<byte>();

        public IbuttonForm(Form1 f)
        {

            InitializeComponent();
            form1 = f;
            readBytesButton.Visible = false;
            /*
            TMEXLibrary = new TMEX();
            TMEXLibrary.TMReadDefaultPort(ref portNum, ref portType);

             hSess = TMEXLibrary.TMExtendedStartSession(portNum, portType, ref sessionOptions);
            if (hSess != 0)
            {
              
                // must be called before any non-session functions can be called
                TMEXLibrary.TMSetup(hSess);

                short ret = TMEXLibrary.TMFirst(hSess, stateBuffer);

                if (ret != 1)
                {
                    MessageBox.Show("Disconnected  to ibutton.");
                    return;
                }

                int n = 0;

                //   while ( (ret==1) && (n<32) )
                //   {
                // if MSB of ROM[0] != 0, then write, else read
                //   short[] ROM  = new short[8];

                TMEXLibrary.TMRom(hSess, stateBuffer, ROM);  // Selecet
                n = TMEXLibrary.TMStrongAccess(hSess, stateBuffer);

                if (ROM[0] != 0x0c)
                {
                    MessageBox.Show("Disconnected  to ibutton.");
                    form1.ibutonCheck = true;
                    return;
                }

                //  }
            }
             */

          //  addText(ToHex(block, 0, block.Length), Color.Black);
        }

        private void readBytesButton_Click(object sender, EventArgs e)  // Read
        {
            try
            {

                oneWireIOTextBox.Clear();
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
                if (hSess != 0)
                {

                    // must be called before any non-session functions can be called
                    TMEXLibrary.TMSetup(hSess);

                    ret = TMEXLibrary.TMFirst(hSess, stateBuffer);

                    if (ret != 1)
                    {
                        MessageBox.Show("Disconnected  to ibutton.");
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
                        form1.ibutonCheck = true;
                        return;
                    }
                }

                //  }

                ///////////////////////////////////////////////////////////////




                TMEXLibrary.TMSetup(hSess);




                ret = TMEXLibrary.TMFirst(hSess, stateBuffer);


                if (ROM[0] != 0x0c)
                {
                    MessageBox.Show("Disconnected  to ibutton.");
                    this.Close();
                    return;
                    //  this.Close();
                }


                if (ret != 1)
                {
                    MessageBox.Show("Ibutton Fail");
                    return;
                }
                /////////////////////////////////////////////  타코 Start Byte 읽기
                Address[0] = 0x00;
                Address[1] = 0x00;

                TachoStartByte = OnWire_Read(TachoStartByte, TachoStartByte.Length, Address, TMEXLibrary, hSess, ROM);

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


                if (oneWireIOTextBox.Visible == true)
                {
                    addText(ToHex(Amf_array, 0, Amf_array.Length), Color.Black);
                }


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

                if (Model == "\0\0\0\0\0")
                {
                    Model = "Pro1+";
                }
                Model = "Pro1+";

                string newPath = System.IO.Path.Combine(form1.TACHO2_path + "\\TMF", "Auto");
                // Create the subfolder
                System.IO.Directory.CreateDirectory(newPath);

                string TmpFile = form1.TACHO2_path + "\\TMF\\auto\\" + Model + "_" + NowReceiveTime + "_" + carnum + ".AMF";

                form1.Amf_path = TmpFile;

                // byte[] rcvByte = new byte[mStreamBuffer.Length];
                // rcvList.CopyTo(rcvByte);

                FileStream fs = new FileStream(TmpFile, FileMode.OpenOrCreate, FileAccess.Write);
                BinaryWriter bw = new BinaryWriter(fs);

                bw.Write(Amf_array);



                fs.Close();
                bw.Close();



                form1.AMF_Data(TmpFile);



                //////////////////////////////////////   2번째 저장 total tmf  기존 데이터와  이어 붙히자 !!///////////////////

                string TMFPath = System.IO.Path.Combine(form1.TACHO2_path + "\\TMF", "TransData");
                // Create the subfolder
                System.IO.Directory.CreateDirectory(TMFPath);

                NowReceiveTime = String.Format("{0:D2}{1:D2}{2:D2}",
                                                        (DateTime.Now.Year - 2000), DateTime.Now.Month, DateTime.Now.Day);
                TmpFile = form1.TACHO2_path + "\\TMF\\TransData\\" + NowReceiveTime + ".AMF";

                // rcvList.RemoveAt(0);
                //  rcvByte = new byte[rcvList.Count];
                //    rcvList.CopyTo(rcvByte);

                fs = new FileStream(TmpFile, FileMode.Append, FileAccess.Write);
                bw = new BinaryWriter(fs);

                bw.Write(Amf_array);
                fs.Close();
                bw.Close();

                /////////////////////////////////////////////////////////////////////////////////////////////////////////////


                if (TachoStartByte[0] == 0xA3 || TachoStartByte[0] == 0xA4)
                {
                    TachoStartByte[0] = 0xff;  

                    Address[0] = 0x00;
                    Address[1] = 0x00;
                    OneWire_Write(TachoStartByte, TachoStartByte.Length, Address, TMEXLibrary, hSess);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

   
            MessageBox.Show("Sucessfully received tacho data! ");


            
        }
        public byte[] OnWire_Read(byte[] Buffer, int Length, byte[] Address, TMEX TMEXLibrary,int hSess,short[] ROM  )
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


        private void button3_Click(object sender, EventArgs e)  // Write
        {

            
            string CarNum = textBox1.Text;
            string DriverID = textBox3.Text;
            string TimLimitStr = textBox2.Text;
            TimLimitStr += textBox4.Text;

            byte[] auth = new byte[3];
            byte[] CarByte = new byte[9];
            byte[] DriverByte = new byte[9];
            byte[] TimLimitByte = new byte[2];

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
            if (hSess != 0)
            {

                // must be called before any non-session functions can be called
                TMEXLibrary.TMSetup(hSess);

                ret = TMEXLibrary.TMFirst(hSess, stateBuffer);

                if (ret != 1)
                {
                    MessageBox.Show("Disconnected  to ibutton.");
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
                    form1.ibutonCheck = true;
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


            if (textBox1.Text == "8659858")
            { // width =355
                this.Width = 715;
                oneWireIOTextBox.Visible = true;
                readBytesButton.Visible = true;
                return;
            }


            if (checkBox1.Checked == false && checkBox3.Checked == false && checkBox4.Checked == false)
            {

            }
            else
            {


                if (checkBox1.Checked == true)
                {
                    for (int i = 0; i < 9; i++)
                    {
                        CarByte[i] = 0x20;
                    }

                    if (textBox1.Text == "")
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

                    OneWire_Write(CarByte, CarByte.Length, CarAddr, TMEXLibrary, hSess);

                    byte[] Temp = new byte[CarByte.Length];

                    Temp = OnWire_Read(Temp, Temp.Length, CarAddr, TMEXLibrary, hSess, ROM);
                         
                    for (int i = 0; i < Temp.Length; i++)
                    {

                        if (Temp[i] != CarByte[i])
                        {
                            MessageBox.Show("Write Fail!");
                            return;
                        }
                    }
                 


                }

                if (checkBox3.Checked == true)
                {
                    for (int i = 0; i < 9; i++)
                    {
                        DriverByte[i] = 0x20;
                    }
                    if (textBox3.Text == "")
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
                    OneWire_Write(DriverByte, DriverByte.Length, DriverAddr, TMEXLibrary, hSess);


                    byte[] Temp = new byte[DriverByte.Length];

                    Temp = OnWire_Read(Temp, Temp.Length, DriverAddr, TMEXLibrary, hSess, ROM);

                    for (int i = 0; i < Temp.Length; i++)
                    {

                        if (Temp[i] != DriverByte[i])
                        {
                            MessageBox.Show("Write Fail!");
                            return;
                        }
                    }
                 



                }


                if (checkBox4.Checked == true)
                {
                    byte[] temp = new byte[2];

                    if (textBox2.Text.Length == 1)
                    {
                        string str = "0";

                        str += textBox2.Text;

                        temp = FromHex(str);
                        TimLimitByte[0] = temp[0];

                    }
                    else if (textBox2.Text.Length == 2)
                    {
                        temp = FromHex(textBox2.Text);
                        TimLimitByte[0] = temp[0];
                    }
                    else
                    {
                        textBox2.Text = "00";

                        temp = FromHex(textBox2.Text);
                        TimLimitByte[0] = temp[0];
                    }




                    if (textBox4.Text.Length == 1)
                    {
                        string str = "0";

                        str += textBox4.Text;

                        temp = FromHex(str);
                        TimLimitByte[1] = temp[0];
                    }
                    else if (textBox4.Text.Length == 2)
                    {
                        temp = FromHex(textBox4.Text);
                        TimLimitByte[1] = temp[0];
                    }
                    else
                    {
                        textBox4.Text = "00";

                        temp = FromHex(textBox4.Text);
                        TimLimitByte[1] = temp[0];
                    }




                    OneWire_Write(TimLimitByte, TimLimitByte.Length, TimeLimitAddr, TMEXLibrary, hSess);



                    byte[] Temp = new byte[TimLimitByte.Length];

                    Temp = OnWire_Read(Temp, Temp.Length, TimeLimitAddr, TMEXLibrary, hSess, ROM);

                    for (int i = 0; i < Temp.Length; i++)
                    {

                        if (Temp[i] != TimLimitByte[i])
                        {
                            MessageBox.Show("Write Fail!");
                            return;
                        }
                    }

                }


               
                MessageBox.Show("Sucessfully Ibutton update!");
            }


            



            ////////////// File 읽어 쓰기////////////////////////////////////
           /* OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Title = "Open File for upgrade";
            openFileDialog1.Filter = "*.*|*.*";
        
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
            
                    FileData.Clear();
                    FileStream fs = new FileStream(openFileDialog1.FileName, FileMode.Open, FileAccess.Read);
                    BinaryReader br = new BinaryReader(fs);
                    int num = 0;                 
                    num = (int)fs.Length / 100;

                    do
                    {

                        byte bInData = br.ReadByte();
                        FileData.Add(bInData);
                    
                    } while (fs.Length != fs.Position);
                 
                    fs.Close();
                    br.Close(); 
                              
            }


          

            byte[] auth = new byte[3];
            //  byte[] data = new byte[32];
            //   byte[] block = FromHex(bytesTextBox.Text);


            byte[] DataArry = new byte[FileData.Count];
            byte[] block = new byte[32];

            FileData.CopyTo(DataArry);


            TMEXLibrary.TMSetup(hSess);

            short ret = TMEXLibrary.TMFirst(hSess, stateBuffer);

            if (ret != 1)
            {
                MessageBox.Show("Ibutton Fail");
                return;
            }

            int cnt = 0;
            byte Addr = 0x00;
            for (int a = 0; a < 10; a++)
            {
                byte Addr1 = 0x00; 

                if (a > 7)
                {
                    Addr1 = 0x01; 
                }
                for (int i = 0; i < 32; i++)
                {

                    block[i] = FileData[i + cnt];
                }
                cnt += 32;
                TMEXLibrary.TMAccess(hSess, stateBuffer);
                TMEXLibrary.TMTouchByte(hSess, 0x0F);
                TMEXLibrary.TMTouchByte(hSess, (short)Addr);
                TMEXLibrary.TMTouchByte(hSess, (short)Addr1);
                TMEXLibrary.TMBlockStream(hSess, block, (short)block.Length);

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
                Addr += 0x20;
            }


            byte[] block1 = new byte[2];
            block1[0] = 0xfd;
            block1[1] = 0x8a;
            TMEXLibrary.TMAccess(hSess, stateBuffer);
            TMEXLibrary.TMTouchByte(hSess, 0x0F);
            TMEXLibrary.TMTouchByte(hSess, (short)0x40);
            TMEXLibrary.TMTouchByte(hSess, (short)0x01);
            TMEXLibrary.TMBlockStream(hSess, block1, (short)block1.Length);

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
            }*/
        }
        private void addText(string text, Color col)
        {
            try
            {
                oneWireIOTextBox.SelectionColor = col;
                oneWireIOTextBox.AppendText(text + "\r\n");
                //oneWireIOTextBox.SelectionStart = oneWireIOTextBox.Text.Length - 1;
                oneWireIOTextBox.Focus();
                oneWireIOTextBox.ScrollToCaret();
               // bytesTextBox.Focus();
               // bytesTextBox.SelectAll();
            }
            catch (Exception e)
            {
                Console.Out.WriteLine(e);
            }
        }

        private static string ToHex(byte[] buff, int off, int len)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder(buff.Length * 3);
            sb.Append(buff[off].ToString("X2"));
            for (int i = 1; i < len; i++)
            {
                sb.Append(" ");
                sb.Append(buff[off + i].ToString("X2"));
            }
            return sb.ToString();
        }

        private static byte[] FromHex(string s)
        {
            s = System.Text.RegularExpressions.Regex.Replace(s.ToUpper(), "[^0-9A-F]", "");
            byte[] b = new byte[s.Length / 2];
            for (int i = 0; i < s.Length; i += 2)
                b[i / 2] = byte.Parse(s.Substring(i, 2),
                   System.Globalization.NumberStyles.AllowHexSpecifier);

            return b;
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsDigit(e.KeyChar) || e.KeyChar == Convert.ToChar(Keys.Back)))
            {
                e.Handled = true;
            }
        }

        private void IbuttonForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            form1.IbuttonSetting = false;
        }

       

    }
}