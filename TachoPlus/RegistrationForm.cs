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
using System.Security.Cryptography;  

namespace TachoPlus
{
    public partial class RegistrationForm : Form
    {
        private iniClass inicls = new iniClass();
        public string Hdd_Serial = "";
        public string Inicode = "";
        public string Encrypt = "";
        public string Dencrypt = "";
        const string SMTP_SERVER = "smtp.naver.com"; // SMTP 서버 주소

        const int SMTP_PORT = 587; // SMTP 포트
         const string MAIL_ID = "kairae@naver.com"; // 보내는 사람의 이메일

        const string MAIL_ID_NAME = "kairae"; // 보내는사람 계정 ( 네이버 로그인 아이디 ) 

        const string MAIL_PW = "dnflskfk63";  // 보내는사람 패스워드 ( 네이버 로그인 패스워드 )

        public RegistrationForm()
        {
            InitializeComponent();
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
        private void button2_Click(object sender, EventArgs e)
        {
            string drive = "";
            if (drive == "" || drive == null)
            {

                drive = "C";

            }
            ManagementObject disk = new ManagementObject("win32_logicaldisk.deviceid=\"" + drive + ":\"");
            disk.Get();
            //   MessageBox.Show(disk["VolumeSerialNumber"].ToString());

            Hdd_Serial = disk["VolumeSerialNumber"].ToString();


            Encrypt = EncryptString(Hdd_Serial,"JIE");

            Dencrypt = DecryptString(Encrypt, "JIE");

             try
            {


                MailAddress mailFrom = new MailAddress(MAIL_ID, MAIL_ID_NAME, Encoding.UTF8); // 보내는사람의 정보를 생성
                MailAddress mailTo = new MailAddress("kairae@naver.com"); // 받는사람의 정보를 생성


                SmtpClient client = new SmtpClient(SMTP_SERVER, SMTP_PORT); // smtp 서버 정보를 생성

                MailMessage message = new MailMessage(mailFrom, mailTo);

                message.Subject = "TachoPlus 인증"; // 메일 제목 프로퍼티

                message.Body =    "보내는  사람 : " + textBox2.Text + "\n" +
                                  "         이름 : " + textBox3.Text + "\n" +
                                  "  Encryp Code : " + Encrypt;     // 내용 

                message.BodyEncoding = Encoding.UTF8; // 메세지 인코딩 형식

                message.SubjectEncoding = Encoding.UTF8; // 제목 인코딩 형식




                client.EnableSsl = true; // SSL 사용 유무 (네이버는 SSL을 사용합니다. )

                client.DeliveryMethod = SmtpDeliveryMethod.Network;

                client.Credentials = new System.Net.NetworkCredential(MAIL_ID, MAIL_PW); // 보안인증 ( 로그인 )

                client.Send(message);  //메일 전송 
                System.Windows.Forms.MessageBox.Show("Authentication complete e-mail request.");



            }

            catch (Exception ex)

            {

                MessageBox.Show(ex.Message);

            }

            /*
            //////////////////////////////////////////////////////////////////  email send

            // !!!!!!!!!!!!!!!!!! 지메일 보안인 낮은 수준의 앱을 사용을 반드시 허락으로 설정하여야 사용이 가능하다.!!!!!!
            MailMessage mail = new MailMessage();
                // !!!!!!!!!!!!!!!!!! 지메일 보안인 낮은 수준의 앱을 사용을 반드시 허락으로 설정하여야 사용이 가능하다.!!!!!!
            mail.From = new MailAddress("kairae0307@gmail.com");  // 보내는 사람 이메일
            mail.To.Add("kairae@naver.com");  //받는 사람 이메일
            mail.Subject = "Tacho Plus 인증건";  // 제목 
            mail.Body = "보내는  사람 : " + textBox2.Text + "\n" +
                       "         이름 : " + textBox3.Text + "\n" +
                       "  Encryp Code : " + Encrypt;     // 내용 



            //  System.Net.Mail.Attachment attachmnt;  // 첨부 파일  
            //   attachmnt = new System.Net.Mail.Attachment("c:\\1111.PNG"); // 파일이름
            //mail.Attachments.Add(attachmnt); // 첨부파일 붙히기i
              
            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);  // 지메일 포트 설정
            smtp.UseDefaultCredentials = false;
            smtp.EnableSsl = true;

            smtp.Credentials = new NetworkCredential("kairae03071@gmail.com", "jie8659858");  // 아뒤 , 비밀 번호 

            try
            {
                smtp.Send(mail);
                System.Windows.Forms.MessageBox.Show("Authentication complete e-mail request.");
            }
            catch (SmtpException ex)
            {
                MessageBox.Show(ex.Message);
            }*/
            
            ///////////////////////////////////////////////////////////////////
        }

        private void RegistrationForm_Load(object sender, EventArgs e)
        {

            string path = Application.StartupPath + "\\TachoPlus.ini";

             Inicode = inicls.GetIniValue("Tacho Init", "Registration", path);

             string drive = "";
             if (drive == "" || drive == null)
             {

                 drive = "C";

             }
             ManagementObject disk = new ManagementObject("win32_logicaldisk.deviceid=\"" + drive + ":\"");
             disk.Get();
             //   MessageBox.Show(disk["VolumeSerialNumber"].ToString());

             Hdd_Serial = disk["VolumeSerialNumber"].ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
         //   string path = Application.StartupPath + "\\TachoPlus.ini";

         //   inicls.SetIniValue("Tacho Init", "Registration",textBox1.Text, path);  //전송방식 쓰기

            RegistryKey reg;
            reg = Registry.LocalMachine.CreateSubKey("SOFTWARE").CreateSubKey("TachoPlus");
            reg.SetValue("Text", textBox1.Text);
            reg.SetValue("Num", 5);
            reg.SetValue("Check", true);
            reg.Close();
            this.DialogResult = DialogResult.OK;


          

        }

        private void RegistrationForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            
               // Application.Exit();
            
        }
    }
}
