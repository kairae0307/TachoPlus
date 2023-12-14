using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Sockets;

namespace TachoPlus
{
    class PacketHandler
    {
        //테스트 용
        static public void Handle_ChangeDoor(byte[] databyte, PacketSocket client)
        {
            try
            {
                lock (client.GetClientManager())
                {
                    Pt_ChangeDoor ptChangeDoor = new Pt_ChangeDoor(databyte);
                    System.Console.WriteLine("From {0} CHANGEDOOR {1}번 문을 {2}번 상태로 바꿉니다.", client.GetSocket().Handle, ptChangeDoor.GetDoorNum(), ptChangeDoor.GetChangeState());

                    Pt_GeneralAck ptAck = new Pt_GeneralAck(ePt_Type.PT_CHANGE_DOOR, ePt_AckType.OK);

                    client.BroadCast(ptAck.GetBytes(), ptAck.GetSize());
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message.ToString());           
            }
        }

        //테스트 용
        static public void Handle_GeneralAck(byte[] databyte, PacketSocket client)
        {
            try
            {
                lock (client.GetClientManager())
                {
                    Pt_GeneralAck ptGeneralAck = new Pt_GeneralAck(databyte);
                    System.Console.WriteLine("Ack타입 {0}\nAck상태 {1}", ptGeneralAck.GetRequestType(), ptGeneralAck.GetResult());
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message.ToString());  
            }
        }

        //채팅 메세지
        static public void Handle_ChatMsg(byte[] databyte, PacketSocket client)
        {
            try
            {
                lock (client.GetClientManager())
                {
                  //  Pt_ChatMsg ptChatMsg = new Pt_ChatMsg(databyte);
                  //  string strMsg = Encoding.Default.GetString(ptChatMsg.GetChatMsg());

                    client.BroadCast(databyte, databyte.Length);
                  //  Console.WriteLine("From {0} CHAT_MSG: {1}", client.GetSocket().Handle, strMsg);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message.ToString());
            }
        }

        //로그인
        static public void Handle_LogIn(byte[] databyte, PacketSocket client)
        {
            try
            {
                lock (client.GetClientManager())
                {
                    Pt_LogIn ptLogIn = new Pt_LogIn(databyte);
                    string strID = Encoding.Default.GetString(ptLogIn.GetID());
                    string strPW = Encoding.Default.GetString(ptLogIn.GetPW());

                    int index = 0;
                    while (strID[index++] != '\0') ;
                    strID = strID.Substring(0, index - 1);

                    index = 0;
                    while (strPW[index++] != '\0') ;
                    strPW = strPW.Substring(0, index - 1);


                  /*  if (SQLDB_LogIn.GetInstance().IsLogInOK(strID, strPW) == true)
                    {
                        Pt_LogInAck ptLogInAck = new Pt_LogInAck(eLogInState.SUCESS);
                        client.Send(ptLogInAck.GetBytes(), ptLogInAck.GetSize());
                    }
                    else
                    {
                        Pt_LogInAck ptLogInAck = new Pt_LogInAck(eLogInState.FAIL);
                        client.Send(ptLogInAck.GetBytes(), ptLogInAck.GetSize());
                    }
                    */
                    Pt_LogInAck ptLogInAck = new Pt_LogInAck(eLogInState.SUCESS);
                    client.Send(ptLogInAck.GetBytes(), ptLogInAck.GetSize());
                    Console.WriteLine("From {0} LOGIN ID {1} PW {2}", client.GetSocket().Handle, strID, strPW);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message.ToString());
            }
        }
    }
}
