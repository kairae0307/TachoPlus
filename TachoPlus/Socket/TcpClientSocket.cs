using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Windows.Forms;

namespace TachoPlus
{
    class TcpClientSocket
    {
        public static int mStreamBufferSize = 1024*66;
        public static int mTempStreamBufferSize = 1024*66;
   
        protected int mStreamBufferCurrentPoint = 0;
        protected int mStreamBufferCurrentSize = 0;

        protected Socket mFdes = null;
        protected TcpServerSocket mServer = null;
        protected TcpClientSocketManager mClientMgr = null;

        public byte[] mStreamBuffer = null;
        public byte[] mTempStreamBuffer = null;


        public TcpClientSocket()
        {
            mFdes = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            mStreamBuffer = new byte[mStreamBufferSize];
            mTempStreamBuffer = new byte[mTempStreamBufferSize];
        }

        public TcpClientSocket(Socket o)
        {
            mFdes = o;
            mStreamBuffer = new byte[mStreamBufferSize];
            mTempStreamBuffer = new byte[mTempStreamBufferSize];
        }

        ~TcpClientSocket()
        {
        }

        public void SetClientManager(TcpClientSocketManager o)
        {
            mClientMgr = o;
        }

        public TcpClientSocketManager GetClientManager()
        {
            return mClientMgr;
        }

        public void SetSocket(Socket o)
        {
            mFdes = o;
        }

        public TcpServerSocket GetServerSocket()
        {
            return mServer;
        }

        public void SetServerSocket(TcpServerSocket o)
        {
            mServer = o;
        }

        public Socket GetSocket()
        {
            return mFdes;
        }

        public void SetTempStreamBuffer(byte[] o)
        {
            mTempStreamBuffer = o;
        }

        public byte[] GetTempStreamBuffer()
        {
            return mTempStreamBuffer;
        }

        public bool Connect(string host, int port)
        {
            mFdes.Connect(host, port);

            if (mFdes.Connected)
                return true;

            return false;
        }

        public void CloseSocket()
        {
            try
            {
                mFdes.Close();
            }
            catch (Exception e)
            {
                e.Message.ToString();
            }
        }

        public int Send(byte[] o, int size)
        {
            try
            {
                return mFdes.Send(o, size, SocketFlags.None);
            }
            catch (Exception e)
            {
                Console.WriteLine("From {0} : {1}", mFdes.Handle, e.Message.ToString());
                return -1;
            }
        }

        public IAsyncResult BeginSendClient(byte[] o, int size)
        {
            try
            {
                return mFdes.BeginSend(o, 0, size, SocketFlags.None, new AsyncCallback(EndSendClient), this);
            }
            catch (Exception e)
            {
                Console.WriteLine("From {0} : {1}", mFdes.Handle, e.Message.ToString());
                return null;
            }
        }

        public void EndSendClient(IAsyncResult ar)
        {
        }

        public int Recv(byte[] o, int size)
        {
            return mFdes.Receive(o, size, SocketFlags.None);
        }

        public void BroadCast(byte[] o, int size)
        {
            mServer.BroadCast(o, size);
        }

        public bool IsAbleToRecv()
        {
            if (mTempStreamBufferSize <= (mStreamBufferSize - mStreamBufferCurrentSize))
                return true;
          
            return false;
        }

        public int RecvToStreamBuffer()
        {
            if (mFdes.Connected)
            {
                // mTempStreamBufferSize만큼의 여유 공간이 있을때만 Receive해 준다.
                if (IsAbleToRecv() == true)
                {
                    int nSize = mFdes.Receive(mTempStreamBuffer, 0, mTempStreamBufferSize, SocketFlags.None);
                    for (int i = 0; i < nSize; i++)
                    {
                        mStreamBuffer[mStreamBufferCurrentPoint] = mTempStreamBuffer[i];
                        mStreamBufferCurrentPoint++;

                        if (mStreamBufferSize <= mStreamBufferCurrentPoint)
                            mStreamBufferCurrentPoint = 0;
                    }


                  
                    mStreamBufferCurrentSize += nSize;
                    return nSize;
                }
            }

            return -1;
        }

       
        //비동기 용
        public int RecvToStreamBufferForAsync(int nSize)
        {
            try
            {
                bool Amf_Start = false;

                if (mFdes.Connected)
                {
                    for (int i = 0; i < nSize; i++)
                    {
                        mStreamBuffer[mStreamBufferCurrentPoint] = mTempStreamBuffer[i];
                        mStreamBufferCurrentPoint++;

                        if (mTempStreamBuffer[i] == 0xfd)
                        {
                            Amf_Start = true;
                        }

                        if (mStreamBufferSize <= mStreamBufferCurrentPoint)
                            mStreamBufferCurrentPoint = 0;
                    }

                    if (mStreamBuffer[0] == 0xA3)
                    {
                        
                            mServer.mStreamBuffer = mStreamBuffer;
                          /*  if (nSize != 0x2000)
                            {
                                mServer.MakeAMfFile();
                            }*/
                            if (Amf_Start == true)
                            {
                                mServer.MakeAMfFile();
                            }
                      
                    }
                    if (Amf_Start == true)
                    {
                        mStreamBufferCurrentPoint = 0;
                        mStreamBufferCurrentSize += nSize;
                    }


                    return nSize;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return -1;
        }
    }
}
    
