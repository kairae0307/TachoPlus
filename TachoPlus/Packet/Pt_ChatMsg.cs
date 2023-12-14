using System;
using System.Collections.Generic;
using System.Text;

namespace TachoPlus
{
    enum eChatState
    {
        ROOM = 0,
        SERVER,
        ALL_SERVER,
        END_STATE
    }

    class Pt_ChatMsg : Pt_General
    {
        public static int iMaxMsgLen = 256;

        ushort mChatState;
        byte[] mChatMsg = new byte[iMaxMsgLen];

        public Pt_ChatMsg(byte[] databyte)
        {
            //아래에도 중복되는 코드가 있다. 생성자에서 패킷 사이즈와 종류를 정하는 로직을 정리할 필요가 있다.
            size = default_size + 2 + 256; //default_size + mChatState + mChatMsg
            kind = 2;

            PutBytes(databyte);
        }

        public void SetChatState(eChatState o)
        {
            mChatState = (ushort)o;
        }

        public void SetChatMsg(byte[] msg)
        {
            if (msg.Length <= iMaxMsgLen)
                mChatMsg = msg;
        }

        public eChatState GetChatState()
        {
            return (eChatState)mChatState;
        }

        public byte[] GetChatMsg()
        {
            return mChatMsg;
        }

        public Pt_ChatMsg(byte[] _Msg, eChatState _MsgState)
        {
            if (_Msg.Length <= iMaxMsgLen)
            {
                size = default_size + 2 + 256; //default_size + mChatState + mChatMsg
                kind = 2;
                mChatMsg = _Msg;
                mChatState = (ushort)_MsgState;
            }
        }

        public override void PutBytes(byte[] databyte)
        {
            mChatState = (ushort)((databyte[5] * 256) + databyte[4]);
            for (int i = 0; i < iMaxMsgLen; i++)
            {
                mChatMsg[i] = databyte[5 + 1 + i];
            }
        }

        public override byte[] GetBytes()
        {
            databyte = new byte[size];
            for (int i = 0; i < size; i++)
                databyte[i] = 0;

            int ChatState1, ChatState2;
            int Size1, Size2;

            Size1 = size % 256;
            Size2 = size / 256;

            ChatState1 = mChatState % 256;
            ChatState2 = mChatState / 256;

            databyte[0] = (byte)Size1;
            databyte[1] = (byte)Size2;
            databyte[2] = (byte)kind;
            databyte[3] = 0;
            databyte[4] = (byte)ChatState1;
            databyte[5] = (byte)ChatState2;

            for (int i = 0; i < mChatMsg.Length; i++)
            {
                databyte[5 + 1 + i] = mChatMsg[i];
            }

            return databyte;
        }
    }
}
