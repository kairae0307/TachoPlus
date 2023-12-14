using System;
using System.Collections.Generic;
using System.Text;

namespace TachoPlus
{
    enum ePt_Type
    {
        PT_GENERAL_ACK,
        PT_CHANGE_DOOR,
        PT_CHAT_MSG,
        PT_LOGIN,
        PT_LOGIN_ACK,
        END_TYPE
    }

    abstract class Pt_General
    {
        protected           ushort size;
        protected           ushort kind;
        protected byte[] databyte;
        protected const ushort default_size = 4;
        public abstract void PutBytes(byte[] databyte);
        public    abstract  byte[] GetBytes();

        public ushort GetSize()
        {
            return size;
        }

        public ushort GetKind()
        {
            return kind;
        }
    }
}
