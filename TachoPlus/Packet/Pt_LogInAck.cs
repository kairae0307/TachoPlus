using System;
using System.Collections.Generic;
using System.Text;

namespace TachoPlus
{
    enum eLogInState
    {
        SUCESS = 0,
        FAIL,
        END_STATE
    }

    class Pt_LogInAck : Pt_General
    {
        ushort  mLogInState;

        public Pt_LogInAck(byte[] databyte)
        {
            PutBytes(databyte);
        }

        public Pt_LogInAck(eLogInState eLogInState_)
        {
            size = default_size + 2; //default_size + mLogInState
            kind = 4;

            mLogInState = (ushort)eLogInState_;
        }

        public void SetLogInState(eLogInState o)
        {
            mLogInState = (ushort)o;
        }

        public eLogInState GetLogInState()
        {
            return (eLogInState)mLogInState;
        }

        public override void PutBytes(byte[] databyte)
        {

            mLogInState = (ushort)((databyte[5] * 256) + databyte[4]);
        }

        public override byte[] GetBytes()
        {
            databyte = new byte[size];

            int iLogInState1, iLogInState2;

            iLogInState1 = mLogInState % 256;
            iLogInState2 = mLogInState / 256;

            databyte[0] = (byte)size;
            databyte[1] = 0;
            databyte[2] = (byte)kind;
            databyte[3] = 0;
            databyte[4] = (byte)iLogInState1;
            databyte[5] = (byte)iLogInState2;

            return databyte;
        }
    }
}
