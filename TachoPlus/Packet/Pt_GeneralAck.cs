using System;
using System.Collections.Generic;
using System.Text;

namespace TachoPlus
{
    enum ePt_AckType
    {
        OK,
        FAIL,
        END_TYPE
    }

    class Pt_GeneralAck : Pt_General
    {
        ushort RequestType;
        ushort Result;

        public Pt_GeneralAck(byte[] databyte)
        {
            PutBytes(databyte);
        }

        public Pt_GeneralAck(ePt_Type Request, ePt_AckType AckState)
        {
            size = default_size + 2 + 2; //default_size + RequestType + Result
            kind = 0;
            RequestType = (ushort)Request;
            Result = (ushort)AckState;
        }

        public ePt_AckType GetResult()
        {
            return (ePt_AckType)Result;
        }

        public ePt_Type GetRequestType()
        {
            return (ePt_Type)RequestType;
        }

        public override void PutBytes(byte[] databyte)
        {
            RequestType = (ushort)((databyte[5] * 256) + databyte[4]);
            Result      = (ushort)((databyte[7] * 256) + databyte[6]);
        }

        public override byte[] GetBytes()
        {
            databyte = new byte[size];

            int RequestType1, RequestType2;
            int Result1, Result2;

            Result1 = Result % 256;
            Result2 = Result / 256;

            RequestType1 = RequestType % 256;
            RequestType2 = RequestType / 256;

            databyte[0] = (byte)size;
            databyte[1] = 0;
            databyte[2] = (byte)kind;
            databyte[3] = 0;
            databyte[4] = (byte)RequestType1;
            databyte[5] = (byte)RequestType2;
            databyte[6] = (byte)Result1;
            databyte[7] = (byte)Result2;

            return databyte;
        }
    }
}
