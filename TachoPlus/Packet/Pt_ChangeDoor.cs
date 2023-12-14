using System;
using System.Collections.Generic;
using System.Text;

namespace TachoPlus
{
    enum eDoorState
    {
        OPEN = 0,
        SUSTAIN,
        CLOSE,
        END_STATE
    }

    class Pt_ChangeDoor : Pt_General
    {
        ushort      DoorNum;
        ushort  ChangeState;

        public Pt_ChangeDoor(byte[] databyte)
        {
            PutBytes(databyte);
        }

        public Pt_ChangeDoor(short _DoorNum, eDoorState _ChangeState)
        {
            size = default_size + 2 + 2; //default_size + DoorNum + ChangeState
            kind = 1;
            DoorNum = (ushort)_DoorNum;
            ChangeState = (ushort)_ChangeState; 
        }

        public int GetDoorNum()
        {
            return DoorNum;
        }

        public int GetChangeState()
        {
            return ChangeState;
        }

        public override void PutBytes(byte[] databyte)
        {

            DoorNum     = (ushort)((databyte[5] * 256) + databyte[4]);
            ChangeState = (ushort)((databyte[7] * 256) + databyte[6]);
        }

        public override byte[] GetBytes()
        {
            databyte = new byte[size];

            int DoorNum1, DoorNum2;
            int ChangeState1, ChangeState2;

            DoorNum1 = DoorNum % 256;
            DoorNum2 = DoorNum / 256;

            ChangeState1 = ChangeState % 256;
            ChangeState2 = ChangeState / 256;


            databyte[0] = (byte)size;
            databyte[1] = 0;
            databyte[2] = (byte)kind;
            databyte[3] = 0;
            databyte[4] = (byte)DoorNum1;
            databyte[5] = (byte)DoorNum2;
            databyte[6] = (byte)ChangeState1;
            databyte[7] = (byte)ChangeState2;

            return databyte;
        }
    }
}
