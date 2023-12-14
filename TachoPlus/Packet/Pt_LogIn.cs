using System;
using System.Collections.Generic;
using System.Text;

namespace TachoPlus
{
    class Pt_LogIn : Pt_General
    {
        public static int iMaxLen_ID = 10;
        public static int iMaxLen_PW = 10;

        byte[] mID = new byte[iMaxLen_ID];
        byte[] mPW = new byte[iMaxLen_PW];

        public Pt_LogIn(byte[] databyte)
        {
            PutBytes(databyte);
        }

        public Pt_LogIn(byte[] szID_, byte[] szPW_)
        {
            mID = szID_;
            mPW = szPW_;

            size = default_size + 10 + 10; //default_size + mID + mPW
            kind = 3;
        }

        public void SetID(byte[] o)
        {
            mID = o;
        }

        public void SetPW(byte[] o)
        {
            mPW = o;
        }

        public byte[] GetID()
        {
            return mID;
        }

        public byte[] GetPW()
        {
            return mPW;
        }

        public override void PutBytes(byte[] databyte)
        {
            for (int i = 0; i < iMaxLen_ID; i++)
                mID[i] = databyte[4 + i];

            for (int i = 0; i < iMaxLen_PW; i++)
                mPW[i] = databyte[4 + iMaxLen_ID + i];
        }

        public override byte[] GetBytes()
        {
            byte[] databyte = new byte[size];
            for (int i = 0; i < databyte.Length; i++)
                databyte[i] = 0;

            databyte[0] = (byte)size;
            databyte[1] = 0;
            databyte[2] = (byte)kind;
            databyte[3] = 0;

            for (int i = 0; i < mID.Length; i++)
                databyte[3 + 1 + i] = mID[i];

            for (int i = 0; i < mPW.Length; i++)
                databyte[iMaxLen_ID + 3 + 1 + i] = mPW[i];

            return databyte;
        }
    }
}
