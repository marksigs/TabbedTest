using System;
using Microsoft.VisualBasic;
using System.Runtime.InteropServices;
namespace omLockManager
{
    public class guidAssistEx
    {

        public static string CreateGUID() 
        {
            guidAssistEx.Guid guidNewGuid = new guidAssistEx.Guid(null);
            string strBuffer = String.Empty;
            // -----------------------------------------------------------------------------------------
            // description:  Creates a Global Unique Identifier
            // return:
            // ------------------------------------------------------------------------------------------
            WinCoCreateGuid(guidNewGuid);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D1), (short)8, false);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D2), (short)4, false);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D3), (short)4, false);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D4[0]), (short)2, false);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D4[1]), (short)2, false);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D4[2]), (short)2, false);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D4[3]), (short)2, false);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D4[4]), (short)2, false);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D4[5]), (short)2, false);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D4[6]), (short)2, false);
            strBuffer = PadRight0(strBuffer, Conversion.Hex(guidNewGuid.D4[7]), (short)2, false);
            return strBuffer;
        }

        private static string PadRight0(string vstrBuffer, string vstrBit, short intLenRequired, bool bHyp) 
        {
            // -----------------------------------------------------------------------------------------
            // description:
            // return:
            // ------------------------------------------------------------------------------------------
            return vstrBuffer + 
                vstrBit + 
                new String("0"[0], Math.Abs(intLenRequired - Strings.Len(vstrBit)));
        }

        [ DllImport( "OLE32.DLL.dll", CharSet = CharSet.Unicode, PreserveSig = false )]
        public static extern int WinCoCreateGuid(guidAssistEx.Guid guidNewGuid);


        public struct Guid {
            public int D1;
            public short D2;
            public short D3;
            public byte[] D4;

            public Guid(object _object_)  {
                D1 = 0;
                D2 = 0;
                D3 = 0;
                D4 = new byte[8 + 1];
            }

        }


    }

}
