using System;
using Microsoft.VisualBasic;
using Vertex.Fsd.Omiga.VisualBasicPort;
namespace omLockManager
{
    public class ConvertAssistEx
    {

        public static void StringToByteArray(byte[] rbytByteArray, string vstrText) 
        {
            string strFunctionName = String.Empty;
            int lngTextLength = 0;
            int lngIndex = 0;
            // header ----------------------------------------------------------------------------------
            // description:
            // copies each character in vstrText into the relevant position of rbytByteArray
            // pass:
            // rbytByteArray()  array of bytes into which the text characters are to be transferred
            // strText          string containing the text to be transferred into the byte array
            // Raise Errors:
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO StringToByteArrayVbErr is not supported
            try 
            {
                strFunctionName = "StringToByteArray";
                lngTextLength = Strings.Len(vstrText);
                // Check string plus null terminator will fit into the byte array
                if (lngTextLength > Information.UBound(rbytByteArray, 1))
                {
                    errAssistEx.errThrowError("ConvertAssist", Convert.ToInt16(
                        strFunctionName), Convert.ToString(
                        OMIGAERROR.oeInvalidParameter));
                }

                for(lngIndex = 1; lngIndex <= lngTextLength; lngIndex += 1)
                {
                    rbytByteArray[lngIndex - 1] = Convert.ToByte(Strings.Asc(vstrText.Substring(lngIndex - 1, 1)));
                }
                return ;
                // WARNING: StringToByteArrayVbErr: is not supported 
            }
            catch(Exception exc)
            {
                // re-raise error for business object to interpret as appropriate
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }

        public static string ByteArrayToString(byte[] rbytByteArray) 
        {
            string strFunctionName = String.Empty;
            int lngArrayLimit = 0;
            string strString = String.Empty;
            int lngIndex = 0;
            // header ----------------------------------------------------------------------------------
            // description:
            // creates a string representation of a byte array
            // pass:
            // rbytByteArray()  array of bytes from which the string is to be formed
            // Raise Errors:
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO ByteArrayToStringVbErr is not supported
            try 
            {
                strFunctionName = "ByteArrayToString";
                lngArrayLimit = Information.UBound(rbytByteArray, 1);
                lngIndex = 0;
                while(lngIndex <= lngArrayLimit && rbytByteArray[lngIndex] != Convert.ToByte(0))
                {
                    strString = strString + (char)(rbytByteArray[lngIndex]);
                    lngIndex = lngIndex + 1;
                } 
                return strString;
                // WARNING: ByteArrayToStringVbErr: is not supported 
            }
            catch(Exception exc)
            {
                // re-raise error for business object to interpret as appropriate
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }

        public static DateTime CSafeDate(object vvntExpression) 
        {
            string strFunctionName = String.Empty;
            DateTime dteRetValue;
            // header ----------------------------------------------------------------------------------
            // description:
            // creates a date representation of an expression
            // pass:
            // vvntExpression  Expression to be converted to a date
            // Returns:
            // CSafeDate   Converted Expression
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO CSafeDateVbErr is not supported
            try 
            {
                strFunctionName = "CSafeDate";
                if (Strings.Len(vvntExpression) > 0)
                {
                    dteRetValue = Convert.ToDateTime(vvntExpression);
                }
                return dteRetValue;
                // WARNING: CSafeDateVbErr: is not supported 
            }
            catch(Exception exc)
            {
                // re-raise error for business object to interpret as appropriate
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }

        public static int CSafeLng(object vvntExpression) 
        {
            string strFunctionName = String.Empty;
            int lngRetValue = 0;
            // header ----------------------------------------------------------------------------------
            // description:
            // creates a long representation of an expression
            // pass:
            // vvntExpression  Expression to be converted to a long
            // Returns:
            // CSafeLng   Converted Expression
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO CSafeLngVbErr is not supported
            try 
            {
                strFunctionName = "CSafeLng";
                if (Strings.Len(vvntExpression) > 0)
                {
                    lngRetValue = Convert.ToInt32(vvntExpression);
                }
                return lngRetValue;
                // WARNING: CSafeLngVbErr: is not supported 
            }
            catch(Exception exc)
            {
                // re-raise error for business object to interpret as appropriate
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }

        public static double CSafeDbl(object vvntExpression) 
        {
            string strFunctionName = String.Empty;
            double dblRetValue = 0;
            // header ----------------------------------------------------------------------------------
            // description:
            // creates a double representation of an expression
            // pass:
            // Expression  Expression to be converted to a double
            // Returns:
            // CSafeDbl   Converted Expression
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO CSafeDblVbErr is not supported
            try 
            {
                strFunctionName = "CSafeDbl";
                if (Strings.Len(vvntExpression) > 0)
                {
                    dblRetValue = Convert.ToDouble(vvntExpression);
                }
                return dblRetValue;
                // WARNING: CSafeDblVbErr: is not supported 
            }
            catch(Exception exc)
            {
                // re-raise error for business object to interpret as appropriate
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }

        public static byte CSafeByte(object vvntExpression) 
        {
            string strFunctionName = String.Empty;
            byte bytRetValue = 0;
            // header ----------------------------------------------------------------------------------
            // description:
            // creates a byte representation of an expression
            // pass:
            // Expression  Expression to be converted to a byte
            // Returns:
            // CSafeByte   Converted Expression
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO CSafeByteVbErr is not supported
            try 
            {
                strFunctionName = "CSafeByte";
                if (Strings.Len(vvntExpression) > 0)
                {
                    bytRetValue = Convert.ToByte(Strings.Asc(Convert.ToString(vvntExpression)));
                }
                return bytRetValue;
                // WARNING: CSafeByteVbErr: is not supported 
            }
            catch(Exception exc)
            {
                // re-raise error for business object to interpret as appropriate
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }

        public static short CSafeInt(object vvntExpression) 
        {
            string strFunctionName = String.Empty;
            short bytRetValue = 0;
            // header ----------------------------------------------------------------------------------
            // description:
            // creates an integer representation of an expression
            // pass:
            // Expression  Expression to be converted to an integer
            // Returns:
            // CSafeInt   Converted Expression
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO CSafeIntVbErr is not supported
            try 
            {
                strFunctionName = "CSafeInt";
                if (Strings.Len(vvntExpression) > 0)
                {
                    bytRetValue = (short)(Convert.ToInt32(vvntExpression));
                }
                return bytRetValue;
                // WARNING: CSafeIntVbErr: is not supported 
            }
            catch(Exception exc)
            {
                // re-raise error for business object to interpret as appropriate
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }

        public static bool CSafeBool(object vvntExpression) 
        {
            string strFunctionName = String.Empty;
            bool bytRetValue = false;
            // header ----------------------------------------------------------------------------------
            // description:
            // creates an boolean representation of an expression
            // pass:
            // Expression  Expression to be converted to a boolean
            // Returns:
            // CSafeInt   Converted Expression
            // ------------------------------------------------------------------------------------------
            // WARNING: On Error GOTO CSafeBoolVbErr is not supported
            try 
            {
                strFunctionName = "CSafeBool";
                if (Strings.Len(vvntExpression) > 0)
                {
                    bytRetValue = Convert.ToBoolean(vvntExpression);
                }
                return bytRetValue;
                // WARNING: CSafeBoolVbErr: is not supported 
            }
            catch(Exception exc)
            {
                // re-raise error for business object to interpret as appropriate
                Information.Err().Raise(Information.Err().Number, Information.Err().Source, Information.Err().Description, null, null);
            }
        }


    }

}
