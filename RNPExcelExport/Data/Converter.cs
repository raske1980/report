using System;

namespace RNPExcelExport.Data
{
    public class Converter
    {
        public static DateTime ToDateTime(object value)
        {
            if (value == null)
                return DateTime.MinValue;

            DateTime dt;
            if (DateTime.TryParse(value.ToString(), out dt))
                return dt;
            else
                return DateTime.MinValue;
        }       

        public static bool ToBoolean(object value)
        {
            if (value == null || value is System.DBNull)
                return false;

            string res = ToString(value);
            if (res.ToLower().IndexOf("false") > -1 || res.ToLower().IndexOf("true") > -1)
            {
                try
                {
                    return System.Convert.ToBoolean(value);
                }
                catch
                {
                    return false;
                }
            }

            try
            {
                return System.Convert.ToBoolean(ToInt32(value));
            }
            catch
            {
                return false;
            }
        }

        public static bool? ToBooleanNullable(object value)
        {
            if (value == null || value is System.DBNull)
                return null;

            string res = ToString(value);
            if (res.ToLower().IndexOf("false") > -1 || res.ToLower().IndexOf("true") > -1)
            {
                try
                {
                    return System.Convert.ToBoolean(value);
                }
                catch
                {
                    return null;
                }
            }

            try
            {
                return System.Convert.ToBoolean(ToInt32(value));
            }
            catch
            {
                return null;
            }
        }

        public static string ToStringNullable(object value)
        {
            if (value == null)
                return null;

            return System.Convert.ToString(value);
        }

        public static string ToString(object value)
        {
            if (value == null)
                return "";

            return System.Convert.ToString(value);
        }

        public static int ToInt32(object value)
        {
            return ToInt32(value, false);
        }

        public static int ToInt32(object value, bool isZero)
        {
            if (value == null)
                return isZero ? 0 : -1;

            int dt;
            if (int.TryParse(value.ToString(), out dt))
                return dt;
            else
            {
                try
                {
                    return System.Convert.ToInt32(value);
                }
                catch
                {
                    return isZero ? 0 : -1;
                }
            }
        }

        public static double? ToDoubleNullable(object value)
        {
            if (value == null)
                return null;

            double dt;
            if (double.TryParse(value.ToString(), out dt))
                return dt;
            else
            {
                try
                {
                    return System.Convert.ToDouble(value);
                }
                catch
                {
                    return null;
                }
            }
        }

        public static double ToDouble(object value)
        {
            return ToDouble(value, false);
        }

        public static double ToDouble(object value, bool isZero)
        {
            if (value == null)
                return isZero ? 0 : -1;

            double dt;
            if (double.TryParse(value.ToString(), out dt))
                return dt;
            else
            {
                try
                {
                    return System.Convert.ToDouble(value);
                }
                catch
                {
                    return isZero ? 0 : -1;
                }
            }
        }

        public static int? ToInt32Nullable(object value)
        {
            if (value == null)
                return null;

            int dt;
            if (int.TryParse(value.ToString(), out dt))
                return dt;
            else
            {
                try
                {
                    return System.Convert.ToInt32(value);
                }
                catch
                {
                    return null;
                }
            }
        }

        public static int? ExtractInt32Nullable(object value)
        {
            if (value == null)
                return null;

            string val = value.ToString();
            int startIndex = val.LastIndexOf('-') + 1;
            int lenght = val.Length - startIndex;
            val = val.Substring(startIndex, lenght);

            int dt;
            if (int.TryParse(val.ToString(), out dt))
                return dt;
            else
            {
                try
                {
                    return System.Convert.ToInt32(val);
                }
                catch
                {
                    return null;
                }
            }
        }

        public static float ToFloat32(object value, bool isZero)
        {
            if (value == null)
                return isZero ? 0 : -1;

            float dt;
            if (float.TryParse(value.ToString(), out dt))
                return dt;
            else
                return isZero ? 0 : -1;
        }

        public static float? ToFloatNullable(object value)
        {
            if (value == null)
                return null;

            float dt;
            if (float.TryParse(value.ToString(), out dt))
                return dt;
            else
                return null;
        }

        public static decimal ToDecimal(object value)
        {
            if (value == null)
                return 0;

            decimal d;
            if (decimal.TryParse(value.ToString(), out d))
                return d;
            else
            {
                if (decimal.TryParse(value.ToString().Replace(".", ","), out d))
                    return d;
                else
                    return 0;
            }
        }

        public static decimal? ToDecimalNullable(object value)
        {
            if (value == null)
                return null;

            decimal d;
            if (decimal.TryParse(value.ToString(), out d))
                return d;
            else
            {
                if (decimal.TryParse(value.ToString().Replace(".", ","), out d))
                    return d;
                else
                    return null;
            }
        }

        public static short? ToInt16Nullable(object value)
        {
            if (value == null)
                return null;

            short dt;
            if (short.TryParse(value.ToString(), out dt))
                return dt;
            else
                return null;
        }

        public static short ToInt16(object value)
        {
            return ToInt16(value, false);
        }

        public static short ToInt16(object value, bool isZero)
        {
            if (value == null)
                return (short)(isZero ? 0 : -1);

            short dt;
            if (short.TryParse(value.ToString(), out dt))
                return dt;
            else
                return (short)(isZero ? 0 : -1);
        }

        public static byte? ToByteNullable(object value)
        {
            if (value == null)
                return null;

            byte dt;
            if (byte.TryParse(value.ToString(), out dt))
                return dt;
            else
                return null;
        }

        public static byte ToByte(object value)
        {
            if (value == null)
                return 0;

            byte dt;
            if (byte.TryParse(value.ToString(), out dt))
                return dt;
            else
                return 0;
        }

        public static Guid? ToGuidNullable(object value)
        {
            if (value == null)
                return null;

            Guid? dt;
            try
            {
                dt = new Guid(value.ToString());
            }
            catch
            {
                dt = null;
            }

            return dt;
        }

        public static Guid ToGuid(object value)
        {
            if (value == null)
                return Guid.Empty;

            Guid dt = Guid.Empty;
            try
            {
                dt = new Guid(value.ToString());
            }
            catch
            {
                dt = Guid.Empty;
            }

            return dt;
        }

        public static DateTime? ToDateTimeNullable(object value)
        {
            if (value == null)
                return null;

            DateTime dt;
            if (DateTime.TryParse(value.ToString(), out dt))
                return dt;
            else
                return null;
        }

        public static string ToDateTimeToString(object value)
        {
            DateTime? dt = ToDateTimeNullable(value);
            if (dt == null)
                return "";
            else
                return dt.Value.ToShortDateString();
        }

        public static string ToDateTimeToStringFormat(object value)
        {
            DateTime? dt = ToDateTimeNullable(value);
            if (dt == null)
                return "";
            else
                return dt.Value.ToString("dd.MM.yyyy HH:mm");
        }

        public static Int64 ToInt64(object value)
        {
            return ToInt64(value, false);
        }

        public static Int64 ToInt64(object value, bool isZero)
        {
            if (value == null)
                return isZero ? 0 : -1;

            long dt;
            if (long.TryParse(value.ToString(), out dt))
                return dt;
            else
            {
                try
                {
                    return System.Convert.ToInt64(value);
                }
                catch
                {
                    return isZero ? 0 : -1;
                }
            }
        }

        public static Int64? ToInt64Nullable(object value)
        {
            if (value == null)
                return null;

            long dt;
            if (long.TryParse(value.ToString(), out dt))
                return dt;
            else
            {
                try
                {
                    return System.Convert.ToInt64(value);
                }
                catch
                {
                    return null;
                }
            }
        }

        public static double RoundToNearest(double Amount, double RoundTo)
        {
            double ExcessAmount = Amount % RoundTo;
            if (ExcessAmount < (RoundTo / 2))
            {
                Amount -= ExcessAmount;
            }
            else
            {
                Amount += (RoundTo - ExcessAmount);
            }

            return Amount;
        }
        
        public static decimal RoundToNearestDecimal(decimal Amount, int DigitsToRound)
        {
            return Math.Round(Amount, DigitsToRound, MidpointRounding.AwayFromZero);
        }
    }
}