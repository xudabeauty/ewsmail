using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS_NTLM_MAIL
{
    class Util
    {
        //日期格式2017/11/15 12:15:45
        public static DateTime convertToDateTime(string dateTimeStr)
        {

            DateTime dt = DateTime.ParseExact(dateTimeStr, "yyyy/MM/dd HH:mm:ss", System.Globalization.CultureInfo.CurrentCulture);
            return dt;
        }

        //将数组转为字符串,将前5个去掉（-url a username a -password）
        public static string connectstr(string[] arr)
        {
            string str = "";
            for (int i = 6; i < arr.Length; i++)
            {
                //if (i == arr.Length - 1)
                //    str += arr[i];
                //else
                str += arr[i] + " ";

            }
            return str;
        }

        public static string connectCommand(string[] arr)
        {
            string str = "";
            for (int i = 0; i < 5; i++)
            {
                //if (i == arr.Length - 1)
                //    str += arr[i];
                //else
                str += arr[i] + " ";

            }
            return str;
        }




        //日期转为字符串
        public static string convertFromDate(DateTime dt)
        {
            string str = "";
            str += dt.Year;
            string month = Convert.ToString(dt.Month);
            if (month.Length == 1)
            {
                str += 0 + month;
            }
            else
            {
                str += month;
            }
            string day = Convert.ToString(dt.Day);
            if (day.Length == 1)
            {
                str += 0 + day;
            }
            else
            {
                str += day;
            }
            str += connect(Convert.ToString(dt.Hour));
            str += connect(Convert.ToString(dt.Minute));
            str += connect(Convert.ToString(dt.Second));

            return str;
        }


        public static string connect(string str)
        {
            if (str.Length == 1)
            {
                return 0 + str;
            }
            else
            {
                return str;
            }

        }

        //讲字符串转为数组
        public static string[] getArr(string str)
        {
            if (null != str)
            {
                string[] arr = str.Split('\\');
                return arr;
            }
            return null;
        }

        //创建路径
        public static string createDirectionary(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            return path;
        }


        //去掉所有的！特殊符号
        public static string removeAllSpecialChar(string filename)
        {
            string[] arr = { "?", ",", "/", "*", "\"", "\"", "<", ">", "|" };
            {
                filename = System.Text.RegularExpressions.Regex.Replace(filename, "[:]", "");
                filename = System.Text.RegularExpressions.Regex.Replace(filename, "[//]", "");

                return filename;
            }

        }

        //去掉命令中的subject:

        public static  string getOutputstring(int emailcount)
        {
            return "OK!                 共收取 " + emailcount + " 封邮件";
        }



    }
}
