using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS_NTLM_MAIL.enumSpace
{
    class Command
    {
        public string COMMAND_URL = "-url";
        public string COMMAND_USERNAME = "-username";
        public string COMMAND_HASH = "-hash";
        public string COMMMAND_LISTFOLD = "-listFolder";
        public string COMMAND_GETMAIL = "-GetMail  FolderName|All";
        public string COMMAND_TIME = "-Time";
        public string COMMAND_SEARCHMAIL = "-SearchMail subject:xxx | From:xxx | ToRecipient:xxx";
        private Dictionary<string, string> commandMap = new Dictionary<string, string>();

        public Dictionary<string, string> CommandMap
        {
            get { return commandMap; }
            set { commandMap = value; }
        }



        public Command()
        {
            commandMap.Add(COMMAND_URL, "EWS的wsdl地址");
            commandMap.Add(COMMAND_USERNAME, "用户名");
            commandMap.Add(COMMAND_HASH, "NTLMhash");
            commandMap.Add(COMMMAND_LISTFOLD, "显示所有邮件文件夹");
            commandMap.Add(COMMAND_GETMAIL, "获取邮件");
            commandMap.Add(COMMAND_TIME, "获取邮件的起始时间");
            commandMap.Add(COMMAND_SEARCHMAIL, "根据条件获取时间");

        }

        //输出所有的命令
        public string outputCommand()
        {
            string printString = "";
            foreach (KeyValuePair<string, string> pair in this.CommandMap)
            {
                printString += pair.Key + fillString(pair.Key) + "  " + pair.Value + "\n";
            }
            return printString;

        }

        public int getMostLength()
        {
            int mostleng = 0;
            foreach (KeyValuePair<string, string> pair in this.CommandMap)
            {
                int len = AdUtil.DisplayLength(pair.Key);
                mostleng = mostleng > len ? mostleng : len;
            }
            return mostleng;
        }
        public string fillString(string str)
        {
            string fillstring = "";
            int numBlank = getMostLength() - str.Length;
            for (int i = 0; i < numBlank; i++)
            {
                fillstring += " ";
            }
            return fillstring;
        }


    }

    //计算字符占得字节数
    public static class AdUtil
    {
        public static int DisplayLength(this string str)
        {
            int lengthCount = 0;
            var splits = str.ToCharArray();
            for (int i = 0; i < splits.Length; i++)
            {
                if (splits[i] == '\t')
                {
                    lengthCount += 8 - lengthCount % 8;
                }
                else
                {
                    lengthCount += Encoding.GetEncoding("GBK").GetByteCount(splits[i].ToString());
                }
            }
            return lengthCount;
        }
    }
}
