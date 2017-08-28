using EWS_NTLM_MAIL.validate;
using EWSAPINTLM;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS_NTLM_MAIL.mail.mailservice
{
    class SearchMail
    {
        public CommandLineArgumentParser arguments;

        private Mail mail;

        public Mail Mail
        {
            get { return mail; }
            set { mail = value; }
        }

       
         public string execute(string []args, bool isHash)
        {
            arguments = CommandLineArgumentParser.Parse(args);
            int EMLCount = 0;
            if (arguments.Has("-Time"))
            {
                EMLCount = searchByTime(args, isHash);
             }
            else
            {
                EMLCount = searchMail(args, isHash);
            }
            return Util.getOutputstring(EMLCount);
        }
       
        public int searchMail(string[] args, bool isHash)
        {
            //命令解析器
             arguments = CommandLineArgumentParser.Parse(args);
             int EmlCount = 0;
            if (isHash)
            {
                string url = arguments.Get("-url").Next;
                string username = arguments.Get("-username").Next;
                string hash = arguments.Get("-hash").Next;
               this.Mail=new Mail(url,username,hash,ExchangeVersion.Exchange2013);
          }
            else
            {
                string url = arguments.Get("-url").Next;
                string username=arguments.Get("-username").Next;
                string password=arguments.Get("-password").Next;
                this.Mail=new Mail(url,username,password);
            }
            //根据参数得到过滤条件
            List<SearchFilter> filterList = getSearchFilterList(args);
            Dictionary<string, string> map = getFilterMap("-Search");
            string path=mail.Path+"\\"+getPartPath(map);
            
            if (map["ToRecipien"] == null ||!map["ToRecipien"].Equals("")||map["ToRecipien"].Equals(" "))
            {
               EmlCount= mail.searchMail(filterList, path);
            }
            else
            {
               EmlCount= mail.searchMailByToRecipient(filterList, map["ToRecipien"],path);
            }
            return EmlCount;
        }

        public int searchByTime(string[] args,bool isHash)
        {
            if (isHash)
            {
                string url = arguments.Get("-url").Next;
                string username = arguments.Get("-username").Next;
                string hash = arguments.Get("-hash").Next;
                this.Mail = new Mail(url, username, hash, ExchangeVersion.Exchange2013);
            }
            else
            {
                string url = arguments.Get("-url").Next;
                string username = arguments.Get("-username").Next;
                string password = arguments.Get("-password").Next;
                this.Mail = new Mail(url, username, password);
            }
            int EmlCount = 0;
            List<SearchFilter> filter = new List<SearchFilter>();
            string time=arguments.Get("-Time").Next.ToString();
            filter.Add(Filters.filterByTimeRange(time));
            string path = mail.Path + "\\" + "[Time]" + Util.removeAllSpecialChar(time);
            EmlCount=mail.searchMail(filter, path);
            return EmlCount;
        }
        
        public List<SearchFilter> getSearchFilterList(string[] args)
        {
            List<SearchFilter> filterList = new List<SearchFilter>();
            Dictionary<string, string> map = new Dictionary<string, string>();
             arguments = CommandLineArgumentParser.Parse(args);
            if (arguments.Has("-SearchMail"))
            {
                map = getFilterMap("-SeaarchFilter");
                if(map["subject"]!=null&&!map["subject"].Equals(""))
                {
                    filterList.Add(Filters.filterBySubject(map["subject"]));
                }
                if (map["From"] != null && !map["From"].Equals(""))
                {
                    filterList.Add(Filters.filterBySubject(map["From"]));
                }
            }
            return filterList;
        }

        //从字符串截取过滤条件字符串
        public string getSubstring(string str)
        {
            str = System.Text.RegularExpressions.Regex.Replace(str, "[ ]", "");
            int index = str.IndexOf(":") + 1;
            //Console.WriteLine(index);
            string newstring = str.Substring(index, str.Length - index);
            return newstring;
        }


          public Dictionary<string,string> getFilterMap(string command)
        {
              Dictionary<string ,string> map=new Dictionary<string,string>();
            if (arguments.Has(command))
            {
                var next = arguments.Get(command).Next;
                var p = next;
                while (true)
                {
                    if (null == p || p.Equals("") || null == p.Next)
                        break;
                    if (p.ToString().Contains("subject"))
                    {
                        map.Add("subject", getSubstring(p));
                    }
                    else if (p.ToString().Contains("From"))
                    {
                        map.Add("From", getSubstring(p));
                    }
                    else if (p.ToString().Contains("ToRecipient"))
                    {
                        map.Add("ToRecipien",getSubstring(p));
                    }
                    p = next;
                    next = next.Next;

                }
            }
              return map;
          }

        public string getPartPath(Dictionary<string,string> map)
          {
            string  path="";
            if (null != map)
            {
                if (map["subject"] != null && !map["subject"].Equals("") && !map["subject"].Equals(" "))
                path+="[subject]"+map["subject"];
                if (map["From"] != null && !map["From"].Equals("") && !map["From"].Equals(" "))
                    path+="[From]"+map["From"];
                if (map["ToRecipien"] != null && !map["ToRecipien"].Equals("") && !map["ToRecipien"].Equals(" "))
                    path += "[ToRecipien]" + map["ToRecipien"];
            }
            return path;
          }

      
        
    }
}
