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
    class GetMail
    {
        public CommandLineArgumentParser arguments;

        private Mail mail;

        internal Mail Mail
        {
            get { return mail; }
            set { mail = value; }
        }

        public string execute(string[] args, bool isHash)
        {
            int EMLCount = getMail(args, isHash);
            return Util.getOutputstring(EMLCount);
        }

        public int getMail(string[] args, bool isHash)
        {
            //命令解析器
            int EmlCount = 0;
            arguments = CommandLineArgumentParser.Parse(args);
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
            string path="";
            if (arguments.Get("-GetMail").Next.ToString().Equals("ALL"))
            {
                path=mail.Path + "\\folderEmails" + Util.convertFromDate(DateTime.Now);

            }
            else
            {
                path = mail.Path + "\\folderEmails" + arguments.Get("-GetMail").Next.ToString();
            }
            EmlCount=GetMailBY(args, path);
            return EmlCount;


        }
        public int  GetMailBY(string[] args, string path)
        {
            int EmlCount = 0;
            arguments = CommandLineArgumentParser.Parse(args);

            if (arguments.Get("-GetMail").Next.ToString().Equals("ALL"))
            {
            
               EmlCount =mail.getMail(null, path);
             }
            else
            {
               EmlCount= mail.getMail(arguments.Get("-GetMail").Next.ToString(),path);
            }
            return EmlCount;
        }

       

    }
}
