using EWSAPINTLM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EWS_NTLM_MAIL.enumSpace;
using EWS_NTLM_MAIL.validate;
using Microsoft.Exchange.WebServices.Data;

namespace EWS_NTLM_MAIL.mail.mailservice
{
   
    class AsistMail
    {
        Mail mail;

        internal Mail Mail
        {
            get { return mail; }
            set { mail = value; }
        }
        CommandLineArgumentParser arguments;

        //实现接口
        public string execute(string []args, bool isHash)
        {
            arguments = CommandLineArgumentParser.Parse(args);
            if (arguments.Has("-help"))
            {
                return listCommand();
           }
            else
            {
                if (arguments.Has("-listFolder"))
                {
                    return listAllFolder(args,isHash);
                }
            }
            return null;
        }

        //list所有指令
       public string listCommand()
        {
            
            return new Command().outputCommand();
        }


        public string listAllFolder(string []args,bool isHash)

       {
           string outputstring = "";
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
               string url = arguments.Get("url").Next;
               string username = arguments.Get("-username").Next;
               string password = arguments.Get("-password").Next;
               this.Mail = new Mail(url, username, password);
           }
           outputstring+= mail.listFolders();
           return outputstring;
       }

       
    }
}
