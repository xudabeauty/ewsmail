using EWS_NTLM_MAIL.enumSpace;
using EWS_NTLM_MAIL.mail.mailcommandimpl;
using EWS_NTLM_MAIL.mail.mailservice;
using EWS_NTLM_MAIL.validate;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS_NTLM_MAIL.mail.MailCommandInterface
{
    class MailApplication 
    {
        private CommandEnum commandEnumration;

        internal CommandEnum CommandEnumration
        {
            get { return commandEnumration; }
            set { commandEnumration = value; }
        }

       
        public MailApplication(string []args)
        {
            var argumetns = CommandLineArgumentParser.Parse(args);
            if (argumetns.Has("-GetMail"))
                commandEnumration = CommandEnum.getmail;
            else if (argumetns.Has("-SearchMail"))
                commandEnumration = CommandEnum.searchmail;
            else if (argumetns.Has("-Time"))
                commandEnumration = CommandEnum.timeMail;
            else
                commandEnumration = CommandEnum.asistMail;

        }

        public string getWay(CommandEnum commandEnum, string[] args, bool isHash)
        {
            string outputstring = "";
           switch ((int)commandEnum)
           {
               case 1:
                   {
                       AsistMail mService = new AsistMail();
                       outputstring=mService.execute(args, isHash);
                   };
           	break;
               case 2: 
                   { GetMail mail = new GetMail();
                   outputstring=mail.execute(args, isHash);
                   };
            break;
               case 3: 
                   {
                       SearchMail searchMail = new SearchMail();
                       outputstring=searchMail.execute(args, isHash);
                   };
                   break;;
               case 4:
                   {
                       SearchMail searchMail = new SearchMail();
                       outputstring = searchMail.execute(args, isHash);
                   };
            break;
           }
           return outputstring;

        }

        public  bool ishash(string[] args)
        {
            bool ishash = false;
            var arguments = CommandLineArgumentParser.Parse(args);
            if (arguments.Has("-url") && arguments.Has("username") && arguments.Has("-hash"))
            {
                ishash = true;
            }
            return ishash;
        }

        
    }
}
