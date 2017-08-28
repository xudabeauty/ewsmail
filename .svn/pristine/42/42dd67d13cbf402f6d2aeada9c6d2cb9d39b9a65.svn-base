using EWS_NTLM_MAIL.enumSpace;
using EWS_NTLM_MAIL.mail.mailcommandimpl;
using EWS_NTLM_MAIL.mail.MailCommandInterface;
using EWS_NTLM_MAIL.validate;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS_NTLM_MAIL
{
    class Program
    {
        static void Main(string[] arg)
        {
            string url="https://192.168.1.120/EWS/Exchange.asmx@netseclab.com ";
            string[] args= { "-url", url, "-username", "xhXH159519", "-password", "xhXH159519", "-GetMail", "收件箱\\inbox2" };
            string[] args1 = { "-url",url, "-username", "xhXH159519", "-password", "xhXH159519", "-GetMail", "ALL" };
            string[] args2 = { "-url", url, "-username", "xhXH159519", "-password", "xhXH159519", "-help" };
            string[] args3 = { "-url", url, "-username", "xhXH159519", "-password", "xhXH159519", "-Time","2017/07/17" };
            string[] args4= { "-url", url, "-username", "xhXH159519", "-password", "xhXH159519", "-Time", "2017/08/17" };
            MailApplication application = new MailApplication(args3);
            CommandEnum command = application.CommandEnumration;
             string output=application.getWay(command, args3, application.ishash(args3));
            Console.WriteLine(output);
        }

       
    }
}
