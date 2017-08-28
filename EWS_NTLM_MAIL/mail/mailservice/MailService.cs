using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS_NTLM_MAIL.mail.mailservice
{
    public interface MailService
    {
         string execute(string args,bool isHash);
    }  
}
