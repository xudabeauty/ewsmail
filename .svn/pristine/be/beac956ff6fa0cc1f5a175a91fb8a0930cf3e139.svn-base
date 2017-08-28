using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS_NTLM_MAIL.mail.mailservice
{
    class Filters
    {

        //根据发件人过滤
        public static SearchFilter filterBYSender(string senderAddress)
        {
            EmailAddress sender = new EmailAddress(senderAddress);

            SearchFilter senderFilter =
                new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, sender);
            return senderFilter;
        }

        //根据收件人过滤
        public static SearchFilter filterByReciver(string fromAddress)
        {
            EmailAddress emailAddress = new EmailAddress(fromAddress);

            SearchFilter senderFilter =
                new SearchFilter.IsEqualTo(EmailMessageSchema.ReceivedBy, emailAddress);
            return senderFilter;
        }


        //根据时间范围过滤:message里的日期格式 "2017/08/17 00:00:00;

        public static SearchFilter filterByTimeRange(string beginDate)
        {
            beginDate += " 00:00:00";
            DateTime begindt = Util.convertToDateTime(beginDate);
            SearchFilter searcFilter = new SearchFilter.IsGreaterThanOrEqualTo(EmailMessageSchema.DateTimeReceived, begindt);
            return searcFilter;
        }

        //主题过滤
        public static SearchFilter filterBySubject(string subjectPart)
        {
            return new SearchFilter.ContainsSubstring(EmailMessageSchema.Subject, subjectPart);

        }

        //过滤未读邮件
        public static SearchFilter filterUnRead()
        {
            return new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false);
        }

        //改变邮件已读状态
        public static void updateIsRead(EmailMessage message, bool readStatus)
        {
            message.IsRead = readStatus;
            message.Update(ConflictResolutionMode.AlwaysOverwrite);
            EmailMessage em;

        }

        //主题发件人过滤
        public static SearchFilter filterSubFrom(string sub, string senderAddress)
        {
            SearchFilter searchFilter1 = filterBySubject(sub);
            SearchFilter searchFilter2 = filterBYSender(senderAddress);
            return new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilter1, searchFilter2);

        }

        //主题收件人过滤
        public static SearchFilter filterBySubTO(string sub, string receiveAddress)
        {
            SearchFilter searchFilter1 = filterBySubject(sub);
            SearchFilter searchFilter2 = filterByReciver(receiveAddress);
            return new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilter1, searchFilter2);

        }
        //收件人发件人过滤
        public static SearchFilter filterByFromTo(string from, string to)
        {
            SearchFilter searchFilter1 = filterBYSender(from);
            SearchFilter searchFilter2 = filterByReciver(to);
            return new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilter1, searchFilter2);
        }

        //主题收件人发件人过滤
        public static SearchFilter filterBySubFromTo(string sub, string from, string to)
        {
            SearchFilter searchFilter3 = filterBySubject(sub);
            SearchFilter searchFilter1 = filterBYSender(from);
            SearchFilter searchFilter2 = filterByReciver(to);
            return new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilter3, searchFilter1, searchFilter2);
        }
    }
}
