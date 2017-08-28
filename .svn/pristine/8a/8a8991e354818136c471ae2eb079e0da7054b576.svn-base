using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System.Collections.ObjectModel;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using System.Net;
using System.IO;
using System.Threading;

using System.Runtime.InteropServices;

using EWS_NTLM_MAIL;
using EWS_NTLM_MAIL.mail;
namespace EWSAPINTLM
{
    class Mail
    {

        static Mail()
        {
            // 获取验证证书的回调函数
            ServicePointManager.ServerCertificateValidationCallback += RemoteCertificateValidate;
        }




        private ExchangeService service;
        private string url;
        private string path;
        private string username;
        private string password;

        public string Password
        {
            get { return password; }
            set { password = value; }
        }

        public string Username
        {
            get { return username; }
            set { username = value; }
        }


        public string Path
        {
            get { return path; }
            set { path = value; }
        }

        public string Url
        {
            get { return url; }
            set { url = value; }
        }

        public ExchangeService Service
        {
            get { return service; }
            set { service = value; }
        }


       

        //哈希登陆初始化
        public Mail(String url, string userName, string hash, ExchangeVersion version)
        {
            this.url = url;
            this.username = userName;
            service = new ExchangeService(version);
            // service.Credentials = new WebCredentials("hui.xu","xhXH159519", "netseclab");
            //service.UseDefaultCredentials = true;
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.DebugMessage;
            service.Url = new Uri("url");
            //初始化创建路径
            this.path = Util.createDirectionary("E:\\" + userName + "@" + "netseclab.com");
        }


        //用户名密码登陆初始化
        public Mail(string url, String Username, string password)
        {
            this.url = url;
            this.username = Username;
            this.password = password;
            service = new ExchangeService(ExchangeVersion.Exchange2013);
            service.Credentials = new WebCredentials("hui.xu", "xhXH159519", "netseclab");
            //service.UseDefaultCredentials = true;
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.DebugMessage;
            service.Url = new Uri("https://192.168.1.120/EWS/Exchange.asmx");
            //初始化创建路径
            this.path = Util.createDirectionary("E:\\" + Username + "@" + "netseclab.com");
        }

        //回调
        private static bool RemoteCertificateValidate(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors error)
        {
            //为了通过证书验证，总是返回true
            return true;
        }

        //列出所有的文件夹
        public string listFolders()
        {
            int count = 0;
            List<string> folderNameList=new List<string>();
            Collection<Folder> folders=MailUtil.getAllFolder(service);
            foreach (Folder folder in folders)
            {
                if (folder.FolderClass == "IPF.Note")
                {
                    folderNameList.Add(folder.DisplayName);
                    //Console.Write("{0,-20}", folder.DisplayName);
                    // Console.Write("{0,-10}", folder.TotalCount);
                    //Console.WriteLine();
                    MailUtil.travelFolder(folder,folderNameList);
                    count++;
                }
            }
            return MailUtil.outputAllFounder(folderNameList);

        }

        public int getMail(string folderPath, string path)
        {
            int getEmlCount = 0;
           
            Collection<Folder> folders = MailUtil.getAllFolder(service);
            List<EmailMessage> messageList = new List<EmailMessage>();
            FindItemsResults<Item> findResults = null;
            //收取所有
            if (folderPath==null||folderPath.Equals("")||folderPath.Equals(" "))
            {
                MailUtil.getAllFolder(service);
                foreach (Folder folder in folders)
                {
                    findResults = service.FindItems(folder.Id, new ItemView(int.MaxValue));
                    messageList = MailUtil.receiveItem(findResults, path, service);
                    if (null != messageList && messageList.Count != 0)
                    {
                        getEmlCount += MailUtil.saveEmailList(messageList, path);
                    }
                }
            }
            //根据文件夹路径收取    
            else
            {
                FolderId folderid = MailUtil.visit(folderPath, service);
                if (folderid == null)
                    return 0;
                findResults = service.FindItems(folderid, new ItemView(int.MaxValue));
                messageList = MailUtil.receiveItem(findResults, path,service);
                if (null != messageList && messageList.Count != 0)
                {
                    getEmlCount =MailUtil.saveEmailList(messageList, path);
                }
            }
            return getEmlCount;
        }


        public int searchMail(List<SearchFilter> filterList,string path)
        {
            List<EmailMessage> messageList=getsearchMailList(filterList);
            if (null == messageList || messageList.Count == 0)
                return 0;
            return MailUtil.saveEmailList(messageList,path);
        }
       

        //ToRecipient ewsmapi没有相关的过滤,，扩展实现
        public int searchMailByToRecipient(List<SearchFilter> filterList, string ToRecipient, string path)
        {
            List<EmailMessage> messageList = getsearchMailList(filterList);
            List<EmailMessage> filterMessgelist = new List<EmailMessage>();
          if(null!=messageList&&messageList.Count!=0)
          {
              foreach (EmailMessage message in messageList)
              {
                  if (message != null && message.ToRecipients != null)
                  {
                      for (int i = 0; i < message.ToRecipients.Count; i++)
                      {
                          if (message.ToRecipients[i].Address.ToString().Equals(ToRecipient))
                          {
                              filterMessgelist.Add(message);

                         }
                     }
                  }
              }
            }
            return MailUtil.saveEmailList(filterMessgelist,path);
            
        }

        public List<EmailMessage> getsearchMailList(List<SearchFilter> filterList)
        {
            SearchFilter[] filterArray = filterList.ToArray();
            SearchFilter filter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, filterArray);
            List<EmailMessage> messageList = new List<EmailMessage>();
            Collection<Folder> folders = MailUtil.getAllFolder(service);
            if (null == folders || folders.Count == 0)
                return null;
            foreach (Folder folder in folders)
            {
                FindItemsResults<Item> findResults = service.FindItems(folder.Id, filter, new ItemView(int.MaxValue));
                messageList = MailUtil.receiveItem(findResults, path, service);
            }
            return messageList;
        }

              






    }
}
