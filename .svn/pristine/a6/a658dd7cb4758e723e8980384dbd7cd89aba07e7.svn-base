using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EWS_NTLM_MAIL.mail
{
    class MailUtil
    {



        //itemId 转mailMessage
        public static List<EmailMessage> BatchGetEmailIlems(ExchangeService service, Collection<ItemId> itemid)
        {
            //绑定属性
            PropertySet propertySet = new PropertySet(BasePropertySet.IdOnly,
                                                       EmailMessageSchema.InternetMessageId,//Message Id 
                                                       EmailMessageSchema.InternetMessageHeaders,//Message Hesder

                                                       EmailMessageSchema.HasAttachments,
                                                       EmailMessageSchema.Attachments,
                                                       EmailMessageSchema.ConversationIndex,//Thread-Inder
                                                       EmailMessageSchema.ConversationTopic,//Thread-Topic
                                                       EmailMessageSchema.From,//From
                                                       EmailMessageSchema.ToRecipients,		//To
                                                       EmailMessageSchema.CcRecipients,		//Cc
                                                       EmailMessageSchema.ReplyTo,				//ReplayTo
                                                       EmailMessageSchema.InReplyTo,			//In-Replay-To
                                                       EmailMessageSchema.References,			//References
                                                       EmailMessageSchema.Subject,				//Subject
                                                       EmailMessageSchema.TextBody,
                                                       EmailMessageSchema.Attachments,
                                                       EmailMessageSchema.DateTimeSent,		//Sent Date
                                                       EmailMessageSchema.DateTimeReceived, //Received Date
                                                       EmailMessageSchema.Sender,
                                                       EmailMessageSchema.MimeContent,//导出到eml文件
                                                       EmailMessageSchema.IsRead,
                                                       EmailMessageSchema.EffectiveRights,
                                                       EmailMessageSchema.ReceivedBy);//邮件已读状态

            ServiceResponseCollection<GetItemResponse> respone = service.BindToItems(itemid, propertySet);
            List<EmailMessage> messageItems = new List<EmailMessage>();
            foreach (GetItemResponse getItemRespone in respone)
            {
                try
                {
                    Item item = getItemRespone.Item;
                    EmailMessage message = (EmailMessage)item;

                    messageItems.Add(message);
                }
                catch (Exception ex)
                {
                    // Console.WriteLine("Exception with getting a message:{0}", ex.Message);
                    throw;
                }
            }
            return messageItems;
        }

        public static Collection<Folder> getAllFolder(ExchangeService service)
        {
            FolderView view = new FolderView(int.MaxValue);

            // ExtendedPropertyDefinition isHiddenProp = new ExtendedPropertyDefinition(0x10f4, MapiPropertyType.Boolean);
            // extended property.
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName,
                                                                       FolderSchema.Id,
                                                                       FolderSchema.ParentFolderId,
                                                                       FolderSchema.PolicyTag,
                                                                       FolderSchema.TotalCount,
                                                                       FolderSchema.UnreadCount,
                                                                       FolderSchema.FolderClass
                                                                       );

            view.Traversal = FolderTraversal.Shallow;
            FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.MsgFolderRoot, view);
            return findFolderResults.Folders;
        }


        //得到EmailMessage list
        public static List<EmailMessage> receiveItem(FindItemsResults<Item> findResults, string path,ExchangeService service)
        {
            int count = 0;
            Collection<ItemId> itemIds = new Collection<ItemId>();
            List<EmailMessage> messageList = new List<EmailMessage>();
            if (findResults != null)
            {
                if (findResults.Items != null && null != findResults.Items && findResults.Items.Count != 0)
                {
                    foreach (Item item in findResults.Items)
                    {
                        itemIds.Add(item.Id);
                    }

                    messageList = BatchGetEmailIlems(service, itemIds);

                }
            }
            return messageList;
        }


        //将邮件list写入eml,返回收取成功邮件数量
        public static int saveEmailList(List<EmailMessage> EmailList, string path)
        {
            Util.createDirectionary(path);
            int count = 0;
            if (EmailList != null && EmailList.Count != 0)
            {
                foreach (EmailMessage message in EmailList)
                {
                    try
                    {
                        SaveAsEml(message, path);
                        count++;
                    }
                    catch (IOException ex)
                    {
                        Console.WriteLine(ex);
                    }
                }
            }
            return count;
        }


        //将邮件写入eml
        public static string SaveAsEml(EmailMessage message, string path)
        {
            string filedirect = "";

            if (null != message && null != message.DateTimeReceived && null != message.Subject)
            {

                string filePath = "\\" + "subject" + "[" + message.Subject + "]" + Util.convertFromDate(message.DateTimeReceived) + ".eml";//根据接收时间给文件名赋值
                filePath = Util.removeAllSpecialChar(filePath);
                filedirect = @path + filePath;//邮件名称=messagesubject+date
                try
                {
                    if (!File.Exists(path + filePath))
                    {
                        if (message != null && null != message.MimeContent)
                        {
                            //Console.WriteLine(message.EffectiveRights.ToString());
                            FileStream fs = new FileStream(filedirect, FileMode.Create, FileAccess.Write);

                            fs.Write(message.MimeContent.Content, 0, message.MimeContent.Content.Length);
                            fs.Close();

                        }
                    }

                }
                catch (IOException ex)
                {
                    throw ex;
                }
            }
            return filedirect;
        }

        //遍历文件夹（两层遍历，有时间想想递归怎么实现）
        public static void travelFolder(Folder folder,List<string>folderNameList)
        {
            string root = folder.DisplayName;
            List<string> paths = new List<string>();
            FindFoldersResults result = folder.FindFolders(new FolderView(int.MaxValue));

            foreach (Folder folderChild in result.Folders)
            {
                string parentPath = root + "\\" + folderChild.DisplayName;
                folderNameList.Add(parentPath);
                //Console.Write("{0,-20}",parentPath );
                // Console.Write("{0,-10}", folderChild.TotalCount);
                //Console.WriteLine();
                //paths.Add(parentPath);
                FindFoldersResults foldchildren = folderChild.FindFolders(new FolderView(int.MaxValue));
                foreach (Folder foldSubChild in foldchildren)
                {
                    string subpath = root + "\\" + folderChild.DisplayName + "\\" + foldSubChild.DisplayName;
                    //Console.Write( "{0,-20}",subpath);
                    //Console.Write("{0,-10}", foldSubChild.TotalCount);
                    folderNameList.Add(subpath);
                    paths.Add(subpath);
                }
            }

        }
        public static string outputAllFounder(List<string> flderNames)
        {
            string folderNameStirng = "";
            foreach (string folderName in flderNames)
            {
                folderNameStirng += folderName + "\n";
                

            }

            return folderNameStirng;
        }



        //递归遍历得到folder
        public static FolderId visit(string path,ExchangeService service)
        {
            string[] folderPath = Util.getArr(path);
            //string[] folderPath = { "inbox2", "box2" };
            FolderId folderid = null;//递归得到的folderid
            // service.FindItems(new FolderId(""), new ItemView(int.MaxValue));
            FolderView view = new FolderView(int.MaxValue);
            view.Traversal = FolderTraversal.Shallow;
            ExtendedPropertyDefinition isHiddenProp = new ExtendedPropertyDefinition(0x10f4, MapiPropertyType.Boolean);
            FindFoldersResults findFolderResults = service.FindFolders(WellKnownFolderName.MsgFolderRoot, view);
            FindFoldersResults result = travel(findFolderResults, folderPath, 0);
            if (null != result && result.Folders != null && result.Folders.Count != 0)
            {
                folderid = result.Folders[0].ParentFolderId;
            }
            return folderid;
        }


        //递归得到想要收取邮件的文件夹
        public static FindFoldersResults travel(FindFoldersResults finderFolderResult, string[] arr, int index)
        {
            FindFoldersResults folders = null;
            if (index == arr.Length)
            {

                return finderFolderResult;
            }
            foreach (Folder folder in finderFolderResult.Folders)
            {
                if (folder.DisplayName.Equals(arr[index]))
                {
                    folders = folder.FindFolders(new FolderView(int.MaxValue));
                }
            }
            return travel(folders, arr, index + 1);
        }

        



    }
}
