using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Data;
using Newtonsoft.Json;
using System.Globalization;

namespace wsEmailExchange.Util
{
    public static class MailUtil
    {
        static string APIUrl = ConfigurationManager.AppSettings["APIUrl"].ToString();
        static string userName = ConfigurationManager.AppSettings["UserName"].ToString();
        static string passWord = ConfigurationManager.AppSettings["Password"].ToString();
        static int pageSize = Convert.ToInt32(ConfigurationManager.AppSettings["PageSize"].ToString());
        static string localPath = ConfigurationManager.AppSettings["LocalPath"].ToString();
        static string connStr = ConfigurationManager.ConnectionStrings["ConnStr"].ConnectionString.ToString();
        const string EmailDen = "DEN";
        const string EmailDi = "DI";

        public static ExchangeService InitExchangeObj()
        {
            try
            {
                ExchangeService exchangeService = new ExchangeService();
                exchangeService.Url = new Uri(APIUrl);
                exchangeService.Credentials = new WebCredentials(userName, passWord);
                exchangeService.Timeout = 10000;
                return exchangeService;
            }
            catch (Exception ex)
            {
                Common.GhiLog("InitExchangeObj", ex);
                return null;
            }
        }

        public static void ReadMailUnreadInbox(ExchangeService exchangeService)
        {
            try
            {
                if (exchangeService == null) return;
                SearchFilter sf = new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                ItemView itemView = new ItemView(pageSize);
                PropertySet prop = new PropertySet(BasePropertySet.IdOnly);
                itemView.PropertySet = prop;
                itemView.OrderBy.Add(EmailMessageSchema.LastModifiedTime, SortDirection.Descending);
                FindItemsResults<Item> items = exchangeService.FindItems(WellKnownFolderName.Inbox, sf, itemView);

                foreach (var item in items.Items)
                {
                    EmailMessage email = EmailMessage.Bind(exchangeService, item.Id);
                    List<Attachment> attachments = email.Attachments.ToList();
                    // Email có đính kèm
                    if (attachments.Any())
                    {
                        List<AttachmentObj> lstAttachObj = SaveAttachmentsToLocal(attachments);
                        if (lstAttachObj.Any())
                        {
                            using (SqlConnection connection = new SqlConnection(connStr))
                            {
                                if (connection.State != ConnectionState.Open)
                                {
                                    connection.Open();
                                }
                                using (SqlTransaction transaction = connection.BeginTransaction())
                                {
                                    try
                                    {
                                        int id = 0;
                                        id = SaveEmailToDB(connection, email, transaction);
                                        if (id > 0)
                                        {
                                            SaveAttachmentToDB(lstAttachObj, id, connection, transaction);
                                            transaction.Commit();
                                            Common.GhiLog("Lưu email có file đính kèm thành công.", "");
                                        }
                                        else
                                        {
                                            Common.GhiLog("Lưu email có file đính kèm thất bại", id.ToString());
                                            transaction.Rollback();
                                            continue;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Common.GhiLog("Lưu email có file đính kèm thất bại", ex);
                                        transaction.Rollback();
                                        continue;
                                    }
                                }
                            }
                            UpdateStatusEmail(email);
                        }
                    }
                    // email không có file đính kèm
                    else
                    {
                        using (SqlConnection connection = new SqlConnection(connStr))
                        {
                            if (connection.State != ConnectionState.Open)
                            {
                                connection.Open();
                            }
                            using (SqlTransaction transaction = connection.BeginTransaction())
                            {
                                try
                                {
                                    SaveEmailToDB(connection, email, transaction);
                                    transaction.Commit();
                                    Common.GhiLog("Lưu email không có file đính kèm thành công", "");
                                }
                                catch (Exception ex)
                                {
                                    Common.GhiLog("Lưu email không có file đính kèm thất bại", ex);
                                    transaction.Rollback();
                                    continue;
                                }
                            }
                        }
                        UpdateStatusEmail(email);
                    }
                }
            }
            catch (Exception ex)
            {
                Common.GhiLog("ReadMailUnreadInbox", ex);
            }
        }

        private static void UpdateStatusEmail(EmailMessage emailMessage)
        {
            try
            {
                emailMessage.IsRead = true;
                emailMessage.Update(ConflictResolutionMode.AlwaysOverwrite);
                Common.GhiLog("Cap nhat IsRead mail thanh cong", emailMessage.Id.UniqueId);
            }
            catch (Exception ex)
            {
                Common.GhiLog("Cap nhat IsRead mail loi", ex);
            }
        }


        private static void SaveAttachmentToDB(List<AttachmentObj> lstAttachObj, int idEmail, SqlConnection connection, SqlTransaction transaction = null)
        {
            string sqlAttach = "INSERT INTO DFiles (TenantId,LoaiDinhKem,TenFile,LoaiFile,DungLuong,DuongDan,NgayDinhKem,Guid,IdCha) " +
                "VALUES (@TenantId,@LoaiDinhKem,@TenFile,@LoaiFile,@DungLuong,@DuongDan,GETDATE(),NEWID(),@IdCha)";
            using (SqlCommand command = new SqlCommand(sqlAttach, connection, transaction))
            {
                command.Parameters.Add(new SqlParameter("@TenantId", SqlDbType.Int) { Value = 2 });
                command.Parameters.Add(new SqlParameter("@LoaiDinhKem", SqlDbType.NVarChar, 100) { Value = "DEmails" });
                command.Parameters.Add(new SqlParameter("@TenFile", SqlDbType.NVarChar, 200));
                command.Parameters.Add(new SqlParameter("@LoaiFile", SqlDbType.NVarChar, 100));
                SqlParameter parameterDungLuong = new SqlParameter("@DungLuong", SqlDbType.Decimal);
                parameterDungLuong.Precision = 18;
                parameterDungLuong.Scale = 2;
                command.Parameters.Add(parameterDungLuong);
                command.Parameters.Add(new SqlParameter("@DuongDan", SqlDbType.NVarChar, 1000));
                command.Parameters.Add(new SqlParameter("@IdCha", SqlDbType.Int) { Value = idEmail });
                command.Prepare();
                foreach (var obj in lstAttachObj)
                {
                    command.Parameters["@TenFile"].Value = obj.TenFile;
                    command.Parameters["@LoaiFile"].Value = obj.LoaiFile;
                    command.Parameters["@DungLuong"].Value = obj.DuongLuong;
                    command.Parameters["@DuongDan"].Value = obj.DuongDan;

                    if (connection.State != ConnectionState.Open)
                    {
                        connection.Open();
                    }
                    command.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Lưu email đến database
        /// </summary>
      
        private static int SaveEmailToDB(SqlConnection connection, EmailMessage email, SqlTransaction transaction = null)
        {
            string sqlAttach = "INSERT INTO DEmails (TenantId,LoaiEmail,NguoiGui,NgayGui,NguoiNhan,NgayNhan,Cc,Bcc,Subject,Body,ConversationId,Guid,MailExchangeId) " +
                "OUTPUT Inserted.Id " +
                "VALUES (@TenantId,@LoaiEmail,@NguoiGui,@NgayGui,@NguoiNhan,@NgayNhan,@Cc,@Bcc,@Subject,@Body,@ConversationId,NEWID(),@MailExchangeId)";
            using (SqlCommand command = new SqlCommand(sqlAttach, connection, transaction))
            {
                command.Parameters.Add(new SqlParameter("@TenantId", SqlDbType.Int) { Value = 2 });
                command.Parameters.Add(new SqlParameter("@LoaiEmail", SqlDbType.NVarChar, 100) { Value = EmailDen });
                command.Parameters.Add(new SqlParameter("@NguoiGui", SqlDbType.VarChar, 100) { Value = email.From.Address });
                command.Parameters.Add(new SqlParameter("@NgayGui", SqlDbType.DateTime) { Value = email.DateTimeSent.ToString("yyyy-MM-dd HH:mm:ss", DateTimeFormatInfo.InvariantInfo) });
                command.Parameters.Add(new SqlParameter("@NguoiNhan", SqlDbType.VarChar, 500) { Value = JsonConvert.SerializeObject(email.ToRecipients.Select(x => x.Address)) });
                command.Parameters.Add(new SqlParameter("@NgayNhan", SqlDbType.DateTime) { Value = email.DateTimeReceived.ToString("yyyy-MM-dd HH:mm:ss", DateTimeFormatInfo.InvariantInfo) });
                command.Parameters.Add(new SqlParameter("@Cc", SqlDbType.NVarChar, 500) { Value = JsonConvert.SerializeObject(email.CcRecipients.Select(x => x.Address)) });
                command.Parameters.Add(new SqlParameter("@Bcc", SqlDbType.NVarChar, 500) { Value = JsonConvert.SerializeObject(email.BccRecipients.Select(x => x.Address)) });
                command.Parameters.Add(new SqlParameter("@Subject", SqlDbType.NVarChar, 1000) { Value = email.Subject });
                command.Parameters.Add(new SqlParameter("@Body", SqlDbType.NVarChar, -1) { Value = email.Body.Text });
                command.Parameters.Add(new SqlParameter("@ConversationId", SqlDbType.VarChar, 300) { Value = email.ConversationId.UniqueId });
                command.Parameters.Add(new SqlParameter("@MailExchangeId", SqlDbType.VarChar, 300) { Value = email.Id.UniqueId });

                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }
                var id = command.ExecuteScalar();
                return SafeConvert.ToInt(id);
            }
        }

        /// <summary>
        /// Lưu file đính kèm tới local
        /// </summary>
        private static List<AttachmentObj> SaveAttachmentsToLocal(List<Attachment> attachments)
        {
            try
            {
                if (attachments.Any())
                {
                    List<AttachmentObj> lst = new List<AttachmentObj>();
                    Parallel.ForEach(attachments, attach =>
                    {
                        if (attach is FileAttachment)
                        {
                            FileAttachment attachment = attach as FileAttachment;
                            string fileName = Path.GetFileNameWithoutExtension(attachment.Name) + "_" + DateTime.Now.ToString("yyyy-MM-dd HH mm") + Path.GetExtension(attachment.Name);
                            string path = Path.Combine(localPath, fileName);
                            attachment.Load(path);
                            lock (lst)
                            {
                                lst.Add(new AttachmentObj
                                {
                                    TenFile = attach.Name,
                                    LoaiFile = Path.GetExtension(attachment.Name),
                                    DuongLuong = attachment.Size,
                                    DuongDan = path
                                });
                            }
                        }
                    });
                    return lst;
                }
                return new List<AttachmentObj>();
            }
            catch (Exception ex)
            {
                Common.GhiLog("SaveAttachments", ex);
                return new List<AttachmentObj>();
            }
        }
    }
}
