using Microsoft.Exchange.WebServices.Data;
using System;
using System.Configuration;
using System.Data.SqlClient;
using System.ServiceProcess;
using System.Timers;
using wsEmailExchange.Util;

namespace wsEmailExchange
{
    public partial class MailExchangeService : ServiceBase
    {
        private Timer timerBegin;
        static int timerInterval = Convert.ToInt32(ConfigurationManager.AppSettings["timerInterval"].ToString());
        static string connStr = ConfigurationManager.ConnectionStrings["ConnStr"].ConnectionString.ToString();


        public MailExchangeService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                Common.GhiLog("OnStart", "Bắt đầu chạy service.", "");
                Common.GhiLog("Connecstring", connStr);
                SqlConnection connection = (SqlConnection)null;
                string msgErr = string.Empty;
                if (!Dao.GetConnect(ref connection, connStr, ref msgErr))
                {
                    Common.GhiLog("OnStart", "Không kết nối được CSDL");
                    this.Stop();
                }
                else
                {
                    connection.Close();
                }
                timerBegin = new Timer();
                timerBegin.Interval = timerInterval;
                timerBegin.Elapsed += new ElapsedEventHandler(timerBegin_Elapsed);
                timerBegin.Start();
            }
            catch (Exception ex)
            {
                Common.GhiLog("OnStart", ex, "");
                this.Stop();
            }
        }

        private void timerBegin_Elapsed(object sender, ElapsedEventArgs e)
        {
            ExchangeService exchangeService = MailUtil.InitExchangeObj();
            if (exchangeService == null)
            {
                Common.GhiLog("Init Service loi", "");
                return;
            }
            MailUtil.ReadMailUnreadInbox(exchangeService);
        }

        protected override void OnStop()
        {
            Common.GhiLog("OnStop", "Stop Service", "");
        }
    }
}
