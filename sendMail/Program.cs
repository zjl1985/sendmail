using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
namespace sendMail
{
    class Program
    {
        SmtpClient SmtpClient = null;   //设置SMTP协议
        MailAddress MailAddress_from = null; //设置发信人地址  当然还需要密码
        MailAddress MailAddress_to = null;  //设置收信人地址  不需要密码
        MailMessage MailMessage_Mai = null;
        static void Main(string[] args)
        {

            sendMail.Program pro = new Program();
            pro.MailMessage_Mai = new MailMessage();
            pro.setSmtpClient("smtp.163.com", 25);
//todo 密码和邮箱
            pro.setAddressform("zjl198544@163.com", "");
            pro.sendtheMail();
            Console.ReadLine();

        }

        #region 设置Smtp服务器信息
        /// <summary>
        /// 设置Ｓmtp服务器信息
        /// </summary>
        /// <param name="ServerName">SMTP服务名</param>
        /// <param name="Port">端口号</param>
        private void setSmtpClient(string ServerHost, int Port)
        {
            SmtpClient = new SmtpClient();
            SmtpClient.Host = ServerHost;//指定SMTP服务名  例如QQ邮箱为 smtp.qq.com 新浪cn邮箱为 smtp.sina.cn等
            SmtpClient.Port = Port; //指定端口号
            SmtpClient.Timeout = 0;  //超时时间

        }
        #endregion

        #region 验证发件人信息
        /// <summary>
        /// 验证发件人信息
        /// </summary>
        /// <param name="MailAddress">发件邮箱地址</param>
        /// <param name="MailPwd">邮箱密码</param>
        private void setAddressform(string MailAddress, string MailPwd)
        {
            //创建服务器认证
            NetworkCredential NetworkCredential_my = new NetworkCredential(MailAddress, MailPwd);
            //实例化发件人地址
            MailAddress_from = new System.Net.Mail.MailAddress(MailAddress, "小张");
            //指定发件人信息  包括邮箱地址和邮箱密码
            SmtpClient.Credentials = new System.Net.NetworkCredential(MailAddress_from.Address, MailPwd);
        }
        #endregion

        private void sendtheMail()
        {
            //检测附件大小 发件必需小于10M 否则返回  不会执行以下代码
            //if (txt_Path.Text != "")
            //{
            //    if (!Attachment_MaiInit(txt_Path.Text.Trim()))
            //    {
            //        return;
            //    }
            //}

            //清空历史发送信息 以防发送时收件人收到的错误信息(收件人列表会不断重复)
            sendMail.Program pro = new Program();
            string strPath = ConfigurationManager.AppSettings["filepath"];
            var lines = File.ReadLines(strPath);

            MailMessage_Mai.To.Clear();
            //添加收件人邮箱地址
            List<string> sendList = new List<string>();
            foreach (var line in lines)
            {
                sendList.Add(line);
            }
            // sendList.Add("zjl198544@163.com");
            foreach (string sendadd in sendList)
            {
                MailAddress_to = new MailAddress(sendadd);
                MailMessage_Mai.To.Add(MailAddress_to);
            }
            Console.WriteLine("收件人：" + MailMessage_Mai.To.Count.ToString() + "  个");

            //发件人邮箱
            MailMessage_Mai.From = MailAddress_from;
            //邮件主题
            string strtitle = "数据中断列表" + DateTime.Now.ToShortDateString();
            MailMessage_Mai.Subject = strtitle;
            MailMessage_Mai.SubjectEncoding = System.Text.Encoding.UTF8;
            //邮件正文
            MailMessage_Mai.Body = pro.Htmlcontent();
            MailMessage_Mai.BodyEncoding = System.Text.Encoding.UTF8;
            MailMessage_Mai.IsBodyHtml = true;
            //清空历史附件  以防附件重复发送
            MailMessage_Mai.Attachments.Clear();
            //添加附件
            //    MailMessage_Mai.Attachments.Add(new Attachment(txt_Path.Text.Trim(), MediaTypeNames.Application.Octet));
            //注册邮件发送完毕后的处理事件
            //SmtpClient.Send(MailMessage_Mai);
            SmtpClient.SendCompleted += new SendCompletedEventHandler(SendCompletedCallback);
            //开始发送邮件
            SmtpClient.SendAsync(MailMessage_Mai, "000000000");

        }

        #region 封装邮件内容
        private string Htmlcontent()
        {
            StringBuilder mailHtml = new StringBuilder();
            SqlConnection conn = new SqlConnection(ConfigurationManager.AppSettings["connString"]);
            int intv = int.Parse(ConfigurationManager.AppSettings["intv"]);
            string str_Time = DateTime.Now.AddMinutes(intv*-1).ToString();
            SqlCommand cmdzd = new SqlCommand("select RoadName,LastTime,XJZone from SLastTime where LastTime<'" + str_Time + "' and RoadName not like '%●%' and RoadName not like '%Δ%' and islook=1 order by xjzone,LastTime DESC", conn);
            SqlCommand cmdnew = new SqlCommand("select RoadName,LastTime,XJZone from SLastTime where LastTime<'" + str_Time + "' and RoadName like '%Δ%' and islook=1 order by xjzone,LastTime DESC", conn);
            SqlDataAdapter da = new SqlDataAdapter(cmdzd);
            DataTable dt1 = new DataTable();
            DataTable dtnew = new DataTable();
            da.Fill(dt1);
            da = new SqlDataAdapter(cmdnew);
            da.Fill(dtnew);
            mailHtml.Append("<div align=\"center\">");
            mailHtml.Append("<p style=\"font-size: 10pt\">以下内容为系统自动发送，请勿直接回复</p><table cellspacing=\"1\" cellpadding=\"3\" border=\"0\" bgcolor=\"000000\" style=\"font-size: 10pt;line-height: 15px;\">");
            
            mailHtml.Append("<tr><td bgcolor=\"#CC9999\" colspan=3 align=\"center\">一期数据中断:" + dt1.Rows.Count + "个</td></tr> ");
            for (int row = 0; row < dt1.Rows.Count; row++)
            {
                mailHtml.Append("<tr>");
                for (int col = 0; col < dt1.Columns.Count; col++)
                {
                    mailHtml.Append("<td bgcolor=\"#FFFFFF\">&nbsp;&nbsp;&nbsp;");
                    mailHtml.Append(dt1.Rows[row][col].ToString());
                    mailHtml.Append("&nbsp;&nbsp;&nbsp;</td>");
                }
                mailHtml.Append("</tr>");
            }

            mailHtml.Append("<tr><td bgcolor=\"#CCFF00\" colspan=3 align=\"center\">新安装数据中断:" + dtnew.Rows.Count + "个</td></tr> ");
            for (int row = 0; row < dtnew.Rows.Count; row++)
            {
                mailHtml.Append("<tr>");
                for (int col = 0; col < dtnew.Columns.Count; col++)
                {
                    mailHtml.Append("<td bgcolor=\"#FFFFCC\">&nbsp;&nbsp;&nbsp;");
                    mailHtml.Append(dtnew.Rows[row][col].ToString());
                    mailHtml.Append("&nbsp;&nbsp;&nbsp;</td>");
                }
                mailHtml.Append("</tr>");
            }


            mailHtml.Append("</table>");
            mailHtml.Append("</div>");
            return mailHtml.ToString();
        }


        #endregion

        #region 发送邮件后所处理的函数
        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            try
            {
                if (e.Cancelled)
                {
                    Console.WriteLine("发送已取消！");
                    Environment.Exit(0);
                }
                if (e.Error != null)
                {

                    Console.WriteLine("邮件发送失败！" + "\n" + "技术信息:\n" + e.ToString());
                    Environment.Exit(0);
                }
                else
                {
                    Console.WriteLine("邮件成功发出!");
                    Environment.Exit(0);
                }
            }
            catch (Exception Ex)
            {
                Console.WriteLine("邮件发送失败！" + "\n" + "技术信息:\n" + Ex.Message);
            }

        }
        #endregion
    }

}
