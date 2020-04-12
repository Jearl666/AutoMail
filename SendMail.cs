using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;

namespace AutoMail_v1
{
    class SendMail
    {
        private string smtphost;
        public SendMail(string host)
        {
            smtphost = host;
        }
        ~SendMail() { }

        /// <summary>
        /// 发送excel报表图片和附件
        /// </summary>
        /// <param name="from"></param>
        /// <param name="cc"></param>
        /// <param name="to"></param>
        /// <param name="subject"></param>
        /// <param name="imageComment"></param>
        /// <param name="imagePATH">仅指定图片名称</param>
        /// <param name="attPath"></param>
        /// <param name="Autopath"></param>
        public void mailinfo(string from, string cc, string to, string subject, string imageComment, string imagePATH, string attPath,string Autopath)
        {
            try
            {
                SmtpClient smc = new SmtpClient();
                smc.Host = smtphost;
                MailMessage mm = new MailMessage();
                mm.From = new MailAddress(from);

                if (cc != "")
                {
                    List<string> ccMaillist = cc.Split(';').ToList();
                    for (int j = 0; j < ccMaillist.Count; j++)
                    {
                        mm.CC.Add(new MailAddress(ccMaillist[j]));
                    }
                }

                List<string> toMaillist = to.Split(';').ToList();
                for (int i = 0; i < toMaillist.Count; i++)
                {
                    mm.To.Add(new MailAddress(toMaillist[i]));
                }
                mm.Subject = subject;
                //mm.Priority = MailPriority.High;
                if (imagePATH != "")
                {
                    mm.IsBodyHtml = true;
                    mm.BodyEncoding = System.Text.Encoding.UTF8;
                    //发送Html格式图片
                    string HtmlBodyContent = null;// "<img src=\"cid:imgurl\"><img src=\"cid:imgurl2\">";

                    List<string> imageCommentList = imageComment.Split(';').ToList();
                    List<string> imagePATHList = imagePATH.Split(';').ToList();

                    for (int i = 0; i < imagePATHList.Count; i++)
                    {
                        if (i < imageCommentList.Count && imageCommentList[i]!="")
                        {
                            HtmlBodyContent = HtmlBodyContent + "<h3>" + imageCommentList[i] + "</h3>" +
                                "<img src=\"cid:imgurl" + i + "\">";
                        }
                        else
                        {
                            HtmlBodyContent = HtmlBodyContent + "<img src=\"cid:imgurl" + i + "\">";
                        }
                    }

                    AlternateView htmlBody = AlternateView.CreateAlternateViewFromString(HtmlBodyContent, null, "text/html");
                    for (int i = 0; i < imagePATHList.Count; i++)
                    {
                        LinkedResource lrImage = new LinkedResource(Autopath+imagePATHList[i], "image/gif");
                        lrImage.ContentId = "imgurl" + i;
                        htmlBody.LinkedResources.Add(lrImage);
                    }
                    mm.AlternateViews.Add(htmlBody);
                }
                if (attPath !="")
                {
                    //发送附件
                    Attachment attr = new Attachment(attPath, MediaTypeNames.Text.Plain);
                    mm.Attachments.Add(attr);
                }

                smc.Send(mm);
                mm.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            
            }         
        }

        /// <summary>
        /// 不运行宏，只发送图片和attach
        /// </summary>
        /// <param name="from"></param>
        /// <param name="cc"></param>
        /// <param name="to"></param>
        /// <param name="subject"></param>
        /// <param name="imageComment"></param>
        /// <param name="imagePATH">指定图片完整路径</param>
        /// <param name="attPath"></param>
        public void mailinfo(string from, string cc, string to, string subject, string imageComment, string imagePATH, string attPath)
        {
            try
            {
                SmtpClient smc = new SmtpClient();
                smc.Host = smtphost;
                MailMessage mm = new MailMessage();
                mm.From = new MailAddress(from);

                if (cc != "")
                {
                    List<string> ccMaillist = cc.Split(';').ToList();
                    for (int j = 0; j < ccMaillist.Count; j++)
                    {
                        mm.CC.Add(new MailAddress(ccMaillist[j]));
                    }
                }

                List<string> toMaillist = to.Split(';').ToList();
                for (int i = 0; i < toMaillist.Count; i++)
                {
                    mm.To.Add(new MailAddress(toMaillist[i]));
                }

                mm.Subject = subject;
                //mm.Priority = MailPriority.High;
                if (imagePATH != "")
                {
                    mm.IsBodyHtml = true;
                    mm.BodyEncoding = System.Text.Encoding.UTF8;
                    //发送Html格式图片
                    string HtmlBodyContent = null;// "<img src=\"cid:imgurl\"><img src=\"cid:imgurl2\">";

                    List<string> imageCommentList = imageComment.Split(';').ToList();
                    List<string> imagePATHList = imagePATH.Split(';').ToList();

                    for (int i = 0; i < imagePATHList.Count; i++)
                    {
                        if (i < imageCommentList.Count && imageCommentList[i] != "")
                        {
                            HtmlBodyContent = HtmlBodyContent + "<h3>" + imageCommentList[i] + "</h3>" +
                                "<img src=\"cid:imgurl" + i + "\">";
                        }
                        else
                        {
                            HtmlBodyContent = HtmlBodyContent + "<img src=\"cid:imgurl" + i + "\">";
                        }
                    }

                    AlternateView htmlBody = AlternateView.CreateAlternateViewFromString(HtmlBodyContent, null, "text/html");
                    for (int i = 0; i < imagePATHList.Count; i++)
                    {
                        LinkedResource lrImage = new LinkedResource(imagePATHList[i], "image/gif");
                        lrImage.ContentId = "imgurl" + i;
                        htmlBody.LinkedResources.Add(lrImage);
                    }
                    mm.AlternateViews.Add(htmlBody);
                }

                if (attPath != "")
                {
                    //发送附件
                    Attachment attr = new Attachment(attPath, MediaTypeNames.Text.Plain);
                    mm.Attachments.Add(attr);
                }

                smc.Send(mm);
                mm.Dispose();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void watchdog(string subject,string errorinfo,string errorfilepath)
        {
            SmtpClient smc = new SmtpClient();
            smc.Host = smtphost;
            MailMessage mm = new MailMessage();
            mm.From = new MailAddress("xxx@xxx.com");
            mm.CC.Add(new MailAddress("xxx@xxx.com"));
            mm.To.Add(new MailAddress("xxx@xxx.com"));
            mm.Subject = "[AutoMail Send Fail-->]"+subject;
            mm.IsBodyHtml = true;
            mm.BodyEncoding = System.Text.Encoding.UTF8;
            string HtmlBodyContent = "<h4>" + errorfilepath + "</h4>" + "<h4>" + errorinfo + "</h4>";
            AlternateView htmlBody = AlternateView.CreateAlternateViewFromString(HtmlBodyContent, null, "text/html");
            mm.AlternateViews.Add(htmlBody);
            smc.Send(mm);
            mm.Dispose();

        }

    }
}
