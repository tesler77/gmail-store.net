using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using System.Threading.Tasks;

namespace GmailQuickstart
{
    class gmail
    {
        public GmailService GmailService { get; set; }
        public gmail(GmailService gmailService)
        {
            this.GmailService = gmailService;
        }

        public ListMessagesResponse getUnreadedMwssages()
        {
            var messagesListRequest = this.GmailService.Users.Messages.List("me");
            messagesListRequest.LabelIds = "UNREAD";
            var messagesList = messagesListRequest.Execute();
            return messagesList;
        }


        public Message getMessageData(string id)
        {
            var messageRequest = this.GmailService.Users.Messages.Get("me", id);
            var message = messageRequest.Execute();
            return message;
        }

        public byte[] getAttachment(string attachmentId, string messageId)
        {
            var attchmentRequest = this.GmailService.Users.Messages.Attachments.Get("me", messageId, attachmentId);
            var attachment = attchmentRequest.Execute();
            string attachmentData = attachment.Data.Replace("-", "+").Replace("_", "/");
            byte[] b = Convert.FromBase64String(attachmentData);
            return b;
        }
    }
}
