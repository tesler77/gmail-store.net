using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace GmailQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/gmail-dotnet-quickstart.json
        static string[] Scopes = { GmailService.Scope.GmailModify };
        static string ApplicationName = "Gmail API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Gmail API service.
            var service = new GmailService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            gmail gmail = new gmail(service);
            ListMessagesResponse MessagesList = gmail.getUnreadedMwssages();

            foreach (var message in MessagesList.Messages)
            {
                var messageData = gmail.getMessageData(message.Id);
                foreach (var data in messageData.Payload.Headers)
                {
                    if (data.Name == "From")
                    {
                        string sender = data.Value.Substring(data.Value.IndexOf("<")).Replace("<", "").Replace(">", "");
                        user user = new user(sender);
                        int checkIndicator = user.check();
                        if (checkIndicator != 0)
                            break;
                        string attachmentId = messageData.Payload.Parts[1].Body.AttachmentId;
                        string fileName = messageData.Payload.Parts[1].Filename;
                        if (attachmentId != null && messageData.Payload.Parts[1].MimeType == "application/vnd.ms-excel")
                        {
                            fileName = fileName.Substring(0, fileName.IndexOf("."));
                            byte[] b = gmail.getAttachment(attachmentId, message.Id);
                            string root = @"D:\gmail-store.net\gmail store\public\";
                            if(!Directory.Exists(root + fileName+ @"\"))                            
                                Directory.CreateDirectory(root + fileName + @"\");
                            File.WriteAllBytes(@"D:\gmail-store.net\gmail store\public\" + fileName + @"\" + sender + ".xls", b);
                        }

                    }
                }
            }
        }
    }
}