﻿using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Configuration;
using System.IO;

namespace ConsoleApp21212
{
    class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = ConfigurationManager.AppSettings["appId"],
                TenantId = ConfigurationManager.AppSettings["tenantId"]
            };

            var pca = PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions).Build();

            // The permission scope required for EWS access
            var ewsScopes = new string[] { "https://outlook.office365.com/EWS.AccessAsUser.All" };

            try
            {
                // Make the interactive token request
                //TODO: need to make this requestion not interactive.  Otherwise you will have to sign in everytime you want to use the application
                var authResult = await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService();
                ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);
                ewsClient.TraceEnabled = true;


                //this should create and item and add an email to it
                var UniqueMessageId = "AAMkADJjY2Q0NTU1LTI3OWUtNDBkMC1hMjAxLTY3ZjBlYWQ4Y2U1YwBGAAAAAAA2dt+UZQfOQpdcyGd7ZWbIBwAfK8aAnKcUT4vAD4zbBMc5AAAAAAEMAAAfK8aAnKcUT4vAD4zbBMc5AAAuXBFmAAA=";
                var ChangeKey = "CQAAABYAAADqaMMi7xPwQZ7yRYM19cghAAAuaT0/";
                var ConversationId = "AAQkADcyNzFlNjRmLWMyZTEtNDcyYi1iNDliLWMzNTNlOTI5MDFkYwAQAARxfwPHQYhHqRJDT6lZiew=";
                var PidTagParentEntryIdPR_PARENT_ENTRYIDptagParentEntryId = "AAAAADlbzB7/6zJPrJvvRGj21psBAOpowyLvE/BBnvJFgzX1yCEAAAAAAQwAAA==";
                var PidTagStoreEntryIdPR_STORE_ENTRYIDptagStoreEntryId = "AAAAADihuxAF5RAaobsIACsqVsIAAEVNU01EQi5ETEwAAAAAAAAAABtV+iCqZhHNm8gAqgAvxFoMAAAAUEgwUFIxME1CNTQwNC5uYW1wcmQxMC5wcm9kLm91dGxvb2suY29tAE/mcXLhwitHtJvDU+kpAdz3W+OEU4nqQoRvj0CV4JLE";
                var PidTagEntryIdPR_ENTRYIDptagEntryId = "AAAAADlbzB7/6zJPrJvvRGj21psHAOpowyLvE/BBnvJFgzX1yCEAAAAAAQwAAOpowyLvE/BBnvJFgzX1yCEAAC54F90AAA==";


                //// This bit of code can update fields in the message
                //Item item = Item.Bind(ewsClient, new ItemId(UniqueMessageId));
                //item.Subject = "test";
                //item.Update(ConflictResolutionMode.AutoResolve);

                //// As a best practice, create a property set that limits the properties returned by the Bind method to only those that are required.
                //PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.Subject, EmailMessageSchema.ToRecipients);
                //// This method call results in a GetItem call to EWS.
                //EmailMessage message = EmailMessage.Bind(ewsClient, UniqueMessageId, propSet);
                //// Send the email message.

                //// This method call results in a SendItem call to EWS.
                ///
                //message.Update(ConflictResolutionMode.AutoResolve);
                //message.Send();
                //Console.WriteLine("An email with the subject '" + message.Subject + "' has been sent to '" + message.ToRecipients[0] + "'.");


                
                //this can replace emails if you have the message id
                EmailMessage email = EmailMessage.Bind(ewsClient, UniqueMessageId);

                string emlFileName = @"C:\Source\Demos\output.eml";
                using (FileStream fs = new FileStream(emlFileName, FileMode.Open, FileAccess.Read))
                {
                    byte[] bytes = new byte[fs.Length];
                    int numBytesToRead = (int)fs.Length;
                    int numBytesRead = 0;
                    while (numBytesToRead > 0)
                    {
                        int n = await fs.ReadAsync(bytes, numBytesRead, numBytesToRead);
                        if (n == 0)
                            break;
                        numBytesRead += n;
                        numBytesToRead -= n;
                    }
                    // Set the contents of the .eml file to the MimeContent property.
                    email.MimeContent = new MimeContent("UTF-8", bytes);
                }

                // Indicate that this email is not a draft. Otherwise, the email will appear as a 
                // draft to clients.
                ExtendedPropertyDefinition PR_MESSAGE_FLAGS_msgflag_read = new ExtendedPropertyDefinition(3591, MapiPropertyType.Integer);
                email.SetExtendedProperty(PR_MESSAGE_FLAGS_msgflag_read, 1);
                // This results in a CreateItem call to EWS. The email will be saved in the Inbox folder.
                email.Update(ConflictResolutionMode.AutoResolve);




                //// this will create and send an email
                //EmailMessage email = new EmailMessage(ewsClient);
                //email.ToRecipients.Add("AdeleV@psbii.onmicrosoft.com");
                //email.Subject = "message throught EWS";
                //email.Body = new MessageBody("<h1>This secondary test email with ews, the application is giving me some other properties.</h1><div>  Just reply if you receive it and i will share the information i am able to pull off the email</div>");
                //email.Send();

                // Make an EWS call
                var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));
                foreach (var folder in folders)
                {
                    Console.WriteLine($"Folder: {folder.DisplayName}");
                    
                }
            }
            catch (MsalException ex)
            {
                Console.WriteLine($"Error acquiring access token: {ex}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex}");
            }

            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine("Hit any key to exit...");
                Console.ReadKey();
            }
        }
    }
}