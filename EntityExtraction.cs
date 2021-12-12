using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp21212
{
    class EntityExtraction
    {
        public static void ExtractEntities(ExchangeService ewsClient, ItemId UniqueMessageId)
        {
            // Create a property set that limits the properties returned 
            // by the Bind method to only those that are required.
            PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.EntityExtractionResult);
            // Get the item from the server.
            // This method call results in an GetItem call to EWS.
            Item item = Item.Bind(ewsClient, UniqueMessageId, propSet);
            Console.WriteLine("The following entities have been extracted from the message:");
            Console.WriteLine(" ");
            // If address entities are extracted from the message, print the results.
            if (item.EntityExtractionResult != null)
            {
                if (item.EntityExtractionResult.Addresses != null)
                {
                    Console.WriteLine("--------------------Addresses---------------------------");
                    foreach (AddressEntity address in item.EntityExtractionResult.Addresses)
                    {
                        Console.WriteLine("Address: {0}", address.Address);
                    }
                    Console.WriteLine(" ");
                }
                // If contact entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.Contacts != null)
                {
                    Console.WriteLine("--------------------Contacts----------------------------");
                    foreach (ContactEntity contact in item.EntityExtractionResult.Contacts)
                    {
                        Console.WriteLine("Addresses:       {0}", contact.Addresses);
                        Console.WriteLine("Business name:   {0}", contact.BusinessName);
                        Console.WriteLine("Contact string:  {0}", contact.ContactString);
                        Console.WriteLine("Email addresses: {0}", contact.EmailAddresses);
                        Console.WriteLine("Person name:     {0}", contact.PersonName);
                        Console.WriteLine("Phone numbers:   {0}", contact.PhoneNumbers);
                        Console.WriteLine("URLs:            {0}", contact.Urls);
                    }
                    Console.WriteLine(" ");
                }
                // If email address entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.EmailAddresses != null)
                {
                    Console.WriteLine("--------------------Email addresses---------------------");
                    foreach (EmailAddressEntity email in item.EntityExtractionResult.EmailAddresses)
                    {
                        Console.WriteLine("Email addresses: {0}", email.EmailAddress);
                    }
                    Console.WriteLine(" ");
                }
                // If meeting suggestion entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.MeetingSuggestions != null)
                {
                    Console.WriteLine("--------------------Meeting suggestions-----------------");
                    foreach (MeetingSuggestion meetingSuggestion in item.EntityExtractionResult.MeetingSuggestions)
                    {
                        Console.WriteLine("Meeting subject:  {0}", meetingSuggestion.Subject);
                        Console.WriteLine("Meeting string:   {0}", meetingSuggestion.MeetingString);
                        foreach (EmailUserEntity attendee in meetingSuggestion.Attendees)
                        {
                            Console.WriteLine("Attendee name:    {0}", attendee.Name);
                            Console.WriteLine("Attendee user ID: {0}", attendee.UserId);
                        }
                        Console.WriteLine("Start time:       {0}", meetingSuggestion.StartTime);
                        Console.WriteLine("End time:         {0}", meetingSuggestion.EndTime);
                        Console.WriteLine("Location:         {0}", meetingSuggestion.Location);
                    }
                    Console.WriteLine(" ");
                }
                // If phone number entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.PhoneNumbers != null)
                {
                    Console.WriteLine("--------------------Phone numbers-----------------------");
                    foreach (PhoneEntity phone in item.EntityExtractionResult.PhoneNumbers)
                    {
                        Console.WriteLine("Original phone string:  {0}", phone.OriginalPhoneString);
                        Console.WriteLine("Phone string:           {0}", phone.PhoneString);
                        Console.WriteLine("Type:                   {0}", phone.Type);
                    }
                    Console.WriteLine(" ");
                }
                // If task suggestion entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.TaskSuggestions != null)
                {
                    Console.WriteLine("--------------------Task suggestions--------------------");
                    foreach (TaskSuggestion task in item.EntityExtractionResult.TaskSuggestions)
                    {
                        foreach (EmailUserEntity assignee in task.Assignees)
                        {
                            Console.WriteLine("Assignee name:    {0}", assignee.Name);
                            Console.WriteLine("Assignee user ID: {0}", assignee.UserId);
                        }
                        Console.WriteLine("Task string:      {0}", task.TaskString);
                    }
                    Console.WriteLine(" ");
                }
                // If URL entities are extracted from the message, print the results.
                if (item.EntityExtractionResult.Urls != null)
                {
                    Console.WriteLine("--------------------URLs--------------------------------");
                    foreach (UrlEntity url in item.EntityExtractionResult.Urls)
                    {
                        Console.WriteLine("URL: {0}", url.Url);
                    }
                    Console.WriteLine(" ");
                }
            }
            // If no entities are extracted from the message, print the result.
            else if (item.EntityExtractionResult == null)
            {
                Console.WriteLine("No entities extracted");
            }
        }
    }
}
