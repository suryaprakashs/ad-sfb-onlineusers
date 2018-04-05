using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using Microsoft.Lync.Model;

namespace SimpleMan.OnlineUsers
{
    class Program
    {
        static ContactSubscription _contactSubscription;

        static void Main(string[] args)
        {
            // really takes long time to finish
            List<string> users = GetADUsers();

            Console.WriteLine("Users in AD: " + users.Count);

            PrintUserAvailability(users);

            Console.ReadKey();

            Console.WriteLine("Unsubscribing contacts...");

            _contactSubscription.Unsubscribe();

            Console.ReadKey();
        }

        private static void PrintUserAvailability(List<string> users)
        {
            try
            {
                var lyncClient = LyncClient.GetClient();

                _contactSubscription = LyncClient.GetClient().ContactManager.CreateSubscription();

                foreach (var usr in users)
                {
                    string user = "sip:" + usr;

                    Contact contact = null;

                    try
                    {
                        contact = lyncClient.ContactManager.GetContactByUri(user);

                        if (contact != null)
                        {
                            contact.ContactInformationChanged += contact_ContactInformationChanged;
                            _contactSubscription.AddContact(contact);
                        }
                    }
                    catch
                    { }
                }

                _contactSubscription.Unsubscribe();
                _contactSubscription.Subscribe(ContactSubscriptionRefreshRate.High, new List<ContactInformationType>() { ContactInformationType.Availability });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        static void contact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            var contact = sender as Contact;
            var currentStatus = contact.GetContactInformation(new List<ContactInformationType>() { ContactInformationType.Availability })[ContactInformationType.Availability];

            if ((int)currentStatus > 0 && (int)currentStatus < 6500)
            {
                Console.WriteLine(contact.Uri + " >> " + "Online");
            }
        }

        private static List<string> GetADUsers()
        {
            List<string> adUsers = new List<string>();

            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, "sensiple.com"))
                {
                    using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                    {
                        foreach (var result in searcher.FindAll())
                        {
                            DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;
                            if (de.Properties["sn"].Value != null && !String.IsNullOrEmpty(de.Properties["sn"].Value.ToString()))
                            {
                                adUsers.Add(de.Properties["userPrincipalName"].Value.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            return adUsers;
        }
    }
}
