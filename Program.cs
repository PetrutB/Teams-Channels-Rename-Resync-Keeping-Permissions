using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;

namespace GraphApp1
{
    class Program
    {
        public static async Task GraphOp (GraphServiceClient graphClient)
        {

            // Reading the Teams List
            List<string> Teams = new List<string>();
            foreach (string line in System.IO.File.ReadLines(@"***\Teams.txt"))
            {
                if (line.Length > 10 && line.Contains("Guid") == false && line.Contains("---") == false)
                    Teams.Add(line);
            }

            foreach (String Team in Teams)
            {
                Console.WriteLine(Team);

                // Get Channels for a Team
                var channels = await graphClient.Teams[Team].Channels
                    .Request()
                    .GetAsync();

                // Going Through all the Channels of a Team
                foreach (Microsoft.Graph.Channel C in channels)
                {
                    // Getting the File Folder for each Channel
                    var FileFolder = await graphClient.Teams[Team].Channels[C.Id].FilesFolder
                        .Request()
                        .GetAsync();

                    // Checking if the Channel Name is the same as the File Folder Name
                    if (C.DisplayName != FileFolder.Name)
                    {
                        Console.WriteLine("Wrong Name " + C.DisplayName + " " + FileFolder.Name);

                        // Get the Teams Drive ID
                        var TeamsDrive = await graphClient.Groups[Team].Drive
                            .Request()
                            .GetAsync();

                        Console.WriteLine("Teams Drive ID used here:" + FileFolder.Id);

                        // Trying to get the List of Childrens for My Channel Folder
                        try
                        {
                            // Get Children Elements of the Channels SPO Folder
                            var DriveItemsUnderMainFolder = await graphClient.Drives[TeamsDrive.Id].Items[FileFolder.Id].Children
                                .Request()
                                .GetAsync();

                            // Checking if there are any Files under the Folder
                            if (DriveItemsUnderMainFolder.Count != 0)
                            {

                                Console.WriteLine("First Children Id: " + DriveItemsUnderMainFolder[0].Id);

                                // Going through the Files under the Folder
                                foreach (DriveItem i in DriveItemsUnderMainFolder)
                                {
                                    Console.WriteLine(i.Name);

                                    // Getting the Permissions of each Files under the Folder
                                    var DriveItemPermissions = await graphClient.Drives[TeamsDrive.Id].Items[i.Id].Permissions
                                        .Request()
                                        .GetAsync();

                                    // Should be Renamed Here the Channel

                                    // Should get the new Files here


                                    // Going through the Permissions of each Files under the Folder
                                    foreach (Permission j in DriveItemPermissions)
                                    {
                                        // Checking for the Permissions to be Coming from an URL
                                        if (j.Link != null && j.InheritedFrom == null)
                                        {
                                            // Displaying the Shareable Link
                                            Console.WriteLine("This is the Shareable Link: " + j.Link.WebUrl + " " + j.InheritedFrom);

                                            // Type and Scope of the Sharing Link already existing
                                            var type = j.Link.Type;
                                            var scope = j.Link.Scope;

                                            // List were I'll try to get each User with which the URL was shared
                                            List<User> users = new List<User>();

                                            if (j.GrantedToIdentitiesV2 != null)
                                            {
                                                // Users of the Sharing Link already existing
                                                foreach (IdentitySet k in j.GrantedToIdentitiesV2)
                                                {
                                                    if (k.User.DisplayName != null)
                                                    {
                                                        Console.WriteLine("Found an user: " + k.User.DisplayName);

                                                        var user = await graphClient.Users[k.User.Id]
                                                            .Request()
                                                            .GetAsync();

                                                        Console.WriteLine("With email address: " + user.Mail);
                                                        users.Add(user);
                                                    }
                                                }
                                            }

                                            // Create the new Sharing Link
                                            await graphClient.Drives[TeamsDrive.Id].Items[DriveItemsUnderMainFolder[0].Id]
                                                .CreateLink(type, scope, null, null, null, null)
                                                .Request()
                                                .PostAsync();

                                            // Create a List of Users for the new Sharing Link
                                            if (users != null)
                                            {
                                                List<DriveRecipient> recipients = new List<DriveRecipient>();
                                                foreach (User k in users)
                                                {
                                                    recipients.Add(new DriveRecipient
                                                    {
                                                        Email = k.Mail
                                                    });
                                                }

                                                // Getting the Permissions for our new URL
                                                var permissionsNewURL = await graphClient.Drives[TeamsDrive.Id].Items[DriveItemsUnderMainFolder[0].Id].Permissions
                                                    .Request()
                                                    .GetAsync();

                                                // Encoding the URL in Order to get the ShareID
                                                string encodedUrl = null;
                                                foreach (Permission ij in permissionsNewURL)
                                                {
                                                    if (ij.Link != null && ij.InheritedFrom == null)
                                                    {
                                                        string sharingUrl = ij.Link.WebUrl;
                                                        string base64Value = System.Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(sharingUrl));
                                                        encodedUrl = "u!" + base64Value.TrimEnd('=').Replace('/', '_').Replace('+', '-');
                                                    }
                                                }

                                                // Creating the URL Roles
                                                var roles = new List<String>();

                                                if (type == "edit")
                                                    roles.Add("write");
                                                else
                                                    roles.Add("read");

                                                // Updating the URL Permissions
                                                await graphClient.Shares[encodedUrl].Permission
                                                    .Grant(roles, recipients)
                                                    .Request()
                                                    .PostAsync();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("Children Command Failed - Likley no File in the Folder:" + FileFolder.Name);
                        }
                    }
                    else
                        Console.WriteLine("Correct Name " + C.DisplayName + " " + FileFolder.Name);
                }
            }
        }

        static void Main(string[] args)
        {

            // The client credentials flow requires that you request the
            // /.default scope, and preconfigure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = "****";

            // Values from app registration
            var clientId = "****";
            var clientSecret = "****";

            // using Azure.Identity;
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            Console.WriteLine("Start");
            GraphOp(graphClient).GetAwaiter().GetResult();
            Console.WriteLine("End");
        }
    }
}
