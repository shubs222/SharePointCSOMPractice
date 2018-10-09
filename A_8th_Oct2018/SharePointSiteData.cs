using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using Microsoft.SharePoint.Client;
namespace A_8th_Oct2018
{
    class SharePointSiteData
    {
        ClientContext clientcntx;
        Web webpage;
        public void GetData(string Url, string UserName, SecureString passwrd)
        {
            using (clientcntx = new ClientContext(Url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(UserName, passwrd);
                webpage = clientcntx.Web;
                clientcntx.Load(webpage);
                try
                {
                    clientcntx.ExecuteQuery();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error: " + e);
                    throw e;
                }
                Console.WriteLine("Share Point Site \n Title: " + webpage.Title + "; URL: " + webpage.Url + "; Description: " + webpage.Description);
                Console.ReadKey();

                Console.WriteLine("Do you want to change the name of the site 1. Yes \t 2. press any key to exit");
                string answer = Console.ReadLine();
                if (answer.ToUpper() == "YES")
                {
                    string Title;
                    Console.WriteLine("Enter the title");
                    Title = Console.ReadLine();
                    webpage.Title = Title;
                    webpage.Update();
                    try
                    {
                        clientcntx.ExecuteQuery();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error " + e);
                        throw e;
                    }
                    Console.WriteLine("New web title is: " + webpage.Title);
                    Console.ReadKey();
                }
                else
                {
                    
                }
            }

        }


        public void CreatenewSubsite(string url, string Username, SecureString password)
        {
            using (clientcntx = new ClientContext(url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                WebCreationInformation crete = new WebCreationInformation();
                crete.Url = "newsite1";
                Console.WriteLine("Enter the title for share point site");
                string title = Console.ReadLine();
                crete.Title = title;
                webpage = clientcntx.Web.Webs.Add(crete);
                clientcntx.Load(webpage, w => w.Title);
                try
                {
                    clientcntx.ExecuteQuery();

                }
                catch (Exception e)
                {
                    Console.WriteLine("Error : " + e);
                    throw e;
                }
                Console.WriteLine("New site" + crete.Title);
                Console.ReadKey();
            }

        }


        public void GetsiteList(string url, string Username, SecureString password)
        {
            using (clientcntx = new ClientContext(url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                webpage = clientcntx.Web;
                clientcntx.Load(webpage.Lists, lists => lists.Include(list => list.Title, list => list.Id));

                try
                {
                    clientcntx.ExecuteQuery();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error: " + e);
                    throw e;
                }
                foreach (List list in webpage.Lists)
                {
                    Console.WriteLine("List name: {0} ", list.Title);

                }
                Console.ReadKey();
            }
        }


        public void CreateSharePointList(string url, string Username, SecureString password)
        {

            Console.WriteLine("Do you want to create new List press Yes to continue or any key to exit");
            string answer = Console.ReadLine();
            if (answer.ToUpper() == "YES")
            {
                using (clientcntx = new ClientContext(url))
                {
                    clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                    ListCreationInformation listCreation = new ListCreationInformation();
                    Console.WriteLine("Enter List Title");
                    listCreation.Title = Console.ReadLine();
                    listCreation.Description = Console.ReadLine();
                    listCreation.TemplateType = Convert.ToInt32(ListTemplateType.GenericList);
                    //List list = clientcntx.Web.Lists.Add(listCreation);
                    clientcntx.Load(clientcntx.Web.Lists.Add(listCreation));
                    try
                    {
                        clientcntx.ExecuteQuery();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error: " + e);
                    }
                    Console.WriteLine("Name of the list: " + clientcntx.Web.Lists.Add(listCreation).Title);
                }
            }
            else
            {

            }
        }


        public void DeleteSpList(string url, string Username, SecureString password)
        {
            Console.WriteLine("Do you want to create delete List press Yes to continue or any key to exit");
            string answer = Console.ReadLine();
            if (answer.ToUpper() == "YES")
            {
                using (clientcntx = new ClientContext(url))
                {
                    clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);

                    Console.WriteLine("Enter the name of the list to delete ");
                    List list = clientcntx.Web.Lists.GetByTitle(Console.ReadLine());
                    list.DeleteObject();
                    clientcntx.ExecuteQuery();
                }
            }
            else
            {

            }
        }


        public void UpdateList(string url, string Username, SecureString password)
        {

        }


        public void CreateFolder(string url, string Username, SecureString password)
        {

            using (clientcntx = new ClientContext(url))
            {
                Console.WriteLine("Enter name of the List: ");
                List list = clientcntx.Web.Lists.GetByTitle(Console.ReadLine());
                ListItemCreationInformation listItem = new ListItemCreationInformation();
                listItem.UnderlyingObjectType = FileSystemObjectType.Folder;
                ListItem listnewitm = list.AddItem(listItem);
                Console.WriteLine("Enter Folder name: ");
                string foldername = Console.ReadLine();
                foldername.Trim();
                listnewitm["Title"] = foldername;
                listnewitm.Update();
                try
                {
                    clientcntx.ExecuteQuery();
                    Console.WriteLine("{0} folder is created ", foldername);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error: " + e);
                    throw;
                }
            }
        }


        public void DeleteFolder(string url, string Username, SecureString password)
        {

        }


        public void CreatenewFolder(string url, string Username, SecureString password)
        {
            Console.WriteLine("Do you want to create new folder press Yes to continue or any key to exit");
            string answer = Console.ReadLine();
            if (answer.ToUpper() == "YES")
            {
                using (clientcntx = new ClientContext(url))
                {
                    clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                    Console.WriteLine("Enter the name of the list: ");
                    List list = clientcntx.Web.Lists.GetByTitle(Console.ReadLine());
                    Folder folder = list.RootFolder;
                    clientcntx.Load(folder);
                    clientcntx.ExecuteQuery();
                    Console.WriteLine("Enter name of the folder: ");
                    folder = folder.Folders.Add("Newfolder");
                    clientcntx.ExecuteQuery();
                    Console.WriteLine("folder created");
                    Console.Read();
                }
            }
            else
            {

            }

        }


        public void UploadFiles(string url, string Username, SecureString password)
        {
            Console.WriteLine("Do you want to upload file in List press Yes to continue or any key to exit");
            string answer = Console.ReadLine();
            if (answer.ToUpper() == "YES")
            {
                using (clientcntx = new ClientContext(url))
                {
                    clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                    Console.WriteLine("Enter the name of the list: ");
                    List list = clientcntx.Web.Lists.GetByTitle(Console.ReadLine());
                    Folder folder =
                       clientcntx.Web.Folders.GetByUrl("https://acuvatehyd.sharepoint.com/:f:/t/shubhamtrial/Em-NREM2bJFEmaApFiErwC0BBOEE-HvxHe6r1Kmc0J5aoA?e=7ILmWx");// list.GetSpecialFolderUrl();
                    clientcntx.ExecuteQuery();



                }

                //private static void UploadFile(ClientContext context, string listTitle, string fileName)
                //{
                //    using (var fs = new FileStream(fileName, FileMode.Open))
                //    {
                //        var fi = new FileInfo(fileName);
                //        var list = context.Web.Lists.GetByTitle(listTitle);
                //        context.Load(list.RootFolder);
                //        context.ExecuteQuery();
                //        var fileUrl = String.Format("{0}/{1}", list.RootFolder.ServerRelativeUrl, fi.Name);

                //        Microsoft.SharePoint.Client.File.SaveBinaryDirect(context, fileUrl, fs, true);
                //    }
                //}

            }
        }


        public void AddField(string url, string Username, SecureString password)
        {
            Console.WriteLine("Do you want to Add field in List press Yes to continue or any key to exit");
            string answer = Console.ReadLine();
            if (answer.ToUpper() == "YES")
            {
                using (clientcntx = new ClientContext(url))
                {
                    clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                    Console.WriteLine("Enter the name of the list: ");
                    List list = clientcntx.Web.Lists.GetByTitle(Console.ReadLine());
                    Field f = list.Fields.AddFieldAsXml("<Field DisplayName='Nationality' Type='Text'/>", true, AddFieldOptions.DefaultValue);
                    f.Update();
                    clientcntx.ExecuteQuery();
                }
            }
        }


        public void DeleteField(string url, string Username, SecureString password)
        {
            Console.WriteLine("Do you want to Delete field in List press Yes to continue or any key to exit");
            string answer = Console.ReadLine();
            if (answer.ToUpper() == "YES")
            {
                using (clientcntx = new ClientContext(url))
                {
                    clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                    Console.WriteLine("Enter the name of the list: ");
                    List list = clientcntx.Web.Lists.GetByTitle(Console.ReadLine());
                    Console.WriteLine("Enter name of the field: ");
                    Field f = list.Fields.GetByTitle(Console.ReadLine());
                    f.DeleteObject();
                    clientcntx.ExecuteQuery();
                }
            }
        }


        public void UploadFile(string url, string Username, SecureString password)
        {
            using (clientcntx = new ClientContext(url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                List list = clientcntx.Web.Lists.GetByTitle("MyDocuments");
                FileCreationInformation fcinfo = new FileCreationInformation();
                fcinfo.Url = "MyDocuments/NewFiles/Products1.txt";
                fcinfo.Content = System.IO.File.ReadAllBytes(@"D:\My Tasks\SharePointPractice\A_8th_Oct2018\Products1.txt");
                fcinfo.Overwrite = true;
                File fileToUpload = list.RootFolder.Files.Add(fcinfo);
                clientcntx.Load(list);
                clientcntx.ExecuteQuery();
                Console.WriteLine("Name is : " + fcinfo.Content);
            }
        }


        public void AddListItem(string url, string Username, SecureString password)
        {
            Console.WriteLine("Do you want to Add item in List press Yes to continue or any key to exit");
            string answer = Console.ReadLine();
            if (answer.ToUpper() == "YES")
            {
                using (clientcntx = new ClientContext(url))
                {
                    clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                    Console.WriteLine("Enter the name of the list: ");
                    List list = clientcntx.Web.Lists.GetByTitle(Console.ReadLine());
                    ListItemCreationInformation listCreation = new ListItemCreationInformation();


                    //FieldLookupValue x = new FieldLookupValue();
                    //x.LookupId = 5;
                    ////x.LookupValue. = "Facility Manager";
                    try
                    {
                        // FieldCollection fcc = list.Fields;

                        //clientcntx.Load(fcc);
                        ListItem ls = list.AddItem(listCreation);
                        ls["Title"] = 5;
                        ls["Departments"] = "fD";
                        //    ls.Update();
                        clientcntx.ExecuteQuery();
                        //Console.WriteLine("cnt  :"+fcc.Count);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error: " + e);
                    }


                }
            }
        }

        public void GetUsers(string url, string Username, SecureString password)
        {
            using (clientcntx = new ClientContext(url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                UserCollection users = clientcntx.Web.SiteUsers;
                clientcntx.Load(users);
                clientcntx.ExecuteQuery();
                foreach (User u in users)
                {
                    Console.WriteLine("Users : "+u.Email);
                }
            }
            
        }


        public void AddUsers(string url, string Username, SecureString password)
        {
            using (clientcntx =new ClientContext(url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                Web web = clientcntx.Web;
                User user = web.EnsureUser("venu.kalam@acuvate.com");
                Group group = web.SiteGroups.GetByName("new group");
                group.Users.AddUser(user);

            }
        }


        public void DeleteUser(string url, string Username, SecureString password)
        {
            using (clientcntx = new ClientContext(url))
            {
                clientcntx.Credentials = new SharePointOnlineCredentials(Username, password);
                UserCollection usercolln = clientcntx.Web.SiteGroups.GetByName("InsertUsers").Users;
                clientcntx.Load(usercolln);
                clientcntx.ExecuteQuery();
                foreach (User usr in usercolln)
                {
                    if (usr.Email == "dharanendra.sheetal@acuvate.com")
                    {
                        try
                        {
                            User u1 = usercolln.GetByEmail(usr.Email);
                            usercolln.Remove(u1);
                            clientcntx.ExecuteQuery();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exe  :" + e);
                        }
                    }
                }


            }
        }
    }
}
