using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using Microsoft.SharePoint.Client;

namespace A_8th_Oct2018
{
   
    class Program
    {
        static void Main(string[] args)
        {
            string username;
            //Console.WriteLine("Enter your site name");
            string sitename = "shubhamtrial";//Console.ReadLine();
            SharePointSiteData spdata = new SharePointSiteData();
            //Console.WriteLine("Enter USerName");
            username= /*Console.ReadLine()+*/"arvind.torvi@acuvate.com";
            Console.WriteLine("Enter your password.");
            SecureString password = GetPassword();
            //string site = "shubhamtrial";
            string Url = "https://acuvatehyd.sharepoint.com/teams/"+sitename;
            spdata.GetData(Url,username,password);

            
            //Console.WriteLine("List infrormation: ");
            //spdata.GetsiteList(Url, username, password);
            //spdata.CreateSharePointList(Url, username, password);
            //spdata.DeleteSpList(Url, username, password);
            //spdata.CreatenewFolder(Url, username, password);
            //spdata.AddField(Url, username, password);
            //spdata.DeleteField(Url, username, password);
            //spdata.Getpath(Url, username, password);
            //spdata.AddListItem(Url, username, password);
            spdata.GetUsers(Url, username, password);
            Console.ReadKey();


        }
        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;
            //Get the user's password as a SecureString  
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
    }
}
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Net;
//using System.Security;
//using System.Text;
//using System.Threading.Tasks;
//using Microsoft.SharePoint.Client;

//namespace ConsoleApp2
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            string userName = "khalil@c986.onmicrosoft.com";
//            Console.WriteLine("Enter your password.");
//            SecureString password = GetPassword();
//            // ClienContext - Get the context for the SharePoint Online Site  
//            // SharePoint site URL - https://c986.sharepoint.com  
//            using (var clientContext = new ClientContext("https://c986.sharepoint.com"))
//            {
//                // SharePoint Online Credentials  
//                clientContext.Credentials = new SharePointOnlineCredentials(userName, password);
//                // Get the SharePoint web  
//                Web web = clientContext.Web;
//                // Load the Web properties  
//                clientContext.Load(web);
//                // Execute the query to the server.  
//                clientContext.ExecuteQuery();
//                // Web properties - Display the Title and URL for the web  
//                Console.WriteLine("Title: " + web.Title + "; URL: " + web.Url);
//                Console.ReadLine();
//            }
//        }
//        private static SecureString GetPassword()
//        {
//            ConsoleKeyInfo info;
//            //Get the user's password as a SecureString  
//            SecureString securePassword = new SecureString();
//            do
//            {
//                info = Console.ReadKey(true);
//                if (info.Key != ConsoleKey.Enter)
//                {
//                    securePassword.AppendChar(info.KeyChar);
//                }
//            }
//            while (info.Key != ConsoleKey.Enter);
//            return securePassword;
//        }
//    }
//}
