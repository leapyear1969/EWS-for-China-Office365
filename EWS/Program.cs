using System;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System.Configuration;

namespace EWS
{
    class Program
	{

        static async System.Threading.Tasks.Task Main(string[] args)
        {
            // Using Microsoft.Identity.Client 4.22.0


            // Configure the MSAL client to get tokens
            var pcaOptions = new PublicClientApplicationOptions
            {
                ClientId = ConfigurationManager.AppSettings["appId"],
                TenantId = ConfigurationManager.AppSettings["tenantId"],
                //Instance = ConfigurationManager.AppSettings["https://login.chinacloudapi.cn"]
                AzureCloudInstance = AzureCloudInstance.AzureChina
            };

            var pca = PublicClientApplicationBuilder.CreateWithApplicationOptions(pcaOptions).WithRedirectUri("http://localhost").Build();

            // The permission scope required for EWS access
            var ewsScopes = new string[] { "https://partner.outlook.cn/EWS.AccessAsUser.All" };
            //var ewsScopes = new string[] { "https://outlook.office365.com/EWS.AccessAsUser.All" };


            try
            {
                // Make the interactive token request
                var authResult = await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService();
                ewsClient.Url = new Uri("https://partner.outlook.cn/EWS/Exchange.asmx");

                //ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

                ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);

                // Make a simple EWS call
                var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));
                foreach (var folder in folders)
                {
                    Console.WriteLine($"Folder: {folder.DisplayName}");
                }
                Console.WriteLine("************************************************************************************************************************");

                //Get Calendar View 
                DateTime startDate = DateTime.Now;
                DateTime endDate = startDate.AddDays(30);
                const int NUM_APPTS = 5;
                //const string ATTENDEE = null;
                //initialize the calendar folder obj with folder ID
                CalendarFolder calendar = CalendarFolder.Bind(ewsClient,WellKnownFolderName.Calendar,new PropertySet());
                //set property to get appointments
                CalendarView cView = new CalendarView(startDate,endDate,NUM_APPTS);
                //set propertySet
                cView.PropertySet = new PropertySet(AppointmentSchema.Subject,AppointmentSchema.Start,AppointmentSchema.End,AppointmentSchema.ICalUid,AppointmentSchema.Id);

                FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);
                Console.WriteLine("\nThe first " + NUM_APPTS + " appointments on your calendar from " + startDate.Date.ToShortDateString() +
                                  " to " + endDate.Date.ToShortDateString() + " are: \n");
                foreach (Appointment a in appointments)
                {
                    Console.Write("Subject: " + a.Subject.ToString() + " ");
                    Console.Write("Start: " + a.Start.ToString() + " ");
                    Console.Write("End: " + a.End.ToString() + " ");
                    Console.Write("日历ID:"+ a.ICalUid.ToString() + " ");
                    Console.Write("会议ID:" + a.Id.ToString());
                    Console.WriteLine("\n");
                    Console.WriteLine();
                }
                Console.WriteLine("************************************************************************************************************************");
                
                
                
                //Get Appointment Resource
                Appointment meeting = Appointment.Bind(ewsClient, new ItemId("AQMkADhjNjU4YTY0LWFiOGUtNGMxNy1hOGYyLTRhNDI1ZDRiMzhkOQBGAAADkQf6i4PrrkWzDaTB9jTocgcAPtq8uE+Z902/50OXFrFp3QAAAgENAAAAPtq8uE+Z902/50OXFrFp3QABR+cTkAAAAA=="));
                //Appointment apts = Appointment.Bind(ewsClient, meeting.Id, new PropertySet(AppointmentSchema.Resources)) ;
                for (int i = 0; i < meeting.Resources.Count; i++) { 
                       Console.WriteLine("资源邮箱是 (" + meeting.Resources[i].Address + ")" + ":" + meeting.Resources[i].ResponseType.Value.ToString());
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
