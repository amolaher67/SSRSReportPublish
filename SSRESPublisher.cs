using Microsoft.Deployment.WindowsInstaller;
using Principa.Deployment.Helpers.webserfrence;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Web.Services.Protocols;

namespace Principa.Deployment.Helpers
{
    /// <summary>
    /// SSRS publish using
    /// help :https://docs.microsoft.com/en-us/dotnet/api/reportservice2010.reportingservice2010?view=sqlserver-2016
    /// </summary>
    public static class SSRSPublisher
    {
        /// <summary>
        /// Ping SSRS serer report url to check is ReportService2010 service in SSRS server is running.
        /// </summary>
        /// <param name="session"></param>
        /// <param name="reportServerUrl"></param>
        public static void PingReportServerUrl(Session session, string reportServerUrl,string reportPortUrl)
        {
            session["ERRORMESSAGEDATA"] = string.Empty;
            session["ISSUCCESS"] = "";

            string getUrl = string.Empty;
            if (reportServerUrl != null && reportServerUrl.EndsWith("/"))
            {
                getUrl = $"{reportServerUrl}ReportService2010.asmx";
            }
            else
            {
                getUrl = $"{reportServerUrl}/ReportService2010.asmx";
            }
            
            PingReportServer(session ,getUrl, ErrorMessages.ReportUrlInvalid(getUrl));
            if (session["ERRORMESSAGEDATA"] != "") return;
            PingReportServer(session, reportPortUrl, ErrorMessages.ReportPortalUrlInvalid(reportPortUrl));
            
        }

        private static void PingReportServer(Session session ,string getUrl,string errorMessage)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(getUrl);
                request.UseDefaultCredentials = true;

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    if (response.StatusCode != HttpStatusCode.OK)
                    {
                        session["ERRORMESSAGEDATA"] = errorMessage;
                    }
                }
            }
            catch (Exception ex)
            {
                session.Log(ex.StackTrace);
                session["ERRORMESSAGEDATA"] = errorMessage;
            }
        }

        /// <summary>
        /// Main method - Publish Reports using below method
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="reportServerUrl"></param>
        /// <param name="connectionString"></param>
        /// <param name="folderName"></param>
        public static void Publish(string sourcePath, string reportServerUrl, string connectionString, string folderName = "DecisionSmartV4")
        {
            //copy company logo from selected location to reportsFolder
            using (ReportingService2010 client = new ReportingService2010())
            {
                client.Url = $"{reportServerUrl}/ReportService2010.asmx?wsdl";
                client.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;
                CreateFolder(client, folderName);
                CreateDataSource(client, connectionString, folderName);
                UploadRdlFiles(sourcePath, client, folderName);
            }
        }

        /// <summary>
        /// Read all available files under report folder for publish
        /// </summary>
        /// <param name="reportsFolderPath"></param>
        /// <returns></returns>
        private static FileInfo[] GetReportsFiles(string reportsFolderPath)
        {
            DirectoryInfo d = new DirectoryInfo(reportsFolderPath);//Assuming Test is your Folder
            FileInfo[] files = d.GetFiles("*"); //Getting Text files
            return files;
        }

        /// <summary>
        /// Create Folder to place our report on Root directory
        /// </summary>
        /// <param name="client"></param>
        /// <param name="folderName"></param>
        private static void CreateFolder(ReportingService2010 client, string folderName)
        {
            var propArray = new Property[1];
            propArray[0] = new Property()
            {
                Value = folderName,
                Name = folderName
            };

            //User credential for
            //Reporting Service
            //the current logged system user
            var catalogItems = client.ListChildren("/", true);
            if (catalogItems.FirstOrDefault(s => s.Name == folderName) == null)
                client.CreateFolder(folderName, "/", propArray);
        }

        /// <summary>
        /// Upload all files on SSRS server using ReportService
        /// </summary>
        /// <param name="reportsFolderPath"></param>
        /// <param name="client"></param>
        /// <param name="folderName"></param>
        private static void UploadRdlFiles(string reportsFolderPath, ReportingService2010 client, string folderName)
        {
            //Get All Files From Reports Directory
            var reportFiles = GetReportsFiles(reportsFolderPath); //ConfigurationSettings.AppSettings["ReportsFolderPath"]);
            foreach (FileInfo fileInfo in reportFiles)
            {
                try
                {
                    FileStream stream = fileInfo.OpenRead();
                    var definition = new Byte[stream.Length];
                    stream.Read(definition, 0, (int)stream.Length);
                    stream.Close();
                    string itemType = "Report";
                    Property[] prop = null;
                    if (fileInfo.Extension != ".rdl")
                    {
                        itemType = "Resource";
                        prop = new Property[2];
                        prop[0] = new Property
                        {
                            Name = "MIMEType",
                            Value = "image/jpeg"
                        };
                        prop[1] = new Property
                        {
                            Name = "Hidden",
                            Value = "true"
                        };
                    }

                    //only rdl and jpg needs be published
                    if (fileInfo.Extension == ".rdl" || fileInfo.Extension == ".jpg")
                    {
                        var report = client.CreateCatalogItem(itemType, fileInfo.Name, $"/{folderName}", true,
                            definition,
                            prop, Warnings: out Warning[] warnings);

                        if (report != null)
                        {
                            Console.WriteLine(fileInfo.Name + " Published Successfully ");
                            Console.WriteLine(string.Format("\n"));
                        }

                        if (warnings != null)
                        {
                            foreach (Warning warning in warnings)
                            {
                                Console.WriteLine($"Report: {warning.Message} has warnings");
                                Console.WriteLine(string.Format("\n"));
                            }
                        }
                        else
                            Console.WriteLine($"Report: {fileInfo.Name} created successfully with no warnings");
                    }
                    Console.WriteLine("\n");
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        /// <summary>
        /// Create DataSource
        /// </summary>
        /// <param name="client"></param>
        /// <param name="connectionString"></param>
        /// <param name="folderName"></param>
        private static void CreateDataSource(ReportingService2010 client, string connectionString, string folderName)
        {
            try
            {
                string name = "ReportDataSource";
                string parent = $"/{folderName}";

                DataSourceDefinition definition = new DataSourceDefinition
                {
                    CredentialRetrieval = CredentialRetrievalEnum.Integrated,
                    ConnectString = connectionString,
                    Enabled = true,
                    EnabledSpecified = true,
                    Extension = "SQL",
                    ImpersonateUserSpecified = false,
                    Prompt = null,
                    WindowsCredentials = false
                };
                //Use the default prompt string.  
                try
                {
                    client.CreateDataSource(name, parent, true, definition, null);
                }
                catch (SoapException e)
                {
                    Console.WriteLine(e.Detail.InnerXml.ToString());
                    throw e;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }
    }
}
