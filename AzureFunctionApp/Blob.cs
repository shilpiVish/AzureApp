using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;

namespace AzureFunctionApp
{
    public static class Blob
    {
        [FunctionName("Blob")]
        public static void Run([BlobTrigger("blobcontainer/{name}", Connection = "connection")]Stream myBlob, string name, ILogger log)
        {
            string requestUri = "https://afflue.api.crm8.dynamics.com/api/data/v9.1/contacts";
            log.LogInformation($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");
            var res = CrmRequest(requestUri).Result;
        }
        public static async Task<string> AccessTokenGenerator()
        {
            string clientId = "95acd67c-8dec-4dfe-a1d9-ba46d116fe79"; // Your Azure AD Application ID
            string clientSecret = "Qx:4iBPckgkfZndmd:P--VB5924u7lT6"; // Client secret generated in your App
            string authority = "https://login.microsoftonline.com/1554ade9-21b5-4d2b-bb4e-1483135739ca"; // Azure AD App Tenant ID
            string resourceUrl = "https://afflue.crm8.dynamics.com"; // Your Dynamics 365 Organization URL

            var credentials = new ClientCredential(clientId, clientSecret);
            var authContext = new Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext(authority);
            var result = await authContext.AcquireTokenAsync(resourceUrl, credentials);
            return result.AccessToken;
        }
        public static async Task<bool> CrmRequest(string requestUri, string body = null)
        {
            var response = new HttpResponseMessage();
            var accessToken = await AccessTokenGenerator();
            var client = new HttpClient();
            var message = new HttpRequestMessage(HttpMethod.Post, requestUri);
            message.Headers.Add("OData-MaxVersion", "4.0");
            message.Headers.Add("OData-Version", "4.0");
            message.Headers.Add("Prefer", "odata.include-annotations=\"*\"");
            message.Headers.Add("Authorization", $"Bearer {accessToken}");
            var accountdata = GetEmployees();
           
           foreach(var item in accountdata)
            {
                if (item != null)
                {
                    try
                    {
                        body = JsonConvert.SerializeObject(item);
                        message.Content = new StringContent(body, UnicodeEncoding.UTF8, "application/json");
                        response = await client.SendAsync(message);
                    }
                    catch(Exception ex)
                    {
                        return false;
                    }
                }
            }
            return false;
            // Adding body content in HTTP request 
        }
        public static List<Employee> GetEmployees()
        {
            List<Employee> emp = new List<Employee>();
            DataTable dt = new DataTable();
            //        CloudStorageAccount storageAccount = CloudStorageAccount.Parse(
            //CloudConfigurationManager.GetSetting("StorageConnectionString"));

            //        // Create the blob client.
            //        CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();

            //        // Retrieve reference to a previously created container.
            //        CloudBlobContainer container = blobClient.GetContainerReference("files");

            //        // Retrieve reference to a blob named "imex.xlsx".
            //        CloudBlockBlob blockBlob = container.GetBlockBlobReference("book.xlsx");

            //        // Save blob contents to a file.
            //        //using (var fileStream = System.IO.File.OpenWrite(@"path\myfile"))
            //        //{
            //        //    blockBlob.DownloadToStream(fileStream);
            //        //}
            using (XLWorkbook workBook = new XLWorkbook(@"C:\Users\akumar3\Documents\Book1.xlsx"))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        emp = new List<Employee>();
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }
                for (int count = 0; count < dt.Rows.Count; count++)
                {
                    emp.Add(new Employee() { emailaddress = (dt.Rows[count][0].ToString()), name = dt.Rows[count][1].ToString() });
                }
            }
            return emp;
        }
    }
    public class parameters
    {
        public List<AccountEntity> ae { get; set; }
        public object acc { get; set; }
    }
    public class AccountEntity
    {
        public string name { get; set; }
        public string emailaddress1 { get; set; }
    }
    public class Employee
    {
        public string emailaddress { get; set; }
        public string name { get; set; }
    }
}
/*
 * call method
 * var contacts = CrmRequest(
         HttpMethod.Post,
         "https://afflue.api.crm8.dynamics.com/api/data/v9.1/contacts")
         .Result.Content.ReadAsStringAsync();
 */
