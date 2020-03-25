using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;


namespace AzureFunctionApp
{
    public static class Blob
    {
        [FunctionName("Blob")]
        public static void Run([BlobTrigger("blobcontainer/{name}", Connection = "AzureWebJobsStorage")]Stream myBlob, string name, ILogger log)
        {
            string requestUri = "https://afflue.api.crm8.dynamics.com/api/data/v9.1/accounts";
            log.LogInformation($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");

            var res = GetExcelBlobData("Book.xlsx", "DefaultEndpointsProtocol=https;AccountName=agsblobstorage2285;AccountKey=8YVtcRpENmsLtRlEiLbXEWRDp61/VaBwyOCtoUtVrjQz1qe2dFSW/H0QJtndH5uQmUH5VAE3QMdh0RN2DxzyyQ==;EndpointSuffix=core.windows.net", "blobcontainer", requestUri).Result;
        }
        public static async Task<string> AccessTokenGenerator()
        {
            string clientId = "95acd67c-8dec-4dfe-a1d9-ba46d116fe79";// Your Azure AD Application ID
            string clientSecret = "Qx:4iBPckgkfZndmd:P--VB5924u7lT6"; // Client secret generated in your App
            string authority = "https://login.microsoftonline.com/1554ade9-21b5-4d2b-bb4e-1483135739ca";// Azure AD App Tenant ID
            string resourceUrl = "https://afflue.crm8.dynamics.com"; // Your Dynamics 365 Organization URL


            var credentials = new ClientCredential(clientId, clientSecret);
            var authContext = new Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext(authority);
            var result = await authContext.AcquireTokenAsync(resourceUrl, credentials);
            return result.AccessToken;
        }
        public static async Task<bool> CrmRequest(string requestUri, List<Employee> accountdata, string body = null)
        {
            //string token = AccessTokenGenerator().Result;
            var response = new HttpResponseMessage();
            var accessToken = await AccessTokenGenerator();
            var client = new HttpClient();
            AccountEntity accountEntity = new AccountEntity();
            HttpRequestMessage message = null;
            foreach (var item in accountdata)
            {
                if (item != null)
                {
                    try
                    {
                       message= new HttpRequestMessage(HttpMethod.Post, requestUri);
                        message.Headers.Add("OData-MaxVersion", "4.0");
                        message.Headers.Add("OData-Version", "4.0");
                        message.Headers.Add("Prefer", "odata.include-annotations=\"*\"");
                        message.Headers.Add("Authorization", $"Bearer {accessToken}");
                        accountEntity.emailaddress1 = item.name;
                        accountEntity.name = item.emailaddress;
                        body = JsonConvert.SerializeObject(accountEntity);
                        message.Content = new StringContent(body, UnicodeEncoding.UTF8, "application/json");
                        response = await client.SendAsync(message);
                        var result = response.Content.ReadAsStringAsync();
                      
                    }
                    catch (Exception ex)
                    {
                        return false;
                    }
                }
            }
            return false;
        }


        private static async Task<Employee> GetExcelBlobData(string filename, string connectionString, string containerName, string requestUri)
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            CloudBlobContainer container = blobClient.GetContainerReference(containerName);
            CloudBlockBlob blockBlobReference = container.GetBlockBlobReference(filename);
            Employee emp = new Employee();
            DataSet ds = new DataSet();
            string path = @"C:\Shilpi\";
            string pathfile = path + "book504.xlsx";
            if (!File.Exists(path))
            {
                Directory.CreateDirectory(path);
                await blockBlobReference.DownloadToFileAsync(pathfile, FileMode.Create);
                var res = GetEmployees(pathfile);
                var re = CrmRequest(requestUri, res).Result;
                Directory.Delete(path, true);
            }
            return emp;
        }
        public static List<Employee> GetEmployees(string pathFile)
        {
            List<Employee> emp = new List<Employee>();
            DataTable dt = new DataTable();

            using (XLWorkbook workBook = new XLWorkbook(pathFile))
            {
                IXLWorksheet workSheet = workBook.Worksheet(1);
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
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
 * var contacts = CrmRequest(HttpMethod.Post,"https://afflue.api.crm8.dynamics.com/api/data/v9.1/contacts").Result.Content.ReadAsStringAsync();
 */






