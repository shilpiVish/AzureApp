using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Net.Http;
using System.Text;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Collections.Generic;

namespace AzureFunctionApp
{
    public static class SendMail
    {
        [FunctionName("SendMail")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = "0b06fdda-0458-ea11-a811-000d3a0a7552";// req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

            // Acquiring Access Token  
            var accessToken = await AccessTokenGenerator();

            RootObject data = JsonConvert.DeserializeObject<RootObject>(requestBody);

            var client = new HttpClient();
            // var message = new HttpRequestMessage(HttpMethod.Get, "https://afflue.api.crm8.dynamics.com/api/data/v9.1/contacts?$top=1");
            var message1 = new HttpRequestMessage(HttpMethod.Post, "accounts(0b06fdda-0458-ea11-a811-000d3a0a7552)/Microsoft.Dynamics.CRM.new_CallAction2f6db438c569ea11a812000d3a0a7552");

            var message = new HttpRequestMessage(HttpMethod.Post, "accounts(0b06fdda-0458-ea11-a811-000d3a0a7552)/new_CallAction2f6db438c569ea11a812000d3a0a7552");

            var accountdata = new AccountData();
            accountdata.accountid = "0b06fdda-0458-ea11-a811-000d3a0a7552";// data.PrimaryEntityId;
            accountdata.odataType = "Microsoft.Dynamics.CRM.account";

            parameter p = new parameter();
            p.Account = accountdata;

            // OData related headers  
            message.Headers.Add("OData-MaxVersion", "4.0");
            message.Headers.Add("OData-Version", "4.0");
            message.Headers.Add("Prefer", "odata.include-annotations=\"*\"");
            message.Headers.Add("Authorization", $"Bearer {accessToken}");

            // Adding body content in HTTP request   
            if (p != null)
                message.Content = new StringContent(JsonConvert.SerializeObject(p), UnicodeEncoding.UTF8, "application/json");
            var response = await client.SendAsync(message);
            return new OkObjectResult(response.Content.ReadAsStringAsync().Result);
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

    }
    public class parameter
    {
        public object Account { get; set; }
    }
    public class AccountData
    {
        public string accountid { get; set; }

        [JsonProperty("@odata.type")]
        public string odataType { get; set; }
    }
    public class RootObject
    {
        public string BusinessUnitId { get; set; }
        public string CorrelationId { get; set; }
        public int Depth { get; set; }
        public string InitiatingUserAzureActiveDirectoryObjectId { get; set; }
        public string InitiatingUserId { get; set; }
        public List<object> InputParameters { get; set; }
        public bool IsExecutingOffline { get; set; }
        public bool IsInTransaction { get; set; }
        public bool IsOfflinePlayback { get; set; }
        public int IsolationMode { get; set; }
        public string MessageName { get; set; }
        public int Mode { get; set; }
        public DateTime OperationCreatedOn { get; set; }
        public string OperationId { get; set; }
        public string OrganizationId { get; set; }
        public string OrganizationName { get; set; }
        public List<object> OutputParameters { get; set; }
        public object OwningExtension { get; set; }
        public object ParentContext { get; set; }
        public List<object> PostEntityImages { get; set; }
        public List<object> PreEntityImages { get; set; }
        public string PrimaryEntityId { get; set; }
        public string PrimaryEntityName { get; set; }
        public string RequestId { get; set; }
        public string SecondaryEntityName { get; set; }
        public List<object> SharedVariables { get; set; }
        public int Stage { get; set; }
        public string UserAzureActiveDirectoryObjectId { get; set; }
        public string UserId { get; set; }
    }

}