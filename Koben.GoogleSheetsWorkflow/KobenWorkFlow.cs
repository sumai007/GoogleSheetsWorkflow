using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using Umbraco.Forms.Core;
using Umbraco.Forms.Core.Attributes;
using Umbraco.Forms.Core.Controllers;
using Umbraco.Forms.Core.Enums;
using Umbraco.Web;
using Umbraco.Core.Logging;


namespace Koben.GoogleSheetsWorkFlow
{
    public class KobenWorkFlow : WorkflowType
    {
        [Umbraco.Forms.Core.Attributes.Setting("Google Spread Sheet ID", alias = "spreadsheetID", description = "This is the Id of the spreadsheet you want data from this form to go")]
        public string SpreadsheetID { get; set; }

        private readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private readonly string ApplicationName = "WorkFlow Test Class";
        string path = System.Configuration.ConfigurationManager.AppSettings["CredentialPath"]; // must be an relative path
        private SheetsService service;

        //Function for Umraco Workflow Icon
        public KobenWorkFlow()
        {
            Name = "Send to Google Spreadsheet Workflow";
            Id = new Guid("e4e60c47-eb0b-4ab2-8739-bd48a3312127");
            Description = "Update Spread sheet";
            Icon = "icon-message";
            Group = "Google Sheets";
            ValidateAndReturnPath(SpreadsheetID, path);
            
        }

        public override List<Exception> ValidateSettings()
        {
            List<Exception> exceptions = new List<Exception>();

            if (string.IsNullOrWhiteSpace(SpreadsheetID))
                exceptions.Add(new Exception("The SpreadsheetID is required  "));

            

            try
            {
                ValidateAndReturnPath(SpreadsheetID, path);
            }
            catch (Exception ex)
            {
                exceptions.Add(ex);
            }

            return exceptions;
        }

        //Main function, record holds the form data
        public override WorkflowExecutionStatus Execute(Record record, RecordEventArgs e)
        {
            try
            {
         //       GoogleCredential credential;
                string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, System.Configuration.ConfigurationManager.AppSettings["CredentialPath"]);

                /*          using (var stream =
                           new FileStream(path, FileMode.Open, FileAccess.Read))
                          {
                              // The file token.json stores the user's access and refresh tokens, and is created
                              // automatically when the authorization flow completes for the first time.
                              credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);

                          }
                          */

                UserCredential credential;
                
                using (var stream =
                    new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"token.json");
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            Scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;


                }
                // Create Google Sheets API service.
                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });


                ReadSheet(record);                 //Read Sheet If empty add headers
                CheckForNewHeader(record);         //Check for new If only new fileds ot the form are added update the sheet Note: No action taken if fields fromm the form are deleted 
                AddData(record);                  // Add data to the sheet

                return WorkflowExecutionStatus.Completed;
            }
            catch (Exception ex)
            {
                LogHelper.Info(typeof(KobenWorkFlow), ex.ToString());
                return WorkflowExecutionStatus.Failed;
            }

        }

        //Add Record function
        private void AddData(Record record)
        {
            var range = "A:Z";
            var valueRange = new ValueRange();

            var objectlist = new List<object>();
            foreach (var field in record.RecordFields)
            {

                objectlist.Add(field.Value.Values.FirstOrDefault());


            }

            IList<IList<object>> values = new List<IList<object>> { objectlist };
            valueRange.Values = values;


            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();

        }

        // Create Header function
        private void Createheaders(Record record)
        {
            var range = "A:Z";
            var valueRange = new ValueRange();

            var objectlist = new List<object>();
            foreach (var field in record.RecordFields)
            {

                objectlist.Add(field.Value.Alias);


            }

            IList<IList<object>> values = new List<IList<object>> { objectlist };
            valueRange.Values = values;


            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetID, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();

        }

        private void ReadSheet(Record record)
        {
            var range = "A1:Z1";
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(SpreadsheetID, range);
            var response = request.Execute();

            IList<IList<object>> values = response.Values;

            if (values == null)
            {
                Createheaders(record);
            }
        }

        // Considering the Fields are only added to the Umbraco form 
        private void CheckForNewHeader(Record record)
        {

            var range = "A1:Z1";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(SpreadsheetID, range);

            var response = request.Execute();
            // IList<IList<object>> values = response.Values;

            if (record.RecordFields.Count == response.Values[0].Count)
            {
                return;
            }
            else if (record.RecordFields.Count > response.Values[0].Count)
            {
                CreateNewheaders(record);
            }
            else
            {
                return;
            }

        }

        private void CreateNewheaders(Record record)
        {
            var range = "A1:Z1";
            var valueRange = new ValueRange();

            var objectlist = new List<object>();
            foreach (var field in record.RecordFields)
            {

                objectlist.Add(field.Value.Alias);


            }

            IList<IList<object>> values = new List<IList<object>> { objectlist };
            valueRange.Values = values;


            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetID, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = updateRequest.Execute();

        }

        private string ValidateAndReturnPath(string SpreadsheetID, string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new Exception("Path to the Credential file is required refer to Config file");

            var tempPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);

            if (!File.Exists(tempPath))
                throw new Exception("The Credential file does not exists at: " + tempPath + ".Refer to Config file to change location");            

            return tempPath;
        }



        
    }
}
