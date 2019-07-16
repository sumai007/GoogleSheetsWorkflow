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

        private readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private readonly string ApplicationName = "WorkFlow Test Class";
        private readonly string SpreadsheetID;
        private readonly string path; // must be an relative path
        private readonly SheetsService service;

        //Function for Umraco Workflow Icon
        public KobenWorkFlow()
        {
            Name = "Spread ";
            Id = new Guid("e5e089b7-b760-4ea4-a038-049c3c91d4c2");
            Description = "Update Spread sheet";
            Icon = "icon-message";


            SpreadsheetID = System.Configuration.ConfigurationManager.AppSettings["GoogleSheetID"];
            path = System.Configuration.ConfigurationManager.AppSettings["CredentialPath"]; // must be an relative path

            path = ValidateAndReturnPath(SpreadsheetID, path);

            GoogleCredential credential;

            using (var stream =
                 new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);

            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


        }

        //Main function, record holds the form data
        public override WorkflowExecutionStatus Execute(Record record, RecordEventArgs e)
        {
            try
            {

                ReadSheet(record);
                AddData(record);

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
            var range = "A:Z";
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(SpreadsheetID, range);
            var response = request.Execute();

            IList<IList<object>> values = response.Values;

            if (values == null)
            {
                Createheaders(record);
            }
        }

        private string ValidateAndReturnPath(string SpreadsheetID, string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new Exception("Path to the Credential file is required refer to Config file");

            var tempPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, path);

            if (!File.Exists(tempPath))
                throw new Exception("The Credential file does not exists at: " + tempPath + ".Refer to Config file to change location");
            if (string.IsNullOrWhiteSpace(SpreadsheetID))
                throw new Exception("The SpreadsheetID is required refer to Config file ");

            return tempPath;
        }



        public override List<Exception> ValidateSettings()
        {
            var exceptions = new List<Exception>();
            return exceptions;
        }
    }
}
