using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks; 
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json; 
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ProactiveBot;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class CreateDialog : IDialog<object>
    {
        public async Task StartAsync(IDialogContext context)
        {
            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;
            await context.PostAsync(msg);
            List<string> excelType = new List<string>(){ "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                                            "application/vnd.ms-excel",
                                                            "application/vnd.ms-excel.sheet.macroEnabled.12",
                                                           };
            var message = context.MakeMessage();
            message.Text = "This is the template excel for you to fill up to create an event/survey.";

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            CloudBlobContainer blobContainer = blobClient.GetContainerReference("surveytemplates");
            blobContainer.CreateIfNotExists();
            CloudBlockBlob blob = blobContainer.GetBlockBlobReference("Survey_Template.xlsx");

            // Set the permissions so the blobs are public. //ON HOLD
            BlobContainerPermissions permissions = new BlobContainerPermissions
            {
                PublicAccess = BlobContainerPublicAccessType.Blob
            };
            blobContainer.SetPermissions(permissions);

            message.Attachments.Add(new Attachment()
            {
                ContentUrl = blob.Uri.ToString(),
                ContentType = blob.Properties.ContentType,
                Name = blob.Name
            });
            
            await context.PostAsync(message);
            PromptDialog.Attachment(context, this.ResumeAfterPrompt, "Send me the excel file whenever you are ready.", excelType);
            
        }

        private async Task ResumeAfterPrompt(IDialogContext context, IAwaitable<IEnumerable<Attachment>> result)
        {
            context.UserData.SetValue(ContextConstants.UserId, context.Activity.From.Id);
            try
            {
                //Connecting to storage
                var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
                var tableClient = storageAccount.CreateCloudTableClient();
                CloudTable eventTable = tableClient.GetTableReference("Event");
                eventTable.CreateIfNotExists();

                //Get the Attachment
                var message = await result;
                var msg = context.MakeMessage();
                msg.Type = ActivityTypes.Typing;
                await context.PostAsync(msg);
                var excelName = "";

                string destinationContainer = "surveytemplates";
                CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
                var blobContainer = blobClient.GetContainerReference(destinationContainer);
                blobContainer.CreateIfNotExists();
                String newFileName = "";

                //Receiving the attachment
                foreach (Attachment a in message)
                {
                    try
                    {
                        excelName = a.Name;

                        using (HttpClient httpClient = new HttpClient())
                        {
                            var responseMessage = await httpClient.GetAsync(a.ContentUrl);
                            var contentLenghtBytes = responseMessage.Content.Headers.ContentLength;
                            var fileByte = await httpClient.GetByteArrayAsync(a.ContentUrl);

                            //Set file name
                            newFileName = context.UserData.GetValue<string>(ContextConstants.UserId) + "_" + DateTime.Now.ToString("MMddyyyy-hhmmss") + "_" + excelName;

                            string sourceUrl = a.ContentUrl;

                            // Set the permissions so the blobs are public. //ON HOLD
                            BlobContainerPermissions permissions = new BlobContainerPermissions
                            {
                                PublicAccess = BlobContainerPublicAccessType.Blob
                            };
                            blobContainer.SetPermissions(permissions);

                            var newBlockBlob = blobContainer.GetBlockBlobReference(newFileName);

                            try
                            {
                                using (var memoryStream = new System.IO.MemoryStream(fileByte))
                                {
                                    newBlockBlob.UploadFromStream(memoryStream);
                                    newBlockBlob.Properties.ContentType = a.ContentType;
                                }
                                newBlockBlob.SetProperties();
                            }
                            catch (Exception ex)
                            {
                                await context.PostAsync("Something went wrong with uploading, please type \'/ce\' to try again.");
                                context.Done(this);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        await context.PostAsync("Something went wrong with uploading, please type \'/ce\' to try again.");
                        context.Done(this);
                    }
                }
                try
                {
                    //Opening the Excel File
                    CloudBlockBlob blob = blobContainer.GetBlockBlobReference(newFileName);

                    //Check for unique entity
                    DateTime thisDay = DateTime.Today;
                    String today = FormatDate(thisDay.ToString("d"));
                    
                    var eventCode = GenerateCode();

                    //After checking that there are no duplicates in generated codes
                    EventEntity eventEntity = new EventEntity(context.UserData.GetValue<string>(ContextConstants.UserId), eventCode);
                    
                    using (HttpClient httpClient = new HttpClient())
                    {
                        var fileByte = await httpClient.GetByteArrayAsync(blob.Uri.ToString());
                        try
                        {
                            XSSFWorkbook wb;
                            XSSFSheet sh;

                            using (var memoryStream = new MemoryStream(fileByte))
                            {
                                wb = new XSSFWorkbook(memoryStream);
                            }

                            sh = (XSSFSheet)wb.GetSheetAt(0);

                            //Check event start date and event name
                            if (sh.GetRow(2).GetCell(1) != null && sh.GetRow(2).GetCell(1).ToString() != "" && sh.GetRow(1).GetCell(1) != null && sh.GetRow(1).GetCell(1).ToString() != "")
                            {
                                //DateTime dateTime;
                                DataFormatter dataFormatter = new DataFormatter();
                                var date = dataFormatter.FormatCellValue(sh.GetRow(2).GetCell(1));
                                int totalDays = 1;
                                var error = false;

                                DateTime sDate;
                                DateTime eDate;
                                
                                //Try to parse date in the specified format
                                if (ParseDate(date, thisDay)) {
                                    String[] split = date.Split(new Char[] { '/', '-' });
                                    date = split[0].ToString() + '/' + split[1].ToString() + '/' + split[2];
                                    eventEntity.EventStartDate = date;
                                    DateTime.TryParseExact(dataFormatter.FormatCellValue(sh.GetRow(2).GetCell(1)), new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out sDate);

                                    //Check Event End Date
                                    if (sh.GetRow(3).GetCell(1) != null && sh.GetRow(3).GetCell(1).ToString() != "")
                                    {
                                        date = dataFormatter.FormatCellValue(sh.GetRow(3).GetCell(1));
                                        if (ParseDate(date, sDate))
                                        {
                                            split = date.Split(new Char[] { '/', '-' });
                                            date = split[0].ToString() + '/' + split[1].ToString() + '/' + split[2];
                                            eventEntity.EventEndDate = date;
                                            DateTime.TryParseExact(dataFormatter.FormatCellValue(sh.GetRow(3).GetCell(1)), new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out eDate);
                                            totalDays = (eDate - sDate).Days +1;
                                            if (totalDays > 31)
                                            {
                                                error = true;
                                                await context.PostAsync("Sorry, we currently do not support events that have more than 31 days.");
                                                await context.PostAsync(msg);
                                                await context.PostAsync("Talk to me again if you require my assistance.");
                                                context.Done(this);
                                            }
                                        }
                                        else
                                        {
                                            error = true;
                                            await context.PostAsync("Please check your excel file and input an end date in the format specified(dd-mm-yyyy or dd/mm/yyyy) and not before today's date!");
                                            await context.PostAsync(msg);
                                            await context.PostAsync("Talk to me again if you require my assistance.");
                                            context.Done(this);
                                        }
                                    }
                                    else
                                    {
                                        eventEntity.EventEndDate = eventEntity.EventStartDate;
                                    }
                                    //End of Checking Event's End Date

                                    //Check Attendance Code 1 Time
                                    var startTime = dataFormatter.FormatCellValue(sh.GetRow(4).GetCell(1));
                                    var endTime = dataFormatter.FormatCellValue(sh.GetRow(5).GetCell(1));

                                    var code1 = new List<string>();
                                    var code2 = new List<string>();

                                    if (sh.GetRow(4).GetCell(1) != null && sh.GetRow(4).GetCell(1).ToString() != "" && sh.GetRow(5).GetCell(1) != null && sh.GetRow(5).GetCell(1).ToString() != "")
                                    {
                                        if (ParseTime(startTime, endTime))
                                        {
                                            eventEntity.AttendanceCode1StartTime = startTime;
                                            eventEntity.AttendanceCode1EndTime = endTime;

                                            for (int i = 0; i < totalDays; i++)
                                            {
                                                //Generate unique attendanceCode1
                                                code1.Add(GenerateCode());
                                            }
                                        }
                                        else
                                        {
                                            error = true;
                                            await context.PostAsync("Please check your excel file and valid Attendance Code 1 Start/End Time in the format specified(2359). The end time should not be before the start time.");
                                            await context.PostAsync(msg);
                                            await context.PostAsync("Talk to me again if you require my assistance.");
                                            context.Done(this);
                                        }
                                    }
                                    else if (sh.GetRow(4).GetCell(1).ToString() == "" && sh.GetRow(5).GetCell(1).ToString() == "")
                                    {
                                        eventEntity.AttendanceCode1StartTime = "";
                                        eventEntity.AttendanceCode1EndTime = "";
                                        
                                    }
                                    else
                                    {
                                        error = true;
                                        await context.PostAsync("Please check your excel file and valid Attendance Code 1 Start/End Time in the format specified(2359). The end time should not be before the start time.");
                                        await context.PostAsync(msg);
                                        await context.PostAsync("Talk to me again if you require my assistance.");
                                        context.Done(this);
                                    }
                                    //End of Attendance Code 1 Time

                                    //Check Attendance Code 2 Time
                                    startTime = dataFormatter.FormatCellValue(sh.GetRow(6).GetCell(1));
                                    endTime = dataFormatter.FormatCellValue(sh.GetRow(7).GetCell(1));
                                    if (sh.GetRow(6).GetCell(1) != null && sh.GetRow(6).GetCell(1).ToString() != "" && sh.GetRow(7).GetCell(1) != null && sh.GetRow(7).GetCell(1).ToString() != "")
                                    {
                                        if (ParseTime(startTime, endTime))
                                        {
                                            eventEntity.AttendanceCode2StartTime = startTime;
                                            eventEntity.AttendanceCode2EndTime = endTime;

                                            for (int i = 0; i < totalDays; i++)
                                            {
                                                //Generate unique attendanceCode2
                                                code2.Add(GenerateCode());
                                            }
                                        }
                                        else
                                        {
                                            error = true;
                                            await context.PostAsync("Please check your excel file and valid Attendance Code 2 Start/End Time in the format specified(2359). The end time should not be before the start time.");
                                            await context.PostAsync(msg);
                                            await context.PostAsync("Talk to me again if you require my assistance.");
                                            context.Done(this);
                                        }
                                    }
                                    else
                                    {
                                        eventEntity.AttendanceCode2StartTime = "";
                                        eventEntity.AttendanceCode2EndTime = "";
                                    }
                                    //End of Attendance Code 2 Time

                                    //Check Survey End Date
                                    if (sh.GetRow(8).GetCell(1) != null && sh.GetRow(8).GetCell(1).ToString() != "")
                                    {
                                        date = dataFormatter.FormatCellValue(sh.GetRow(8).GetCell(1));
                                        if (ParseDate(date, sDate))
                                        {
                                            split = date.Split(new Char[] { '/', '-' });
                                            date = split[0] + '/' + split[1] + '/' + split[2];
                                            eventEntity.SurveyEndDate = date;
                                        }
                                        else//Input date is not in the specified format
                                        {
                                            error = true;
                                            await context.PostAsync("Please check your excel file and input a Survey End Date in the format specified(dd-mm-yyyy or dd/mm/yyyy) and not before today's date!");
                                            await context.PostAsync(msg);
                                            await context.PostAsync("Talk to me again if you require my assistance.");
                                            context.Done(this);
                                        }
                                    }
                                    else
                                    {
                                        if (eventEntity.EventEndDate != "")
                                        {
                                            eventEntity.SurveyEndDate = eventEntity.EventEndDate;
                                        }
                                        else
                                        {
                                            eventEntity.SurveyEndDate = eventEntity.EventStartDate;
                                        }
                                    }
                                    //End of checking Survey End Date

                                    //Check Survey End Time
                                    endTime = dataFormatter.FormatCellValue(sh.GetRow(9).GetCell(1));
                                    if (sh.GetRow(9).GetCell(1) != null && sh.GetRow(9).GetCell(1).ToString() != "")
                                    {
                                        if (ParseTime("0000", endTime))
                                        {
                                            eventEntity.SurveyEndTime = endTime;
                                        }
                                        else
                                        {
                                            error = true;
                                            await context.PostAsync("Please check your excel file and your Survey End Time in the format specified(2359).");
                                            await context.PostAsync(msg);
                                            await context.PostAsync("Talk to me again if you require my assistance.");
                                            context.Done(this);
                                        }
                                    }
                                    else
                                    {
                                        eventEntity.SurveyEndTime = "2359";
                                    }
                                    //End of Attendance Code 2 Time

                                    //Start of Anonymity
                                    if (sh.GetRow(11).GetCell(1) != null && sh.GetRow(11).GetCell(1).ToString() != "" && (sh.GetRow(11).GetCell(1).ToString() == "Y" || sh.GetRow(11).GetCell(1).ToString() == "N" || sh.GetRow(11).GetCell(1).ToString() == "Yes" || sh.GetRow(11).GetCell(1).ToString() == "No"))
                                    {
                                        eventEntity.Anonymous = (sh.GetRow(11).GetCell(1).ToString());
                                    }
                                    else
                                    {
                                        error = true;
                                        await context.PostAsync("Please check your excel file and input either Y or N for Anonymity.");
                                        await context.PostAsync(msg);
                                        await context.PostAsync("Talk to me again if you require my assistance.");
                                        context.Done(this);
                                    }
                                    //End of Anonymity

                                    //Start of Email
                                    if (sh.GetRow(12).GetCell(1) != null && sh.GetRow(12).GetCell(1).ToString() != "")
                                    {
                                        try
                                        {
                                            System.Net.Mail.MailAddress mail = new System.Net.Mail.MailAddress(sh.GetRow(12).GetCell(1).ToString());

                                            eventEntity.Email = (sh.GetRow(12).GetCell(1).ToString());
                                        }
                                        catch (FormatException)
                                        {
                                            error = true;
                                            await context.PostAsync("Please check your excel file and input a valid Email address or remove the field.");
                                            await context.PostAsync(msg);
                                            await context.PostAsync("Talk to me again if you require my assistance.");
                                            context.Done(this);
                                        }

                                    }
                                    else
                                    {
                                        eventEntity.Email = "";
                                    }
                                    //End of Email

                                    //If the file got no error parsing all the values
                                    if (!error)
                                    {
                                        //Setting eventEntity variables
                                        eventEntity.Description = dataFormatter.FormatCellValue(sh.GetRow(10).GetCell(1));
                                        eventEntity.EventName = (sh.GetRow(1).GetCell(1).ToString());

                                        if (sh.GetRow(13).GetCell(1).ToString() != "")
                                            eventEntity.Password = Sha256(sh.GetRow(13).GetCell(1).ToString());
                                        else
                                            eventEntity.Password = "";

                                        var survey = new List<QuestionEntity>();
                                        QuestionEntity question = new QuestionEntity();
                                        eventEntity.SurveyCode = GenerateCode();

                                        //retrieving the questions
                                        for (int i = 35; i <= sh.LastRowNum; i++)
                                        {
                                            var answerList = new List<string>();

                                            if (sh.GetRow(i).GetCell(0) != null && sh.GetRow(i).GetCell(0).ToString() != "")
                                            {
                                                int x = 1;
                                                question = new QuestionEntity(eventEntity.SurveyCode, (i - 34).ToString());
                                                question.QuestionText = (sh.GetRow(i).GetCell(0).ToString());

                                                //Retrieving the answer list
                                                while (sh.GetRow(i).GetCell(x) != null && sh.GetRow(i).GetCell(x).ToString() != "")
                                                {
                                                    answerList.Add(sh.GetRow(i).GetCell(x).ToString());
                                                    x++;
                                                }
                                            }
                                            else
                                            {
                                                break;
                                            }
                                            question.AnswerList = JsonConvert.SerializeObject(answerList);
                                            survey.Add(question);
                                        }
                                        //No question is found in excel file
                                        if (survey.Count == 0)
                                        {
                                            await context.PostAsync("Your excel file does not contain any question, please check again! (Read the guidelines in the provided template)");
                                            await context.PostAsync(msg);
                                            await context.PostAsync("Talk to me again if you require my assistance.");
                                            context.Done(this);
                                        }
                                        else//Questions were found in excel file, process inserting of event
                                        {
                                            eventEntity.Survey = JsonConvert.SerializeObject(survey);
                                            for(int i = 0; i < totalDays; i++)
                                            {
                                                if(code1.Count != 0)
                                                    eventEntity.AttendanceCode1 = code1[i];
                                                if (code2.Count != 0)
                                                    eventEntity.AttendanceCode2 = code2[i];
                                                string code = "";
                                                while (true)
                                                {
                                                    code = Guid.NewGuid().ToString().Substring(0, 8);
                                                    TableOperation retrieveOperation = TableOperation.Retrieve<EventEntity>(context.UserData.GetValue<string>(ContextConstants.UserId), code);
                                                    TableResult retrievedResult = eventTable.Execute(retrieveOperation);

                                                    if (retrievedResult.Result == null)
                                                    {
                                                        break;
                                                    }
                                                }
                                                eventEntity.RowKey = code;
                                                eventEntity.Day = (i+1).ToString();
                                                TableOperation insertOperation1 = TableOperation.InsertOrMerge(eventEntity);
                                                eventTable.Execute(insertOperation1);
                                            }
                                            
                                            await context.PostAsync(msg);
                                            string eventMsg = "'" + eventEntity.EventName + "' has been created, with a total of " + survey.Count + " questions. \n\n";
                                            eventMsg += "Date: " + eventEntity.EventStartDate + " to " + eventEntity.EventEndDate + "\n\n";
                                            eventMsg += "Attendance Code 1: \n\n";
                                            for (int i = 0; i < code1.Count; i++) {
                                                eventMsg += "Day " + (i+1) + ": " + (code1[i] == null ? "-" : code1[i]) + "\n\n";
                                            }
                                            eventMsg += "\n\n" + "Attendance Code 2:  \n\n";
                                            for (int i = 0; i < code2.Count; i++)
                                            {
                                                eventMsg += "Day " + (i+1) + ": " + (code2[i] == null ? "-" : code2[i]) + "\n\n";
                                            }
                                            eventMsg += "\n\n" + "Survey Code: " + eventEntity.SurveyCode + "\n\n\n\n"
                                                + "Type /Event to manage your event. Thank you and have a great day ahead!";

                                            await context.PostAsync(eventMsg);
                                            bool a = SharedFunction.EmailEvent(eventEntity, context.UserData.GetValue<string>(ContextConstants.Name), survey, code1, code2);
                                            context.Done(this);
                                        }
                                    }
                                }
                                else//Input date is not in the specified format
                                {
                                    await context.PostAsync("Please check your excel file and input a date in the format specified(dd-mm-yyyy or dd/mm/yyyy) and not before today's date!");
                                    await context.PostAsync(msg);
                                    await context.PostAsync("Talk to me again if you require my assistance.");
                                    context.Done(this);
                                }
                            }
                            else
                            {
                                await context.PostAsync("Please check your excel file to see if you have inputted an Event Name and Event Start Date!");
                                await context.PostAsync(msg);
                                await context.PostAsync("Talk to me again if you require my assistance.");
                                context.Done(this);
                            }
                        }
                        catch (Exception ex)
                        {
                            await context.PostAsync("Sorry there is an error trying to open your file, please type \'/ce\' to try again.");
                            context.Done(this);
                        }
                    }
                }
                catch (Exception e){

                    await context.PostAsync("Sorry there is an error trying to open your file, please type \'/ce\' to try again.");
                    context.Done(this);
                }
            }
            catch (TooManyAttemptsException e)
            {
            }
        }

        private bool ParseDate(String date, DateTime compareDate)
        {
            DateTime dateTime;
            if (DateTime.TryParseExact(date, new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture,
                                DateTimeStyles.None, out dateTime))
            {
                if (DateTime.Compare(compareDate, dateTime) <= 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private bool ParseTime(String startTime, string endTime)
        {
            DateTime sDateTime;
            DateTime eDateTime;
            if (DateTime.TryParseExact(startTime, new string[] { "HHmm" }, CultureInfo.InvariantCulture,
                                DateTimeStyles.None, out sDateTime))
            {
                if (DateTime.TryParseExact(endTime, new string[] { "HHmm" }, CultureInfo.InvariantCulture,
                                DateTimeStyles.None, out eDateTime))
                {
                    if (DateTime.Compare(sDateTime, eDateTime) <= 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        //Generate a unique code
        private string GenerateCode()
        {
            string generatedCode;
            DateTime thisDay = DateTime.Today;
            String today = FormatDate(thisDay.ToString("d"));
            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            CloudTable eventTable = tableClient.GetTableReference("Event");
            eventTable.CreateIfNotExists();
            while (true)
            {
                generatedCode = Guid.NewGuid().ToString().Substring(0, 8);
                String filterA = TableQuery.GenerateFilterCondition("AttendanceCode1", QueryComparisons.Equal, generatedCode);
                String filterB = TableQuery.GenerateFilterCondition("AttendanceCode2", QueryComparisons.Equal, generatedCode);
                String filterC = TableQuery.GenerateFilterCondition("SurveyCode", QueryComparisons.Equal, generatedCode);
                TableQuery<EventEntity> query = new TableQuery<EventEntity>().Where(TableQuery.CombineFilters(TableQuery.CombineFilters(filterA, TableOperators.Or, filterB), TableOperators.Or, filterC));
                var results = eventTable.ExecuteQuery(query);
                var empty = true;
                foreach (EventEntity e in results)
                {
                    empty = false;
                }
                if (empty)
                {
                    break;
                }
            }
            return generatedCode;
        }

        private string Sha256(string randomString)
        {
            var crypt = new System.Security.Cryptography.SHA256Managed();
            var hash = new StringBuilder();
            byte[] crypto = crypt.ComputeHash(Encoding.UTF8.GetBytes(randomString));
            foreach (byte theByte in crypto)
            {
                hash.Append(theByte.ToString("x2"));
            }
            return hash.ToString();
        }

        private String FormatDate(String today)
        {
            String[] split = today.Split('/');
            today = split[1] + '/' + split[0] + '/' + split[2];
            return today;
        }
    }
}