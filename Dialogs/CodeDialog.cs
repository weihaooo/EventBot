using Microsoft.Bot.Builder.Dialogs;
using System;
using System.Threading.Tasks;
using Microsoft.Bot.Connector;
using Microsoft.WindowsAzure.Storage;
using System.Configuration;
using Microsoft.WindowsAzure.Storage.Table;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using Microsoft.WindowsAzure.Storage.Blob;
using ZXing;
using System.IO;
using System.Drawing;
using System.Diagnostics;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class CodeDialog : IDialog<object>
    { 
        public async Task StartAsync(IDialogContext context)
        { 
            String name = context.UserData.GetValue<String>(ContextConstants.Name);
            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;
            await context.PostAsync(msg);
            await context.PostAsync("My name is Eventory, and I am capable of managing event's attendance and survey! You can create an event with the command '/ce', view survey's result with '/event' and '/help' for assistance.");
            await context.PostAsync("Otherwise, please input the attendance or survey code. You can also upload a QR image to proceed.");
            
            context.Wait(this.MessageReceivedAsync);
        }

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> result)
        {
            var message = await result;

            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;
            await context.PostAsync(msg);

            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            CloudTable eventTable = tableClient.GetTableReference("Event");
            eventTable.CreateIfNotExists();

            DateTime thisDay = DateTime.Today;
            String today = FormatDate(thisDay.ToString("d"));

            EventEntity eventEntity = null;
            string codeEntered = "";

            if (message.Attachments.FirstOrDefault() != null)
            {
                string destinationContainer = "qrcode";
                CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
                var blobContainer = blobClient.GetContainerReference(destinationContainer);
                blobContainer.CreateIfNotExists();
                String newFileName = "";
                var a = message.Attachments.FirstOrDefault();

                try
                {
                    using (HttpClient httpClient = new HttpClient())
                    {
                        var responseMessage = await httpClient.GetAsync(a.ContentUrl);
                        var contentLenghtBytes = responseMessage.Content.Headers.ContentLength;
                        var fileByte = await httpClient.GetByteArrayAsync(a.ContentUrl);
                        
                        newFileName = context.UserData.GetValue<string>(ContextConstants.UserId) + "_" + DateTime.Now.ToString("MMddyyyy-hhmmss") + "_" + "qrcode";

                        var barcodeReader = new BarcodeReader();
                        var memoryStream = new MemoryStream(fileByte);
                        string sourceUrl = a.ContentUrl;
                        var barcodeBitmap = (Bitmap)Bitmap.FromStream(memoryStream);
                        var barcodeResult = barcodeReader.Decode(barcodeBitmap);
                        codeEntered = barcodeResult.Text;
                        eventEntity = GetEvent(barcodeResult.Text);
                        memoryStream.Close();
                    }
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.ToString());
                }
            }
            else
            {
                codeEntered = message.Text;
                eventEntity = GetEvent(message.Text);

            }
            //Check database
            if (eventEntity != null)
            {
                var eventName = eventEntity.EventName;

                CloudTable attendanceTable = tableClient.GetTableReference("Attendance");
                attendanceTable.CreateIfNotExists();

                String filterA = TableQuery.GenerateFilterCondition("SurveyCode", QueryComparisons.Equal, eventEntity.SurveyCode);
                String filterB = TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, context.UserData.GetValue<string>(ContextConstants.UserId));
                String filterD = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, eventEntity.RowKey);
                TableQuery<AttendanceEntity> query = new TableQuery<AttendanceEntity>().Where(TableQuery.CombineFilters(filterA, TableOperators.And, filterB));
                TableQuery<AttendanceEntity> todayQuery = new TableQuery<AttendanceEntity>().Where(TableQuery.CombineFilters(TableQuery.CombineFilters(filterA, TableOperators.And, filterB), TableOperators.And, filterD));

                var results = attendanceTable.ExecuteQuery(query);
                var todayResults = attendanceTable.ExecuteQuery(todayQuery);
                Boolean notDone = true;
                //Survey Code
                if (codeEntered == eventEntity.SurveyCode)
                {
                    if (results != null)
                    {
                        foreach (AttendanceEntity a in results)
                        {
                            if (a.Survey == true)
                            {
                                notDone = false;
                                await context.PostAsync("You have already completed the survey for " + "'" + eventName + "'!");
                                await context.PostAsync(msg);
                                await context.PostAsync("Talk to me again if you require my assistance.");
                                context.Done(this);
                            }
                        }

                        if (notDone && !ValidPeriod(eventEntity.SurveyEndDate, eventEntity.SurveyEndTime, eventEntity.EventStartDate))
                        {
                            await context.PostAsync("Sorry, this code is not yet ready for use or have already expired!");
                            await context.PostAsync(msg);
                            await context.PostAsync("Talk to me again if you require my assistance.");
                            context.Done(this);
                        }
                    }
                    context.UserData.SetValue(ContextConstants.Status, "3");
                    context.UserData.SetValue(ContextConstants.Survey, eventEntity.Survey);
                    //Attendance Code 1
                }
                else if (codeEntered == eventEntity.AttendanceCode1)
                {
                    if (todayResults != null)
                    {
                        foreach (AttendanceEntity a in todayResults)
                        {
                            if (a.Morning == true)
                            {
                                notDone = false;
                                await context.PostAsync("You have already registered this attendance for " + "'" + eventName + "'" + " today!");
                                await context.PostAsync(msg);
                                await context.PostAsync("Talk to me again if you require my assistance.");
                                context.Done(this);
                            }
                        }

                        if (notDone && !ValidTime(eventEntity.AttendanceCode1StartTime, eventEntity.AttendanceCode1EndTime, eventEntity.Day, eventEntity.EventStartDate, eventEntity.EventEndDate))
                        {
                            await context.PostAsync("Sorry, this code is not yet ready for use or have already expired!");
                            await context.PostAsync(msg);
                            await context.PostAsync("Talk to me again if you require my assistance.");
                            context.Done(this);
                        }
                    }
                    context.UserData.SetValue(ContextConstants.Status, "1");

                }//Attendance code 2
                else if (codeEntered == eventEntity.AttendanceCode2)
                {
                    if (todayResults != null)
                    {
                        foreach (AttendanceEntity a in todayResults)
                        {
                            if (a.Afternoon == true)
                            {
                                notDone = false;
                                await context.PostAsync("You have already registered this attendance for " + "'" + eventName + "'" + " today!");
                                await context.PostAsync(msg);
                                await context.PostAsync("Talk to me again if you require my assistance.");
                                context.Done(this);
                            }
                        }

                        if (notDone && !ValidTime(eventEntity.AttendanceCode2StartTime, eventEntity.AttendanceCode2EndTime, eventEntity.Day, eventEntity.EventStartDate, eventEntity.EventEndDate))
                        {
                            await context.PostAsync("Sorry, this code is not yet ready for use or have already expired!");
                            await context.PostAsync(msg);
                            await context.PostAsync("Talk to me again if you require my assistance.");
                            context.Done(this);
                        }
                    }

                    context.UserData.SetValue(ContextConstants.Status, "2");
                }
                context.UserData.SetValue(ContextConstants.EventCode, eventEntity.RowKey);
                context.UserData.SetValue(ContextConstants.SurveyCode, eventEntity.SurveyCode);
                context.UserData.SetValue(ContextConstants.Anonymous, eventEntity.Anonymous);
                context.UserData.SetValue(ContextConstants.EventName, eventName);
                context.UserData.SetValue(ContextConstants.Today, today);
                context.UserData.SetValue(ContextConstants.Description, eventEntity.Description);
                context.Done(true);

            }
            else
            {
                await context.PostAsync("I'm sorry, I think you have keyed in the wrong workshop code or have sent an invalid QR Code, let's do this again.");
                context.Wait(this.MessageReceivedAsync);
            }
        }
        
        private EventEntity GetEvent(String code)
        {
            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            CloudTable eventTable = tableClient.GetTableReference("Event");
            eventTable.CreateIfNotExists();
            
            String filterA = TableQuery.GenerateFilterCondition("AttendanceCode1", QueryComparisons.Equal, code);
            String filterB = TableQuery.GenerateFilterCondition("AttendanceCode2", QueryComparisons.Equal, code);
            String filterC = TableQuery.GenerateFilterCondition("SurveyCode", QueryComparisons.Equal, code);
            TableQuery<EventEntity> query = new TableQuery<EventEntity>().Where(TableQuery.CombineFilters(TableQuery.CombineFilters(filterA, TableOperators.Or, filterB), TableOperators.Or, filterC));
            var results = eventTable.ExecuteQuery(query);

            DateTime sDate;
            DateTime eDate;
            DateTime today = DateTime.Today;

            int dayNum = 1;

            List<EventEntity> e = results.ToList();
            
            if(e.Count == 0)
            {
                return null;
            } else if(e.Count == 1)
            {
                return e[0];
            }
            else
            {
                DateTime.TryParseExact(e[0].EventStartDate, new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out sDate);
                DateTime.TryParseExact(e[0].EventEndDate, new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out eDate);
                
                if(DateTime.Compare(sDate, today) <= 0 && DateTime.Compare(today, eDate) <= 0)
                {
                    dayNum = (today - sDate).Days + 1;

                    foreach (EventEntity entity in results)
                    {
                        if(entity.Day == dayNum.ToString())
                        {
                            return entity;
                        }
                    }
                    return null;
                }
                else
                {
                    return null;
                }
            }
        }

        private bool ValidPeriod(string endDate, string endTime, string startEventDate)
        {
            DateTime today = DateTime.Today;
            DateTime now = DateTime.Now.AddHours(8);
                
            DateTime sEventDate;
            DateTime dateTime;
            DateTime.TryParseExact(endDate + " " + endTime, new string[] { "dd/MM/yyyy HHmm", "dd-MM-yyyy HHmm" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime);
            DateTime.TryParseExact(startEventDate, new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out sEventDate);

            if (DateTime.Compare(dateTime,now) > 0 && DateTime.Compare(sEventDate,today) <=0)
            {
                return true;
            }
            return false;
        }

        private bool ValidTime(string attendanceCode1StartTime, string attendanceCode1EndTime, string day, string eventStartDate, string eventEndDate)
        {
            DateTime today = DateTime.Today;
            DateTime now = DateTime.Now.AddHours(8);

            DateTime sEventDate;
            DateTime eEventDate;
            DateTime sTime;
            DateTime eTime;
            DateTime.TryParseExact(attendanceCode1StartTime, new string[] { "HHmm"}, CultureInfo.InvariantCulture, DateTimeStyles.None, out sTime);
            DateTime.TryParseExact(attendanceCode1EndTime, new string[] { "HHmm"}, CultureInfo.InvariantCulture, DateTimeStyles.None, out eTime);
            DateTime.TryParseExact(eventStartDate, new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out sEventDate);
            DateTime.TryParseExact(eventEndDate, new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out eEventDate);
            int dayNum = (today - sEventDate).Days + 1;

            if (DateTime.Compare(eEventDate, today) >= 0 && DateTime.Compare(sEventDate, today) <= 0 && TimeSpan.Compare(sTime.TimeOfDay, now.TimeOfDay) <= 0 && TimeSpan.Compare(eTime.TimeOfDay, now.TimeOfDay) >= 0 && dayNum.ToString().Equals(day))
            {
                return true;
            }
            return false;
        }

        private String FormatDate(String today)
        {
            String[] split = today.Split('/');
            today = split[1] + '/' + split[0] + '/' + split[2];
            return today;
        }
    }
}