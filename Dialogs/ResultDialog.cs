using Microsoft.Bot.Builder.Dialogs;
using System;
using System.Threading.Tasks;
using Microsoft.Bot.Connector;
using Microsoft.WindowsAzure.Storage;
using System.Configuration;
using Microsoft.WindowsAzure.Storage.Table;
using System.Collections.Generic; 
using Newtonsoft.Json; 
using System.Net.Http; 
using ProactiveBot;
using System.Linq; 

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class ResultDialog : IDialog<object>
    {
        private SharedResultFunction sharedResultFunction = new SharedResultFunction();

        private IDictionary<int, string> eventDict = new Dictionary<int, string>();
        private List<string> questionList = new List<string>();
        private List<List<string>> answerList = new List<List<string>>();
        private IDictionary<string, string> uniqueAnswerDict = new Dictionary<string, string>();
        private List<string> responseList = new List<string>();
        private List<string> questionTypeList = new List<string>();
        private IDictionary<string, int> attendanceDict = new Dictionary<string, int>();
        private IDictionary<string, List<string>> uniqueSurveyDict = new Dictionary<string, List<string>>();
        
        private string pwdEventList;  
        private bool sharedAccess = false;
        private string code;

        private int selectedEvent = -1;

        public ResultDialog(string message)
        {
            code = message;
        }

        public async Task StartAsync(IDialogContext context)
        {
            if (code != null)
            { 
                var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
                var tableClient = storageAccount.CreateCloudTableClient();
                CloudTable eventTable = tableClient.GetTableReference("Event");
                TableQuery<EventEntity> query = new TableQuery<EventEntity>().Where(TableQuery.GenerateFilterCondition("SurveyCode", QueryComparisons.Equal, code.Substring(0, 8)));
                List<EventEntity> eventList = eventTable.ExecuteQuery(query).ToList();
                if (eventList.Count == 0)
                {
                    await context.PostAsync("Invalid Survey Code. Talk to me again if you require my assistance");
                    context.Done(this);
                }
                else
                { 
                    pwdEventList = JsonConvert.SerializeObject(eventList);

                    if (context.Activity.From.Id == eventList.FirstOrDefault().PartitionKey)
                    {
                        sharedAccess = true;
                        await generateResult(context);
                        context.Done(this);
                    }
                    else
                    { 
                        PromptDialog.Text(context, ResumeAfterPrompt, "Please enter the Password."); 
                    }
                }
            }
            else
            {
                var reply = context.MakeMessage();
                reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                reply.Attachments = GetCardsAttachments(context.Activity.From.Id, eventDict, uniqueSurveyDict);

                if(reply.Attachments.Count == 0)
                {
                    await context.PostAsync("You do not have any events!");
                } else
                    await context.PostAsync(reply);

                context.Wait(this.MessageReceivedAsync);
            }
        }

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> result)
        {
            var message = await result;


            if (!(string.IsNullOrEmpty(message.Text) || string.IsNullOrWhiteSpace(message.Text)))
            {
                if (message.Text.Contains("ViewResultPrompt"))
                {
                    selectedEvent = Convert.ToInt32(message.Text.Remove(0, message.Text.Count() - 1));

                    if (checkEventResponse(selectedEvent) > 0)
                    {
                        if (eventDict.ContainsKey(selectedEvent))
                        {
                            EventEntity oneEventEntity = JsonConvert.DeserializeObject<EventEntity>(eventDict[selectedEvent]);

                            var newMessage = sharedResultFunction.exportResult(context, eventDict, oneEventEntity, responseList, questionList, answerList,
                                            questionTypeList, uniqueAnswerDict, attendanceDict);
                            sharedResultFunction.emailResult(context, eventDict, oneEventEntity, responseList, questionList, answerList,
                                    questionTypeList, uniqueAnswerDict, attendanceDict);
                            await context.PostAsync(newMessage);
                            await context.PostAsync("Talk to me again if you require my assistance.");
                            context.Done(this);
                        }
                    }
                    else
                    {
                        await context.PostAsync("There is no responses for this survey!");
                        context.Done(this);
                    }
                }
            }
            else
            {
                var reply = context.MakeMessage();

                reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
                reply.Attachments = GetCardsAttachments(context.Activity.From.Id, eventDict, uniqueSurveyDict);
                await context.PostAsync(reply);
                context.Wait(this.MessageReceivedAsync);
            }
        }

        private async Task generateResult(IDialogContext context, string message = null)
        { 
            try
            {
                bool foundEvent = false;
                EventEntity oneEventEntity = new EventEntity();

                if (!sharedAccess && code == null)
                {
                    int option = Convert.ToInt32(message);

                    string oneEvent;
                    if (eventDict.TryGetValue(option, out oneEvent))
                    {
                        oneEventEntity = JsonConvert.DeserializeObject<EventEntity>(oneEvent);
                        foundEvent = true;
                    }
                    else
                    {
                        await context.PostAsync("Your option is out of range! Please enter an option again.");
                    }
                }
                else
                {
                    oneEventEntity = JsonConvert.DeserializeObject<List<EventEntity>>(pwdEventList).FirstOrDefault();
                }

                //Event is found (option is valid)
                if (foundEvent || sharedAccess)
                {
                    sharedResultFunction.StoreResultsToList(oneEventEntity.SurveyCode, responseList, questionList, answerList, questionTypeList, uniqueAnswerDict, attendanceDict);

                    var newMessage = sharedResultFunction.exportResult(context, eventDict, oneEventEntity, responseList, questionList, answerList,
                                    questionTypeList, uniqueAnswerDict, attendanceDict);
                    if ((!sharedAccess && code == null) || (sharedAccess && oneEventEntity.PartitionKey == context.Activity.From.Id))
                    {
                        sharedResultFunction.emailResult(context, eventDict, oneEventEntity, responseList, questionList, answerList,
                                    questionTypeList, uniqueAnswerDict, attendanceDict);
                    }
                    await context.PostAsync(newMessage);
                    await context.PostAsync("Talk to me again if you require my assistance.");
                    if (foundEvent)
                    {
                        context.Done(this);
                    }
                }
            }
            catch (FormatException)
            {
                await context.PostAsync("It seems like your option is not an number. Please try again.");
            } 
        }

        private int checkEventResponse(int selectedEvent)
        {
            int number = 0;

            if (eventDict.ContainsKey(selectedEvent))
            {
                EventEntity oneEventEntity = JsonConvert.DeserializeObject<EventEntity>(eventDict[selectedEvent]);
                sharedResultFunction.StoreResultsToList(oneEventEntity.SurveyCode, responseList, questionList, answerList, questionTypeList, uniqueAnswerDict, attendanceDict);

                number = responseList.Count();
            }
            return number;
        }

        private string displayEventList(string uniqueID, IDictionary<int, string> eventDict)
        {
            string printMessage = "";
            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            CloudTable eventTable = tableClient.GetTableReference("Event");
            string searchPartition = uniqueID;  //UserID of the user

            TableQuery<EventEntity> query = new TableQuery<EventEntity>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, searchPartition));

            int count = 1;
            //Displays all the event name and code 

            IDictionary<string, int> uniqueSurveyDict = new Dictionary<string, int>();
            printMessage += "No. Survey Code :  Event Duration  :  Event Name  \n";
            foreach (EventEntity entity in eventTable.ExecuteQuery(query).OrderByDescending(o => o.Timestamp))
            {
                if (!uniqueSurveyDict.ContainsKey(entity.SurveyCode))
                {
                    printMessage += count + ". " + entity.SurveyCode + " : " + entity.EventStartDate + "-" + entity.EventEndDate + " : " + entity.EventName + "  \n";
                    eventDict.Add(count, JsonConvert.SerializeObject(entity));
                    uniqueSurveyDict.Add(entity.SurveyCode, count); 
                    count++;
                }
            }

            if (count == 1)
                return printMessage = "No event is found..";
            else
                return printMessage;
        }
         
        private async Task ResumeAfterPrompt(IDialogContext context, IAwaitable<string> result)
        {
            try
            {
                var pw = await result; 

                if (!(string.IsNullOrEmpty(pw) || string.IsNullOrWhiteSpace(pw)))
                {
                    string pwd = JsonConvert.DeserializeObject<List<EventEntity>>(pwdEventList).FirstOrDefault().Password;
                    
                    if (pwd == sharedResultFunction.Sha256(pw))
                    {
                        sharedAccess = true;
                        await generateResult(context);
                        context.Done(this);
                    }
                    else
                    {
                        await context.PostAsync("The password is invalid!");
                        await context.PostAsync("Talk to me again if you require my assistance.");
                        context.Done(this);
                    }
                }
                else
                {
                    await context.PostAsync("The password is invalid!");
                    await context.PostAsync("Talk to me again if you require my assistance.");
                    context.Done(this);
                }
            }
            catch (TooManyAttemptsException)
            {
                await context.PostAsync("The password is invalid!");
                await context.PostAsync("Talk to me again if you require my assistance.");
                context.Done(this);
            }
        }

        private static IList<Attachment> GetCardsAttachments(string uniqueID, IDictionary<int, string> eventDict, IDictionary<string, List<string>> uniqueSurveyDict)
        {
            List<Attachment> eventCardList = new List<Attachment>();

            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            CloudTable eventTable = tableClient.GetTableReference("Event");
            string searchPartition = uniqueID;  //UserID of the user

            TableQuery<EventEntity> query = new TableQuery<EventEntity>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, searchPartition));

            uniqueSurveyDict.Clear();
            eventDict.Clear();

            int count = 1;
            foreach (EventEntity entity in eventTable.ExecuteQuery(query).OrderByDescending(o => o.Timestamp))
            {
                if (!uniqueSurveyDict.ContainsKey(entity.SurveyCode))
                {
                    var heroCard = new HeroCard
                    {
                        Title = "No " + count + ". " + entity.EventName + " (" + entity.SurveyCode + ")",
                        Subtitle = "Date: " + entity.EventStartDate + " to " + entity.EventEndDate,
                        Text = "Attendance Code 1: " + entity.AttendanceCode1 + "  \n"
                                + "Attendance Code 2: " + entity.AttendanceCode2 + "  \n"
                                + "Survey Code: " + entity.SurveyCode + "  \n"
                                + entity.Description,
                        Buttons = new List<CardAction>
                        {
                            new CardAction()
                            {
                                Title = "View Result",
                                Type = ActionTypes.PostBack,
                                Value = "ViewResultPrompt" + count
                            }
                        }
                    };

                    eventCardList.Add(heroCard.ToAttachment());
                    eventDict.Add(count, JsonConvert.SerializeObject(entity));
                    uniqueSurveyDict.Add(entity.SurveyCode, new List<string>() { JsonConvert.SerializeObject(entity)  });
                    count++;
                }
                else
                {
                    uniqueSurveyDict[entity.SurveyCode].Add(JsonConvert.SerializeObject(entity)); 
                }
            }
            return eventCardList;
        }
    } 
}