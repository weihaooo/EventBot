using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Sample.ProactiveBot;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json; 
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks; 

namespace ProactiveBot.Dialogs
{
    [Serializable]
    public class DeleteEventDialog : IDialog<object>
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

        private int selectedEvent = -1;
        private string selectedEventName = "event";
        private string selectedSurveyCode = "";

        public async Task StartAsync(IDialogContext context)
        {
            await context.PostAsync(displayCardReply(context)); 
            context.Wait(this.MessageReceivedAsync);
        }

        private IMessageActivity displayCardReply(IDialogContext context)
        {
            var reply = context.MakeMessage();
            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
            reply.Attachments = GetCardsAttachments(context.Activity.From.Id, eventDict, uniqueSurveyDict);

            if (reply.Attachments.Count > 0)
            {
                reply.Text = "Here are all your events! To exit, type 'exit'!";
            }
            else
            {
                reply.Text = "There is no event created yet."; 
            }  
            return reply;
        }

        public virtual async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> result)
        {
            var message = await result;

            try
            {
                if (!(string.IsNullOrEmpty(message.Text) || string.IsNullOrWhiteSpace(message.Text)))
                {
                    selectedEvent = Convert.ToInt32(message.Text.Remove(0, message.Text.Count() - 1));
                    if (eventDict.ContainsKey(selectedEvent))
                    {
                        EventEntity oneEventEntity = JsonConvert.DeserializeObject<EventEntity>(eventDict[selectedEvent]);
                        selectedEventName = oneEventEntity.EventName;
                        selectedSurveyCode = oneEventEntity.SurveyCode;

                        if (message.Text.Contains("DeleteEventPrompt"))
                        {
                            PromptDialog.Choice(context, this.ResumeAfterDeletePrompt,
                                new List<string>() { "Yes", "No" },
                                "Are you sure you want to delete " + selectedEventName + " (" + selectedSurveyCode + ")?",
                                "This is not an option, please choose 'Yes' or 'No'.. Are you sure you want to delete " + selectedEventName + " (" + selectedSurveyCode + ")?",
                                100, PromptStyle.Keyboard);
                        }
                        else if (message.Text.Contains("ViewResultPrompt"))
                        {
                            if (checkEventResponse(selectedEvent) > 0)
                            {
                                var newMessage = sharedResultFunction.exportResult(context, eventDict, oneEventEntity, responseList, questionList, answerList,
                                            questionTypeList, uniqueAnswerDict, attendanceDict);
                                sharedResultFunction.emailResult(context, eventDict, oneEventEntity, responseList, questionList, answerList,
                                    questionTypeList, uniqueAnswerDict, attendanceDict);
                                await context.PostAsync(newMessage);
                            }
                            else
                            {
                                await context.PostAsync("There is no response for this survey!");
                            }
                        }
                        else if (message.Text.Contains("ViewDetails"))
                        {
                            string msg = "";
                            if (uniqueSurveyDict.ContainsKey(oneEventEntity.SurveyCode))
                            {
                                msg = "No." + selectedEvent + ") " + oneEventEntity.EventName + " \n\n"
                                    + "Date: " + oneEventEntity.EventStartDate + " to " + oneEventEntity.EventEndDate + "  \n";


                                for (int i = 1; i <= uniqueSurveyDict[oneEventEntity.SurveyCode].Count; i++)
                                {
                                    foreach (var eventDay in uniqueSurveyDict[oneEventEntity.SurveyCode])
                                    {
                                        EventEntity oneEvent = JsonConvert.DeserializeObject<EventEntity>(eventDay);
                                        if (oneEvent.Day == i.ToString())
                                        {
                                            msg += "Day " + i + "  \n" 
                                                + "Attendance Code 1: " + (oneEvent.AttendanceCode1 == null ? "-" : oneEvent.AttendanceCode1) + "  \n"
                                                + "Attendance Code 2: " + (oneEvent.AttendanceCode2 == null ? "-" : oneEvent.AttendanceCode2) + "  \n"
                                                + "Survey Code: " + oneEvent.SurveyCode + "  \n.  \n";
                                        }
                                    }
                                }  
                            }
                            else
                            {
                                msg = "There is no event detail available.";
                            }
                            await context.PostAsync(msg);
                        }
                    }
                    else
                        await context.PostAsync("Something just went wrong. I think you better type 'exit' to restart the conversation or '/event' to view your events.");
                }
                else
                {
                    await context.PostAsync(displayCardReply(context));
                    context.Wait(this.MessageReceivedAsync);
                }
            }
            catch (Exception e)
            {
                await context.PostAsync("Something just went wrong. I think you better type 'exit' to restart the conversation or '/event' to view your events.");
            } 
        }

        private async Task ResumeAfterDeletePrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            if (message != null)
            {
                switch (message)
                {
                    case "Yes":
                        //Prompt for confirmation if there is response in the event
                        if (checkEventResponse(selectedEvent) > 0)
                        {
                            PromptDialog.Choice(context, this.ResumeAfterDeleteResponsePrompt,
                            new List<string>() { "Yes", "No" },
                            "Are you sure you want to delete " + selectedEventName + " (" + selectedSurveyCode + ") with responses?",
                            "This is not an option, please choose 'Yes' or 'No'.. Are you sure you want to delete " + selectedEventName + " (" + selectedSurveyCode + ") with responses?",
                            100, PromptStyle.Keyboard);
                        }
                        else
                        {
                            deleteEvent(selectedEvent);
                            await context.PostAsync("You have successfully deleted " + selectedEventName + " (" + selectedSurveyCode + ")!");
                            var reply = displayCardReply(context);
                            if (reply.Text.Equals("There is no event created yet."))
                            {
                                await context.PostAsync("You have no more events left. Talk to me again if you require my assistance.");
                                context.Done(this);
                            }
                            else
                            {
                                await context.PostAsync(reply);
                                context.Wait(this.MessageReceivedAsync);
                            }
                        }
                        break;

                    case "No":
                        await context.PostAsync(displayCardReply(context));
                        context.Wait(this.MessageReceivedAsync);
                        break; 
                }
            } 
        }

        private async Task ResumeAfterDeleteResponsePrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            if (message != null)
            {
                switch (message)
                {
                    case "Yes":
                        //Delete Workflow
                        deleteEvent(selectedEvent);
                        await context.PostAsync("You have successfully deleted " + selectedEventName + "!");
                        var reply = displayCardReply(context);
                        if (reply.Text.Equals("There is no event created yet."))
                        {
                            await context.PostAsync("You have no more events left. Talk to me again if you require my assistance.");
                            context.Done(this);
                        }
                        else
                        {
                            await context.PostAsync(reply);
                            context.Wait(this.MessageReceivedAsync);
                        }
                        break;

                    case "No":
                        await context.PostAsync(displayCardReply(context));
                        context.Wait(this.MessageReceivedAsync);
                        break; 
                }
            }
        } 

        private int checkEventResponse(int selectedEvent)
        {
            int  number = 0;
            
            if (eventDict.ContainsKey(selectedEvent))
            {
                EventEntity oneEventEntity = JsonConvert.DeserializeObject<EventEntity>(eventDict[selectedEvent]);
                sharedResultFunction.StoreResultsToList(oneEventEntity.SurveyCode, responseList, questionList, answerList, questionTypeList, uniqueAnswerDict, attendanceDict);

                number = responseList.Count();
            } 
            return number;
        }

        private void deleteEvent(int selectedEvent)
        {
            try
            {
                if (eventDict.ContainsKey(selectedEvent))
                {
                    EventEntity oneEventEntity = JsonConvert.DeserializeObject<EventEntity>(eventDict[selectedEvent]);

                    //Delete from event Table
                    var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
                    var tableClient = storageAccount.CreateCloudTableClient();
                    CloudTable eventTable = tableClient.GetTableReference("Event");
                    eventTable.CreateIfNotExists();

                    //Delete from Feedback Table
                    CloudTable feedbackTable = tableClient.GetTableReference("Feedback");
                    feedbackTable.CreateIfNotExists();

                    //Delete from Attendance
                    CloudTable attendanceTable = tableClient.GetTableReference("Attendance");
                    attendanceTable.CreateIfNotExists();

                    //Delete from Question Table
                    CloudTable questionTable = tableClient.GetTableReference("Question");
                    questionTable.CreateIfNotExists();

                    if (uniqueSurveyDict.ContainsKey(oneEventEntity.SurveyCode))
                    {
                        List<string> deleteEventList = uniqueSurveyDict[oneEventEntity.SurveyCode];

                        foreach (string e in deleteEventList)
                        {
                            EventEntity events = JsonConvert.DeserializeObject<EventEntity>(e);

                            TableQuery<FeedbackEntity> feedbackQuery = new TableQuery<FeedbackEntity>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, events.SurveyCode));
                            foreach (FeedbackEntity entity in feedbackTable.ExecuteQuery(feedbackQuery))
                            {
                                feedbackTable.Execute(TableOperation.Delete(entity));
                            }

                            TableQuery<FeedbackEntity> attendanceQuery = new TableQuery<FeedbackEntity>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, events.RowKey));
                            foreach (FeedbackEntity entity in attendanceTable.ExecuteQuery(attendanceQuery))
                            {
                                attendanceTable.Execute(TableOperation.Delete(entity));
                            }

                            TableQuery<FeedbackEntity> questionQuery = new TableQuery<FeedbackEntity>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, events.SurveyCode));
                            foreach (FeedbackEntity entity in questionTable.ExecuteQuery(questionQuery))
                            {
                                questionTable.Execute(TableOperation.Delete(entity));
                            }

                            eventTable.Execute(TableOperation.Delete(events));
                        }
                    }
                }
            }
            catch (Exception)
            {
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
                        Title = "No." + count + " " + entity.EventName + " (" + entity.SurveyCode + ")",
                        Subtitle = "Date: " + entity.EventStartDate + " to " + entity.EventEndDate,
                        Text = "Attendance Code 1: " + entity.AttendanceCode1 + "  \n"
                                + "Attendance Code 2: " + entity.AttendanceCode2 + "  \n"
                                + "Survey Code: " + entity.SurveyCode + "  \n"
                                + entity.Description,
                        Buttons = new List<CardAction>
                        {
                            new CardAction()
                            {
                                Title = "View Details",
                                Type = ActionTypes.PostBack,
                                Value = "ViewDetails" + count
                            },
                            new CardAction()
                            {
                                Title = "View Result",
                                Type = ActionTypes.PostBack,
                                Value = "ViewResultPrompt" + count
                            },
                            new CardAction()
                            {
                                Title = "Delete Event",
                                Type = ActionTypes.PostBack,
                                Value = "DeleteEventPrompt" + count
                            }
                        }
                    };

                    eventCardList.Add(heroCard.ToAttachment());
                    eventDict.Add(count, JsonConvert.SerializeObject(entity));
                    uniqueSurveyDict.Add(entity.SurveyCode, new List<string>() { JsonConvert.SerializeObject(entity) });
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