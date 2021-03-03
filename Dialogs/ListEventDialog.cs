using Microsoft.Bot.Builder.Dialogs;
using System;
using System.Threading.Tasks;
using Microsoft.Bot.Connector;
using Microsoft.WindowsAzure.Storage;
using System.Configuration;
using Microsoft.WindowsAzure.Storage.Table;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Linq;
using global::AdaptiveCards;
using System.Threading;
using ProactiveBot;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class ListEventDialog : IDialog<object>
    {
        private SharedFunction sharedFunction = new SharedFunction();

        private IDictionary<int, string> eventDict = new Dictionary<int, string>();
        private List<string> questionList = new List<string>();
        private List<List<string>> answerList = new List<List<string>>(); 
        private List<string> eventList = new List<string>();
        private List<string> questionTypeList = new List<string>();
        private int cacheOption = -1;

        public async Task StartAsync(IDialogContext context)
        { 
            //Display a list of event available for display
            String name = context.UserData.GetValue<String>(ContextConstants.Name);
            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;
            await context.PostAsync(msg); 
            await context.PostAsync("Please enter the number of the event that you want to see details from.");
            //Print a list of events
            await context.PostAsync(displayEventList(context.Activity.From.Id, eventDict));

            context.Wait(this.MessageReceivedAsync);
        }

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> result)
        { 
            var message = await result;

            try
            { 
                EventEntity oneEventEntity = new EventEntity();
                bool foundEvent = false;
                int option = Convert.ToInt32(message.Text); 

                string oneEvent;
                //Check the option selected exists for the event
                if (eventDict.TryGetValue(option, out oneEvent))
                {
                    cacheOption = option;
                    oneEventEntity = JsonConvert.DeserializeObject<EventEntity>(oneEvent);
                    foundEvent = true; 
                }
                else //Option is not selected / cannot be found
                {
                    //however if dialog is run for second times, then it uses the cache option to gather event details
                    if (eventDict.TryGetValue(cacheOption, out oneEvent))
                    {
                        oneEventEntity = JsonConvert.DeserializeObject<EventEntity>(oneEvent);
                        foundEvent = true;
                    }

                    // Got an Action Submit (Dialog run for second time; action taken for adaptative card buttons)
                    dynamic value = message.Value;
                    string submitType = value.Type.ToString();
                    switch (submitType)
                    {
                        case "SaveEditSurvey": 
                            SaveEditSurvey(context, message.Value, oneEventEntity);

                            await context.PostAsync("Your event's survey has been successfully saved!"); 
                            context.Done(this);
                            return;

                        case "DeleteEvent":
                            await DeleteEventCardAsync(context);
                            return;

                        case "YesDeleteEvent":
                            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
                            var tableClient = storageAccount.CreateCloudTableClient();
                            CloudTable cloudTable = tableClient.GetTableReference("Event");
                            cloudTable.CreateIfNotExists();
                            cloudTable.Execute(TableOperation.Delete(oneEventEntity));

                            await context.PostAsync("Your event has been successfully deleted!");
                            context.Done(this);
                            return;

                        case "NoDeleteEvent":
                            await context.PostAsync("Okayy, talk to me again if you need assistance!");
                            context.Done(this);
                            return;
                    } 
                    await context.PostAsync("Your option is out of range!  Please enter again or type \"exit\" to exit the conversation.");
                }

                 //Event found
                if (foundEvent)
                {
                    var eventDetail = ""; 
                    eventDetail += "Event Name: " + oneEventEntity.EventName + "\n\n";
                    eventDetail += "Event Date: " + oneEventEntity.EventStartDate + "\n\n";
                    eventDetail += "Attendance Code 1: " + oneEventEntity.AttendanceCode1 + "\n\n";
                    eventDetail += "Attendance Code 2: " + oneEventEntity.AttendanceCode2 + "\n\n";
                    eventDetail += "Survey Code: " + oneEventEntity.RowKey + "\n\n"; 

                    //Adaptive Card starts from here
                    await ShowEventCardAsync(context, eventDetail, oneEventEntity);
                }
            }
            catch (Exception)
            {
                await context.PostAsync("It seems like your option is not an number. Please enter again or type \"exit\" to exit the conversation.");
            }
        }

        //Displays a list of event
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
            printMessage += "No.  Event Code :  Event Date  : Event Name  \n";
            foreach (EventEntity entity in eventTable.ExecuteQuery(query).OrderByDescending(o => o.Timestamp))
            {
                printMessage += count + ") " + entity.RowKey + "  : " + entity.EventStartDate + "   :  " + entity.EventName + " \n";
                eventDict.Add(count, JsonConvert.SerializeObject(entity));
                count++;
            }

            if (count == 1)
                return printMessage = "No event is found.";
            else
                return printMessage;
        }

        private async Task ShowEventCardAsync(IDialogContext context, string eventDetail, EventEntity oneEventEntity)
        {
            AdaptiveCard card = new AdaptiveCard()
            {
                Body = new List<CardElement>()
                {
                    new Container()
                    {
                        Speak = "<s>Hi, this is your selected Event Details</s>",
                        Items = new List<CardElement>()
                        {
                            new ColumnSet()
                            {
                                Columns = new List<Column>()
                                { 
                                    new Column()
                                    {
                                        Size = ColumnSize.Stretch,
                                        Items = new List<CardElement>()
                                        {
                                            new TextBlock()
                                            {
                                                Text =  "Your Selected Event Details",
                                                Weight = TextWeight.Bolder,
                                                IsSubtle = true
                                            },
                                            new TextBlock()
                                            {
                                                Text = eventDetail,
                                                Wrap = true
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                },
                // Buttons
                Actions = new List<ActionBase>() {
                   new SubmitAction()
                    {
                        Title = "Delete Event",
                        Speak = "<s>Delete Event</s>",
                        DataJson = "{ \"Type\": \"DeleteEvent\" }"
                    },
                    new ShowCardAction()
                    {
                        Title = "Edit Survey Questions",
                        Speak = "<s>Edit Survey Questions</s>", 
                        Card = GetEditSurveyCard(oneEventEntity)
                    }
                }
            };

            Attachment attachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card
            };

            var reply = context.MakeMessage();
            reply.Attachments.Add(attachment);

            await context.PostAsync(reply, CancellationToken.None); 
            context.Wait(MessageReceivedAsync);
        }

        private async Task DeleteEventCardAsync(IDialogContext context)
        {
            AdaptiveCard card = new AdaptiveCard()
            {
                Body = new List<CardElement>()
                {
                    new Container()
                    {
                        Speak = "<s>Hi, this is your selected Event Details</s>",
                        Items = new List<CardElement>()
                        {
                            new ColumnSet()
                            {
                                Columns = new List<Column>()
                                { 
                                    new Column()
                                    {
                                        Size = ColumnSize.Stretch,
                                        Items = new List<CardElement>()
                                        {
                                            new TextBlock()
                                            {
                                                Text =  "Delete Event",
                                                Weight = TextWeight.Bolder,
                                                IsSubtle = true
                                            },
                                            new TextBlock()
                                            {
                                                Text = "Are you sure you want to delete this event?",
                                                Wrap = true
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                },
                // Buttons
                Actions = new List<ActionBase>() {
                   new SubmitAction()
                    {
                        Title = "Yes",
                        Speak = "<s>Yes</s>",
                        DataJson = "{ \"Type\": \"YesDeleteEvent\" }"
                    },
                    new SubmitAction()
                    {
                        Title = "No",
                        Speak = "<s>No</s>",
                        DataJson = "{ \"Type\": \"NoDeleteEvent\" }"
                    }
                }
            };

            Attachment attachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card
            };

            var reply = context.MakeMessage();
            reply.Attachments.Add(attachment);

            await context.PostAsync(reply, CancellationToken.None); 
            context.Wait(MessageReceivedAsync);
        }

        private AdaptiveCard GetEditSurveyCard(EventEntity oneEventEntity)
        {
            StoreResultsToList(oneEventEntity, questionList, questionTypeList, answerList);
            AdaptiveCard card = new AdaptiveCard()
            {
                Body = new List<CardElement>()
                { 
                },
                Actions = new List<ActionBase>()
                {
                    new SubmitAction()
                    {
                        Title = "Save Survey",
                        Speak = "<s>Save Survey</s>",
                        DataJson = "{ \"Type\": \"SaveEditSurvey\" }"
                    }
                }
            };

            //Generate cards based on the questions
            for (int i = 0; i < questionList.Count; i++)
            {
                card.Body.Add(new TextInput()
                {
                    Id = "q" + i,
                    Placeholder = "Q" + (i + 1) + ". " + questionList[i],
                    Style = TextInputStyle.Text,
                    IsMultiline = true
                });

                if (questionTypeList[i] == "2") //Multiple-Choice
                {
                    card.Body.Add(new TextBlock()
                    {
                        Text = "Q" + (i+1) + ". Answer: " + string.Join(",", answerList[i].ToArray())
                    });

                }
                else    //OpenEnded Answer (1)
                {
                    card.Body.Add(new TextBlock()
                    {
                        Text = "Q" + (i+1) + ". Answer: Free-text"
                    }); 
                }
            } 
            return card;
        }

        private void StoreResultsToList(EventEntity eventEntity, List<string> questionList, List<string> questionTypeList, List<List<string>> answerList)
        {
            List<QuestionEntity> surveyEntity = new List<QuestionEntity>();

            surveyEntity = JsonConvert.DeserializeObject<List<QuestionEntity>>(eventEntity.Survey);
             
            foreach (QuestionEntity entity in surveyEntity)
            {
                questionList.Add(entity.QuestionText);
                if (JsonConvert.DeserializeObject<List<string>>(entity.AnswerList).Count != 0)
                {
                    answerList.Add(JsonConvert.DeserializeObject<List<string>>(entity.AnswerList));
                    questionTypeList.Add("2");
                }
                else
                {
                    answerList.Add(new List<string>());
                    questionTypeList.Add("1");
                } 
            }
        }

        private void SaveEditSurvey(IDialogContext context, dynamic resultValue, EventEntity oneEventEntity)
        { 
            // Get card from result
            Dictionary<string, string> editedQuestionList = resultValue.ToObject<Dictionary<string, string>>();

            for (int i = 0; i < questionList.Count; i++)
            {
                if (editedQuestionList.ContainsKey("q" + i))
                {
                    if (!(string.IsNullOrEmpty(editedQuestionList["q" + i]) || string.IsNullOrWhiteSpace(editedQuestionList["q" + i])))
                        questionList[i] = editedQuestionList["q" + i];
                }
            }

            //Serialize the object back and save it 
            var survey = new List<QuestionEntity>();
            var question = new QuestionEntity();
            for (int i = 0; i < questionList.Count; i++)
            { 
                question = new QuestionEntity();
                question.PartitionKey = oneEventEntity.SurveyCode;
                question.RowKey = (i + 1).ToString();
                question.AnswerList = JsonConvert.SerializeObject(answerList[i]);
                question.QuestionText = questionList[i];
                survey.Add(question);
            }
             
            oneEventEntity.Survey = JsonConvert.SerializeObject(survey);

            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            CloudTable cloudTable = tableClient.GetTableReference("Event");
            cloudTable.CreateIfNotExists();
            cloudTable.Execute(TableOperation.InsertOrMerge(oneEventEntity)); 
        }
    }
}