using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading;
using System.Threading.Tasks;
using AdaptiveCards;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class SurveyDialog : IDialog<object>
    {
        List<string> question = new List<string>();
        List<List<string>> answerList = new List<List<string>>();
        List<string> questionType = new List<string>();
        List<string> responseList = new List<string>();
        int count = 0;
        int max = 0;

        public async Task StartAsync(IDialogContext context)
        { 
            List<QuestionEntity> survey = JsonConvert.DeserializeObject<List<QuestionEntity>>(context.UserData.GetValue<string>(ContextConstants.Survey));
            max = survey.Count;
            foreach(QuestionEntity q in survey)
            {
                question.Add(q.QuestionText);
                if (JsonConvert.DeserializeObject<List<string>>(q.AnswerList).Count!=0)
                {
                    answerList.Add(JsonConvert.DeserializeObject<List<string>>(q.AnswerList));
                    questionType.Add("2");
                } else
                {
                    answerList.Add(new List<string>());
                    questionType.Add("1");
                }
            }
            
            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;
            await context.PostAsync(msg); 
            await context.PostAsync("There are a total of " + survey.Count + " questions in the survey, shouldn't take more than "+ survey.Count*0.5+" minutes of your time!"); // Count the questions in future
            await context.PostAsync(msg);
            
            if (context.Activity.ChannelId == "webchat" || context.Activity.ChannelId == "emulator")
            {
                context.Wait(this.MessageReceivedAsync);
            }
            else
            {
                context.Wait(this.MessageReceivedAsync);
            } 
        }

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<object> result)
        { 
            var postResult = (Activity)context.Activity; 
            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;
            await context.PostAsync(msg);

            await surveyChoice(context, postResult.Text, postResult.Value);
        }

        private async Task surveyChoice(IDialogContext context, string postResult, dynamic postResultValue = null)
        {
            if (postResult == "Form")    //Form mode for survey
            {
                //Adaptive Card starts from here
                await SurveyQuestionsCard(context);
            }
            else
            {
                if (questionType[count] == "1")
                {
                    PromptDialog.Text(context, this.ResumeAfterPrompt, "Q"+(count+1)+ ". " +question[count]);
                }
                else if (questionType[count] == "2")
                {
                    PromptDialog.Choice(context, this.ResumeAfterPrompt, answerList[count], "Q" + (count + 1) + ". " + question[count], "Q" + (count + 1) + ". " + question[count] + "\n\nPlease choose one of the options.", 100, PromptStyle.Keyboard);
                } 
            }
            
            if (postResultValue != null)
            {
                // Got an Action Submit (Dialog run for second time; action taken for adaptative card buttons)
                dynamic value = postResultValue;
                string submitType = value.Type.ToString();
                switch (submitType)
                {
                    case "SubmitSurvey":
                        SubmitSurvey(context, value);
                        await context.PostAsync("Your survey has been successfully submitted!");
                        context.Done(this);
                        return;
                }
            }
        }

        private async Task SurveyQuestionsCard(IDialogContext context)
        {
            AdaptiveCard card = new AdaptiveCard()
            {
                Body = new List<CardElement>()
                {
                    new Container()
                    { 
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
                                                Text =  context.UserData.GetValue<string>(ContextConstants.Today) + ", " + context.UserData.GetValue<string>(ContextConstants.EventName),
                                                Weight = TextWeight.Bolder,
                                                IsSubtle = true, 
                                                Wrap = true
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                },
                Actions = new List<ActionBase>()
                {
                    new SubmitAction()
                    {
                        Title = "Submit Survey",
                        Speak = "<s>Submit Survey</s>",
                        DataJson = "{ \"Type\": \"SubmitSurvey\" }"
                    }
                }
            };

            //Generate cards based on the questions
            for (int i = 0; i < question.Count; i++)
            {
                card.Body.Add(new TextBlock()
                {
                    Text = "Q" + (i + 1) + ". " + question[i]
                });

                if (questionType[i] == "2") //Multiple-Choice
                {
                    var choices = new List<Choice>();

                    for (int j = 0; j < answerList[i].Count; j++)
                    {
                        choices.Add(new Choice()
                        {
                            Title = answerList[i][j],
                            Value = answerList[i][j]
                        });
                    }
                     
                    var choiceSet = new ChoiceSet()
                    {
                        IsMultiSelect = false,
                        Choices = choices,
                        Style = ChoiceInputStyle.Compact,
                        Id = "ans" + i
                    };
                    card.Body.Add(choiceSet);  
                }
                else    //OpenEnded Answer (1)
                {
                    card.Body.Add(new TextInput()
                    {
                        Id = "ans" + i,
                        Placeholder = "Please enter a comment",
                        Style = TextInputStyle.Text,
                        IsMultiline = true
                    });
                }
            } 

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

        private async Task ResumeAfterPrompt(IDialogContext context, IAwaitable<string> result)
        {
            try
            {
                var response = await result;
                responseList.Add(response);

                count++;
                var msg = context.MakeMessage();
                msg.Type = ActivityTypes.Typing;
                if (count >= max)
                {
                    bool success = false;
                    string anonymous;
                    if (context.UserData.TryGetValue(ContextConstants.AnonymousStatus, out anonymous))
                    {
                        success = SaveFeedback(context.UserData.GetValue<string>(ContextConstants.SurveyCode), context.UserData.GetValue<string>(ContextConstants.UserId), context.UserData.GetValue<string>(ContextConstants.Today), context.UserData.GetValue<string>(ContextConstants.EventName), anonymous, context.UserData.GetValue<string>(ContextConstants.Survey));
                    }
                    else
                    {
                        success = SaveFeedback(context.UserData.GetValue<string>(ContextConstants.SurveyCode), context.UserData.GetValue<string>(ContextConstants.UserId), context.UserData.GetValue<string>(ContextConstants.Today), context.UserData.GetValue<string>(ContextConstants.EventName), context.UserData.GetValue<string>(ContextConstants.Name), context.UserData.GetValue<string>(ContextConstants.Survey));
                    }

                    if (success)
                    {
                        await context.PostAsync(msg);
                        await context.PostAsync("This is the end of the evaluation survey form.");
                        await context.PostAsync(msg);
                        await context.PostAsync("Thank you for your feedback and time. Have a great day ahead!");
                        await context.PostAsync(msg);
                        await context.PostAsync("Talk to me again if you require my assistance."); 
                    }
                    else
                    {
                        await context.PostAsync(msg);
                        await context.PostAsync("Something went wrong, try again or contact my boss!");
                        await context.PostAsync(msg);
                        await context.PostAsync("Talk to me again if you require my assistance.");
                    }
                    context.Done(this);
                }
                else
                {
                    await this.MessageReceivedAsync(context, null);
                }
            }
            catch (TooManyAttemptsException)
            {
            }
        } 

        private bool SaveFeedback(string code, string userid, string today, string workshopName, string name, string survey)
        {
            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            CloudTable cloudTable = tableClient.GetTableReference("Feedback");
            cloudTable.CreateIfNotExists();

            TableOperation retrieveOperation = TableOperation.Retrieve<FeedbackEntity>(code, userid);
            TableResult retrievedResult = cloudTable.Execute(retrieveOperation);
            
            if (retrievedResult.Result == null)
            {
                FeedbackEntity feedback = new FeedbackEntity(code, userid);
                feedback.Date = today;
                feedback.WorkshopName = workshopName;
                feedback.Name = name;
                feedback.Survey = survey;
                feedback.Response = JsonConvert.SerializeObject(responseList);
                
                TableOperation insertOperation = TableOperation.Insert(feedback);
                cloudTable.Execute(insertOperation);

                return true;
            }
            return false;
        }

        private void SubmitSurvey(IDialogContext context, dynamic resultValue)
        {
            // Get card from result
            Dictionary<string, string> selectedAnswerList = resultValue.ToObject<Dictionary<string, string>>();

            for (int i = 0; i < answerList.Count; i++)
            {
                if (selectedAnswerList.ContainsKey("ans" + i))
                { 
                    responseList.Add(selectedAnswerList["ans" + i]);
                }
            } 
            bool success = SaveFeedback(context.UserData.GetValue<string>(ContextConstants.SurveyCode), context.UserData.GetValue<string>(ContextConstants.UserId), context.UserData.GetValue<string>(ContextConstants.Today), context.UserData.GetValue<string>(ContextConstants.EventName), context.UserData.GetValue<string>(ContextConstants.Name), context.UserData.GetValue<string>(ContextConstants.Survey));
        }
    }
}