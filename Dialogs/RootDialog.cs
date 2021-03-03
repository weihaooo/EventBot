using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class RootDialog : IDialog<object>
    {
        bool anonymous = false;
        private static readonly HttpClient client = new HttpClient();

        public Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);

            return Task.CompletedTask;
        }

        private async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> result)
        {
            var activity = await result;
            string name = "";
            string email = "";
            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;
            await context.PostAsync(msg);
            
            context.UserData.SetValue(ContextConstants.UserId, activity.From.Id);
            var authorized = true;
            try
            {
                if (context.Activity.ChannelId == "facebook")
                {
                    var url = "https://graph.facebook.com/v2.12/" + activity.From.Id + "?fields=name,email&access_token="+ ConfigurationManager.AppSettings["WorkChatAccessToken"];
                    var responseString = await client.GetStringAsync(url);
                    dynamic json = JsonConvert.DeserializeObject(responseString);
                    name = json.name;
                    email = json.email;
                    context.UserData.SetValue(ContextConstants.Name, name);

                    if (email != null)
                    {
                        string[] split = email.Split('@');

                        if (!split[1].Equals("jtc.gov.sg") && !split[1].Equals("test.jtc.gov.sg"))
                        {
                            await context.PostAsync("I'm sorry but it appears that you do not have access to my services. Goodbye.");
                            authorized = false;
                            context.Done(this);
                        }
                    }
                    else
                    {
                        await context.PostAsync("I'm sorry but it appears that you do not have access to my services. Goodbye.");
                        authorized = false;
                        context.Done(this);
                    }
                }
            } catch(Exception e)
            {
            }

            if (authorized)
            {
                if (activity.From != null && name == "")
                {
                    context.UserData.SetValue(ContextConstants.Name, activity.From.Name);
                }

                if (!context.UserData.TryGetValue(ContextConstants.Name, out name))
                {
                    name = "sir/madam";
                    PromptDialog.Text(context, this.ResumeAfterPrompt, "Hello " + name + ", may I know your name ? (Please try to key in your full name in order for us to uniquely identify you.) ");
                }
                else
                {
                    context.Call(new CodeDialog(), this.CodeDialogResumeAfter);
                }
            }
        }

        private async Task ResumeAfterPrompt(IDialogContext context, IAwaitable<string> result)
        {
            try
            {
                var name = await result;
                var msg = context.MakeMessage();
                msg.Type = ActivityTypes.Typing;
                await context.PostAsync(msg);
                await context.PostAsync($"Welcome {name}!");
                context.UserData.SetValue(ContextConstants.Name, name);

                context.Call(new CodeDialog(), this.CodeDialogResumeAfter);
            }
            catch (TooManyAttemptsException)
            {
            }
        }

        private async Task CodeDialogResumeAfter(IDialogContext context, IAwaitable<object> result)
        {
            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;
            var eventName = context.UserData.GetValue<String>(ContextConstants.EventName);
            try
            {
                if (context.UserData.GetValue<string>(ContextConstants.Status) == "1")
                {
                    await context.PostAsync(msg);
                    PromptDialog.Choice(context, this.ResumeAfterCodePrompt, 
                        new List<string>() { "Yes" , "No" },
                        "Are you here to register your attendance for " + "'" + eventName + "'" + " today?",
                        "Are you here to register your attendance for " + "'" + eventName + "'" + " today? Please choose either 'Yes' or 'No'",
                        100, PromptStyle.Keyboard);
                } else if (context.UserData.GetValue<string>(ContextConstants.Status) == "2")
                {
                    await context.PostAsync(msg);
                    PromptDialog.Choice(context, this.ResumeAfterCodePrompt,
                        new List<string>() { "Yes", "No" },
                        "Are you here to register your attendance for " + "'" + eventName + "'" + " today?",
                        "Are you here to register your attendance for " + "'" + eventName + "'" + " today? Please choose either 'Yes' or 'No'",
                        100, PromptStyle.Keyboard);
                 } else if (context.UserData.GetValue<string>(ContextConstants.Status) == "3")
                {
                    await context.PostAsync(msg);
                    PromptDialog.Choice(context, this.ResumeAfterCodePrompt,
                        new List<string>() { "Yes", "No" },
                        "Are you here to proceed with the survey for " + "'" + eventName + "'" + " today?",
                        "Are you here to proceed with the survey for " + "'" + eventName + "'" + " today? Please choose either 'Yes' or 'No'",
                        100, PromptStyle.Keyboard);

                }
            }
            catch (TooManyAttemptsException)
            {
                await context.PostAsync(msg);
                await context.PostAsync("I'm sorry, I'm having issues understanding you. Let's try again.");
                await this.StartAsync(context);
            }
        }

        private async Task ResumeAfterCodePrompt(IDialogContext context, IAwaitable<object> result)
        {
            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;

            try
            {
                var choice = await result;
                if (choice.ToString() == "No")
                {
                    await context.PostAsync(msg);
                    await context.PostAsync("I see. Hope you have a pleasant day ahead!");
                    await context.PostAsync(msg);
                    await context.PostAsync("Talk to me again if you require my assistance.");
                    context.Done(this);
                }
                else if(choice.ToString() == "Yes")
                {
                    if (context.UserData.GetValue<string>(ContextConstants.Status) == "1")
                    {
                        RegisterAttendance(context.UserData.GetValue<string>(ContextConstants.SurveyCode), context.UserData.GetValue<string>(ContextConstants.EventCode), context.UserData.GetValue<string>(ContextConstants.UserId), context.UserData.GetValue<string>(ContextConstants.Today), context.UserData.GetValue<string>(ContextConstants.EventName), context.UserData.GetValue<string>(ContextConstants.Status), context.UserData.GetValue<string>(ContextConstants.Name));
                        await context.PostAsync(msg);
                        await context.PostAsync("Thank you " + context.UserData.GetValue<string>(ContextConstants.Name) + ", you have successfully registered your attendance for " + context.UserData.GetValue<string>(ContextConstants.EventName) + "!");
                        await context.PostAsync(msg);
                        await context.PostAsync("Talk to me again if you require my assistance.");
                        context.Done(this);
                    }
                    else if (context.UserData.GetValue<string>(ContextConstants.Status) == "2")
                    {
                        RegisterAttendance(context.UserData.GetValue<string>(ContextConstants.SurveyCode), context.UserData.GetValue<string>(ContextConstants.EventCode), context.UserData.GetValue<string>(ContextConstants.UserId), context.UserData.GetValue<string>(ContextConstants.Today), context.UserData.GetValue<string>(ContextConstants.EventName), context.UserData.GetValue<string>(ContextConstants.Status), context.UserData.GetValue<string>(ContextConstants.Name));
                        await context.PostAsync(msg);
                        await context.PostAsync("Thank you " + context.UserData.GetValue<string>(ContextConstants.Name) + ", you have successfully registered your attendance for " + context.UserData.GetValue<string>(ContextConstants.EventName) + "!");
                        await context.PostAsync(msg);
                        await context.PostAsync("Talk to me again if you require my assistance.");
                        context.Done(this);
                    }
                    else if (context.UserData.GetValue<string>(ContextConstants.Status) == "3")
                    {
                        string description;
                        if (context.UserData.TryGetValue(ContextConstants.Description, out description))
                        {
                            await context.PostAsync(context.UserData.GetValue<string>(ContextConstants.Description)); 
                        }

                        if (context.UserData.GetValue<string>(ContextConstants.Anonymous) == "Y" || context.UserData.GetValue<string>(ContextConstants.Anonymous) == "Yes")
                        {
                            PromptDialog.Confirm(context, this.ResumeAfterAnonymousPrompt,
                                "Would you like to take this survey as Anonymous?",
                                "Would you like to take this survey as Anonymous?" + "\n\nPlease choose one of the options.", 100, PromptStyle.Keyboard);
                        }
                        else
                        {
                            if (context.Activity.ChannelId == "webchat" || context.Activity.ChannelId == "emulator")
                            {
                                PromptDialog.Choice(context, this.ResumeAfterSurveyTypePrompt, new List<string>() { "Chatbot", "Form" },
                                    "How do you want to do your survey? Chatbot or Form?",
                                    "How do you want to do your survey? Chatbot or Form?",
                                    100, PromptStyle.Keyboard);
                            }
                            else
                            {
                                await context.Forward(new SurveyDialog(), this.SurveyDialogResumeAfter, "Chatbot", CancellationToken.None);
                            }
                        } 
                    }
                }
            }
            catch (TooManyAttemptsException)
            {
            }
        }

        private async Task ResumeAfterAnonymousPrompt(IDialogContext context, IAwaitable<bool> result)
        {
            try
            {
                bool response = await result;

                if (response)
                {
                    anonymous = true;
                    context.UserData.SetValue<string>(ContextConstants.AnonymousStatus, "anonymous");
                }
                else
                {
                    context.UserData.SetValue<string>(ContextConstants.AnonymousStatus, context.UserData.GetValue<String>(ContextConstants.Name));
                }
                
                if (context.Activity.ChannelId == "webchat" || context.Activity.ChannelId == "emulator")
                {
                    PromptDialog.Choice(context, this.ResumeAfterSurveyTypePrompt, new List<string>() { "Chatbot", "Form" },
                        "How do you want to do your survey? Chatbot or Form?",
                        "How do you want to do your survey? Chatbot or Form?",
                        100, PromptStyle.Keyboard);
                }
                else
                {
                    await context.Forward(new SurveyDialog(), this.SurveyDialogResumeAfter, "Chatbot", CancellationToken.None);
                } 
            }
            catch (Exception e)
            {
                await context.PostAsync(e.ToString());
            }
        }

        private async Task ResumeAfterSurveyTypePrompt(IDialogContext context, IAwaitable<string> result)
        {
            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;

            var choice = await result;
            await context.Forward(new SurveyDialog(), this.SurveyDialogResumeAfter, choice, CancellationToken.None);
        }


        private void RegisterAttendance(String surveyCode, string eventCode, string userid, string today, string eventName, string status, string name)
        {
            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            CloudTable cloudTable = tableClient.GetTableReference("Attendance");
            cloudTable.CreateIfNotExists();

            String filterA = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, eventCode);
            String filterD = TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, userid);
            TableQuery<AttendanceEntity> query = new TableQuery<AttendanceEntity>().Where(TableQuery.CombineFilters(filterA, TableOperators.And, filterD));

            var results = cloudTable.ExecuteQuery(query);

            AttendanceEntity attendance = new AttendanceEntity(eventCode, userid);
            attendance.SurveyCode = surveyCode;
            attendance.Date = today;
            attendance.EventName = eventName;
            attendance.Name = name;
            foreach(AttendanceEntity a in results)
            {
                attendance.Morning = a.Morning;
                attendance.Afternoon = a.Afternoon;
                attendance.Survey = a.Survey;
            }
            
            if(status == "1")
            {
                attendance.Morning = true;
            } else if (status == "2")
            {
                attendance.Afternoon = true;
            } else if (status == "3")
            {
                attendance.Survey = true;
            }
            
            TableOperation insertOperation = TableOperation.InsertOrMerge(attendance);
            cloudTable.Execute(insertOperation);
        }

        private async Task SurveyDialogResumeAfter(IDialogContext context, IAwaitable<object> result)
        {
            try
            {
                RegisterAttendance(context.UserData.GetValue<string>(ContextConstants.SurveyCode), context.UserData.GetValue<string>(ContextConstants.EventCode), context.UserData.GetValue<string>(ContextConstants.UserId), context.UserData.GetValue<string>(ContextConstants.Today), context.UserData.GetValue<string>(ContextConstants.EventName), context.UserData.GetValue<string>(ContextConstants.Status), context.UserData.GetValue<string>(ContextConstants.Name));

            }
            catch (TooManyAttemptsException)
            {
                await context.PostAsync("I'm sorry, I'm having issues understanding you. Let's try again.");
                await this.StartAsync(context);
            }
        }
    }
}