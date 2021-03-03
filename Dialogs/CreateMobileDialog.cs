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
using ProactiveBot;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class CreateMobileDialog : IDialog<object>
    {
        List<string> eventDetails = new List<string>();
        DateTime sDate;
        DateTime eDate;
        string startTime;
        string endTime;
        int totalDays;
        int qCount = 0;
        List<string> code1 = new List<string>();
        List<string> code2 = new List<string>();
        List<string> surveyList = new List<string>();
        List<int> optionList = new List<int>();

        public async Task StartAsync(IDialogContext context)
        {
            eventDetails = new List<string>();
            var msg = context.MakeMessage();
            msg.Type = ActivityTypes.Typing;
            await context.PostAsync(msg);
            
            PromptDialog.Text(context, this.ResumeAfterPrompt, "What is the Event Name?");
        }

        private async Task ResumeAfterPrompt(IDialogContext context, IAwaitable<string> result)
        {
            context.UserData.SetValue(ContextConstants.UserId, context.Activity.From.Id);

            var eventCode = GenerateCode();
            eventDetails.Add(context.UserData.GetValue<string>(ContextConstants.UserId));//0 - Partition Key
            eventDetails.Add(eventCode);//1 - Row Key

            var message = await result;

            eventDetails.Add(message);//2 - Event Name
            PromptDialog.Text(context, this.ResumeAfterNamePrompt, "When is the Event Start Date(dd/mm/yyyy)?");
        }

        private async Task ResumeAfterNamePrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            DateTime thisDay = DateTime.Today;
            String today = FormatDate(thisDay.ToString("d"));

            if (ParseDate(message, thisDay))
            {
                DateTime.TryParseExact(message, new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out sDate);

                eventDetails.Add(message);//3 - Event Start Date
                
                PromptDialog.Text(context, this.ResumeAfterSDatePrompt, "When is the Event End Date(dd/mm/yyyy)? Enter '-' if it is the same as the start date.");
            }
            else
            {
                await context.PostAsync("The date that you have entered is not valid. Please try again.");
                PromptDialog.Text(context, this.ResumeAfterNamePrompt, "When is the Event Start Date(dd/mm/yyyy)?");
            }
        }

        private async Task ResumeAfterSDatePrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            if (!message.Equals("-"))
            {
                if (ParseDate(message, sDate))
                {
                    eventDetails.Add(message);//4 - Event End Date

                    DateTime.TryParseExact(message, new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out eDate);

                    totalDays = (eDate - sDate).Days + 1;

                    if (totalDays > 31)
                    {
                        await context.PostAsync("Currently I do not accept events lasting more than 31 days.");
                        PromptDialog.Text(context, this.ResumeAfterSDatePrompt, "When is the Event End Date(dd/mm/yyyy)? Enter '-' if it is the same as the start date.");
                    }
                    else
                    {
                        PromptDialog.Confirm(context, this.ResumeAfterEDatePrompt, "Do you require Attendance Code?", "Do you require Attendance Code?", 100, PromptStyle.Keyboard);

                    }
                }
                else
                {
                    await context.PostAsync("The date that you have entered is not valid. Please try again.");
                    PromptDialog.Text(context, this.ResumeAfterSDatePrompt, "When is the Event End Date(dd/mm/yyyy)? Enter '-' if it is the same as the start date.");
                }
            }
            else
            {
                eventDetails.Add(eventDetails[3]);//4 - Event End Date
                DateTime.TryParseExact(eventDetails[3], new string[] { "dd/MM/yyyy", "dd-MM-yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out eDate);

                totalDays = (eDate - sDate).Days + 1;
                PromptDialog.Confirm(context, this.ResumeAfterEDatePrompt, "Do you require Attendance Code?", "Do you require Attendance Code?", 100, PromptStyle.Keyboard);
            }
        }

        private async Task ResumeAfterEDatePrompt(IDialogContext context, IAwaitable<bool> result)
        {
            var message = await result;

            if (message)
            {
                string promptText = "";
                if (code1.Count == 0)
                {
                    promptText = "What is your first Attendance Code Start Time(e.g. 2359)?";
                }
                else
                {
                    promptText = "What is your second Attendance Code Start Time(e.g. 2359)?";
                }

                PromptDialog.Text(context, this.ResumeAfterASTimePrompt, promptText);
            }
            else
            {
                if(code1.Count == 0)
                {
                    eventDetails.Add("");//5 - AttendanceCode1StartTime
                    eventDetails.Add("");//6 - AttendanceCode1EndTime
                }
                if (code2.Count == 0)
                {
                    eventDetails.Add("");//7 - AttendanceCode2StartTime
                    eventDetails.Add("");//8 - AttendanceCode2EndTime
                }
                PromptDialog.Text(context, this.ResumeAfterSEndDatePrompt, "When is the Survey End Date(dd/mm/yyyy)? Enter '-' if it is the same as the end date.");
            }
        }

        private async Task ResumeAfterASTimePrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;
            startTime = message;

            string promptText = "";

            if (code1.Count == 0)
            {
                promptText = "What is your first Attendance Code End Time(e.g. 2359)?";
            }
            else
            {
                promptText = "What is your second Attendance Code End Time(e.g. 2359)?";
            }
            PromptDialog.Text(context, this.ResumeAfterAETimePrompt, promptText);
        }

        private async Task ResumeAfterAETimePrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;
            endTime = message;

            if (ParseTime(startTime, endTime))
            {
                if (code1.Count == 0)
                {
                    eventDetails.Add(startTime);//5 - AttendanceCode1StartTime
                    eventDetails.Add(endTime);//6 - AttendanceCode1EndTime
                    

                    for (int i = 0; i < totalDays; i++)
                    {
                        //Generate unique attendanceCode1
                        code1.Add(GenerateCode());
                    }

                    PromptDialog.Confirm(context, this.ResumeAfterEDatePrompt, "Do you require Attendance Code 2?", "Do you require Attendance Code 2?",100, PromptStyle.Keyboard);
                }
                else
                {
                    eventDetails.Add(startTime);//7 - AttendanceCode2StartTime
                    eventDetails.Add(endTime);//8 - AttendanceCode2EndTime
                    
                    for (int i = 0; i < totalDays; i++)
                    {
                        //Generate unique attendanceCode2
                        code2.Add(GenerateCode());
                    }
                    PromptDialog.Text(context, this.ResumeAfterSEndDatePrompt, "When is the Survey End Date(dd/mm/yyyy)? Enter '-' if it is the same as the end date.");
                }
            }
            else
            {
                await context.PostAsync("Your start and/or end time is/are not valid. Please try again.");

                string promptText = "";

                if (code1.Count == 0)
                {
                    promptText = "What is your first Attendance Code Start Time(e.g. 2359)?";
                }
                else
                {
                    promptText = "What is your second Attendance Code Start Time(e.g. 2359)?";
                }
                PromptDialog.Text(context, this.ResumeAfterASTimePrompt, promptText);
            }
        }

        private async Task ResumeAfterSEndDatePrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            string promptText = "";

            if (!message.Equals("-"))
            {
                if (ParseDate(message, sDate))
                {
                    eventDetails.Add(message);//9 - Survey End Date
                    
                    promptText = "What is your Survey End Time(e.g. 2359)? Enter '-' if you are ok with the default 2359hrs.";
                    PromptDialog.Text(context, this.ResumeAfterSEndTimePrompt, promptText);
                }
                else
                {
                    PromptDialog.Text(context, this.ResumeAfterSEndDatePrompt, "When is the Survey End Date(dd/mm/yyyy)? Enter '-' if it is the same as the end date.");

                }
            }
            else
            {
                eventDetails.Add(eventDetails[4]);//9 - Survey End Date

                promptText = "What is your Survey End Time(e.g. 2359)? Enter '-' if you are ok with the default 2359hrs.";
                PromptDialog.Text(context, this.ResumeAfterSEndTimePrompt, promptText);
            }
        }

        private async Task ResumeAfterSEndTimePrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            string promptText = "";

            if (!message.Equals("-"))
            {
                if (ParseTime("0000", message))
                {
                    eventDetails.Add(message);//10 - Survey End Time

                    promptText = "Enter your Survey Description. Enter '-' if you do not need any.";
                    PromptDialog.Text(context, this.ResumeAfterDescriptionPrompt, promptText);
                }
                else
                {
                    await context.PostAsync("Your Survey End Time is not valid. Please try again.");
                    promptText = "What is your Survey End Time(e.g. 2359)? Enter '-' if you are ok with the default 2359hrs.";
                    PromptDialog.Text(context, this.ResumeAfterSEndTimePrompt, promptText);
                }
            }
            else
            {
                eventDetails.Add("2359");//10 - Survey End Time
                
                promptText = "Enter your Survey Description. Enter '-' if you do not need any.";
                PromptDialog.Text(context, this.ResumeAfterDescriptionPrompt, promptText);
            }
        }

        private async Task ResumeAfterDescriptionPrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            string promptText = "";

            if (!message.Equals("-"))
            {
                eventDetails.Add(message);//11 - description
                
            }
            else
            {
                eventDetails.Add("");//11 - description
             
            }
            promptText = "Do you allow anonymity for the survey?";
            PromptDialog.Confirm(context, this.ResumeAfterAnonPrompt, promptText,promptText,100,PromptStyle.Keyboard);
        }

        private async Task ResumeAfterAnonPrompt(IDialogContext context, IAwaitable<bool> result)
        {
            var message = await result;

            string promptText = "What is your Email Address? Enter '-' if you do not wish to associate it with your email.";

            if (message)
            {
                eventDetails.Add("Y");//12 - anonymity
            }
            else
            {
                eventDetails.Add("N");//12 - anonymity
            }

            PromptDialog.Text(context, this.ResumeAfterEmailPrompt, promptText);
        }

        private async Task ResumeAfterEmailPrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            string promptText = "";

            if (!message.Equals("-"))
            {
                try
                {
                    System.Net.Mail.MailAddress mail = new System.Net.Mail.MailAddress(message);

                    eventDetails.Add(message);//13 - email
                    
                    promptText = "Do you need to share the event results with anyone?";
                    PromptDialog.Confirm(context, this.ResumeAfterSharePrompt, promptText, promptText, 100, PromptStyle.Keyboard);
                }
                catch(Exception e)
                {
                    await context.PostAsync("You have entered an invalid email! Please try again.");
                    promptText = "What is your Email Address? Enter '-' if you do not wish to associate it with your email.";
                    PromptDialog.Text(context, this.ResumeAfterEmailPrompt, promptText);
                }
            }
            else
            {
                eventDetails.Add("");//13 - email

                promptText = "Do you need to share the event results with anyone?";
                PromptDialog.Confirm(context, this.ResumeAfterSharePrompt, promptText, promptText, 100, PromptStyle.Keyboard);
            }
        }

        private async Task ResumeAfterSharePrompt(IDialogContext context, IAwaitable<bool> result)
        {
            var message = await result;

            string promptText = "What would you like to set as the password?";

            if (message)
            {
                PromptDialog.Text(context, this.ResumeAfterPasswordPrompt, promptText);
            }
            else
            {
                eventDetails.Add("");//14 - pw

                qCount++;
                promptText = "For survey question " + qCount + ", what is the question? Enter '-' to end the creation of event.";
                PromptDialog.Text(context, this.ResumeAfterSurveyPrompt, promptText);
            }
        }

        private async Task ResumeAfterPasswordPrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;

            string promptText = "";

            eventDetails.Add(Sha256(message));//14 - pw

            qCount++;
            promptText = "For survey question " + qCount + ", what is the question? Enter '-' to end the creation of event.";
            PromptDialog.Text(context, this.ResumeAfterSurveyPrompt, promptText);
        }

        private async Task ResumeAfterSurveyPrompt(IDialogContext context, IAwaitable<string> result)
        {
            var message = await result;
            EventEntity eventEntity = new EventEntity();
            if (message.Equals("-") && surveyList.Count ==0)
            {
                eventDetails.Add(GenerateCode());//15 - Survey Code

                var survey = new List<QuestionEntity>();
                QuestionEntity question = new QuestionEntity();
                var answerList = new List<string>();
                question = new QuestionEntity(eventDetails[15], "1");
                question.QuestionText = ("This survey has no question. Type anything to exit this survey.");
                question.AnswerList = JsonConvert.SerializeObject(answerList);
                survey.Add(question);

                eventDetails.Add(JsonConvert.SerializeObject(survey));//16 - Survey

                var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
                var tableClient = storageAccount.CreateCloudTableClient();
                CloudTable eventTable = tableClient.GetTableReference("Event");
                eventTable.CreateIfNotExists();

                eventDetails.Add("");//17 - Attendance Code 1
                eventDetails.Add("");//18 - Attendance Code 2
                eventDetails.Add("");//19 - Day

                for (int i = 0; i < totalDays; i++)
                {
                    if (code1.Count != 0)
                        eventDetails[17] = code1[i];//Set Attendance Code 1
                    if (code2.Count != 0)
                        eventDetails[18] = code2[i];//Set Attendance Code 2
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

                    eventDetails[1] = code;//Reset Row Key
                    eventDetails[19] = i+1 + "";//Set Day

                    eventEntity = new EventEntity(eventDetails[0], eventDetails[1]);
                    eventEntity.EventName = eventDetails[2];
                    eventEntity.EventStartDate = eventDetails[3];
                    eventEntity.EventEndDate = eventDetails[4];
                    eventEntity.AttendanceCode1StartTime = eventDetails[5];
                    eventEntity.AttendanceCode1EndTime = eventDetails[6];
                    eventEntity.AttendanceCode2StartTime = eventDetails[7];
                    eventEntity.AttendanceCode2EndTime = eventDetails[8];
                    eventEntity.SurveyEndDate = eventDetails[9];
                    eventEntity.SurveyEndTime = eventDetails[10];
                    eventEntity.Description = eventDetails[11];
                    eventEntity.Anonymous = eventDetails[12];
                    eventEntity.Email = eventDetails[13];
                    eventEntity.Password = eventDetails[14];
                    eventEntity.SurveyCode = eventDetails[15];
                    eventEntity.Survey = eventDetails[16];
                    eventEntity.AttendanceCode1 = eventDetails[17];
                    eventEntity.AttendanceCode2 = eventDetails[18];
                    eventEntity.Day = eventDetails[19];
                    TableOperation insertOperation1 = TableOperation.InsertOrMerge(eventEntity);
                    eventTable.Execute(insertOperation1);
                }

                string eventMsg = "'" + eventEntity.EventName + "' has been created, with a total of " + survey.Count + " questions. \n\n";
                eventMsg += "Date: " + eventEntity.EventStartDate + " to " + eventEntity.EventEndDate + "\n\n";
                eventMsg += "Attendance Code 1: \n\n";
                for (int i = 0; i < code1.Count; i++)
                {
                    eventMsg += "Day " + (i + 1) + ": " + (code1[i] == null ? "-" : code1[i]) + "\n\n";
                }
                eventMsg += "\n\n" + "Attendance Code 2:  \n\n";
                for (int i = 0; i < code2.Count; i++)
                {
                    eventMsg += "Day " + (i + 1) + ": " + (code2[i] == null ? "-" : code2[i]) + "\n\n";
                }
                eventMsg += "\n\n" + "Survey Code: " + eventEntity.SurveyCode + "\n\n\n\n"
                    + "Type /Event to manage your event. Thank you and have a great day ahead!";

                await context.PostAsync(eventMsg);
                bool a = SharedFunction.EmailEvent(eventEntity, context.UserData.GetValue<string>(ContextConstants.Name), survey, code1, code2);
                context.Done(this);

            } else if(message.Equals("-"))
            {
                eventDetails.Add(GenerateCode());//15 - Survey Code

                var survey = new List<QuestionEntity>();
                QuestionEntity question = new QuestionEntity();

                for (int i = 0; i < surveyList.Count; i++)
                {
                    var answerList = new List<string>();
                    question = new QuestionEntity(eventDetails[15], (i + 1).ToString());
                    question.QuestionText = surveyList[i];
                    if (optionList[i] == 2)
                    {
                        answerList.Add("Yes");
                        answerList.Add("No");

                    }
                    else if (optionList[i] == 3)
                    {
                        answerList.Add("Too Long");
                        answerList.Add("Just Right");
                        answerList.Add("Too Short");

                    }
                    else if (optionList[i] == 4)
                    {
                        answerList.Add("Strongly Agree");
                        answerList.Add("Agree");
                        answerList.Add("Somewhat Agree");
                        answerList.Add("Neutral");
                        answerList.Add("Somewhat Disagree");
                        answerList.Add("Disagree");
                        answerList.Add("Strongly Disagree");
                        answerList.Add("NA");
                    }
                    else if (optionList[i] == 5)
                    {
                        answerList.Add("Strongly Satisfied");
                        answerList.Add("Satisfied");
                        answerList.Add("Somewhat Satisfied");
                        answerList.Add("Neutral");
                        answerList.Add("Somewhat Dissatisfied");
                        answerList.Add("Dissatisfied");
                        answerList.Add("Strongly Dissatisfied");
                        answerList.Add("NA");
                    }
                    else if (optionList[i] == 6)
                    {
                        answerList.Add("Highly Relevant");
                        answerList.Add("Relevant");
                        answerList.Add("Somewhat Relevant");
                        answerList.Add("Neutral");
                        answerList.Add("Somewhat Irrelevant");
                        answerList.Add("Irrelevant");
                        answerList.Add("Highly Irrelevant");
                        answerList.Add("NA");
                    }
                    //Retrieving the answer list

                    question.AnswerList = JsonConvert.SerializeObject(answerList);
                    survey.Add(question);
                }
                eventDetails.Add(JsonConvert.SerializeObject(survey));//16 - Survey

                var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
                var tableClient = storageAccount.CreateCloudTableClient();
                CloudTable eventTable = tableClient.GetTableReference("Event");
                eventTable.CreateIfNotExists();
                eventDetails.Add("");//17 - Attendance Code 1
                eventDetails.Add("");//18 - Attendance Code 2
                eventDetails.Add("");//19 - Day

                for (int i = 0; i < totalDays; i++)
                {
                    if (code1.Count != 0)
                        eventDetails[17] = code1[i];//Set Attendance Code 1
                    if (code2.Count != 0)
                        eventDetails[18] = code2[i];//Set Attendance Code 2
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
                    eventDetails[1] = code;//Reset Row Key
                    eventDetails[19] = i + 1 + "";//Set Day
                    eventEntity = new EventEntity(eventDetails[0], eventDetails[1]);
                    eventEntity.EventName = eventDetails[2];
                    eventEntity.EventStartDate = eventDetails[3];
                    eventEntity.EventEndDate = eventDetails[4];
                    eventEntity.AttendanceCode1StartTime = eventDetails[5];
                    eventEntity.AttendanceCode1EndTime = eventDetails[6];
                    eventEntity.AttendanceCode2StartTime = eventDetails[7];
                    eventEntity.AttendanceCode2EndTime = eventDetails[8];
                    eventEntity.SurveyEndDate = eventDetails[9];
                    eventEntity.SurveyEndTime = eventDetails[10];
                    eventEntity.Description = eventDetails[11];
                    eventEntity.Anonymous = eventDetails[12];
                    eventEntity.Email = eventDetails[13];
                    eventEntity.Password = eventDetails[14];
                    eventEntity.SurveyCode = eventDetails[15];
                    eventEntity.Survey = eventDetails[16];
                    eventEntity.AttendanceCode1 = eventDetails[17];
                    eventEntity.AttendanceCode2 = eventDetails[18];
                    eventEntity.Day = eventDetails[19];
                    TableOperation insertOperation1 = TableOperation.InsertOrMerge(eventEntity);
                    eventTable.Execute(insertOperation1);
                }

                string eventMsg = "'" + eventEntity.EventName + "' has been created, with a total of " + survey.Count + " questions. \n\n";
                eventMsg += "Date: " + eventEntity.EventStartDate + " to " + eventEntity.EventEndDate + "\n\n";
                eventMsg += "Attendance Code 1: \n\n";
                for (int i = 0; i < code1.Count; i++)
                {
                    eventMsg += "Day " + (i + 1) + ": " + (code1[i] == null ? "-" : code1[i]) + "\n\n";
                }
                eventMsg += "\n\n" + "Attendance Code 2:  \n\n";
                for (int i = 0; i < code2.Count; i++)
                {
                    eventMsg += "Day " + (i + 1) + ": " + (code2[i] == null ? "-" : code2[i]) + "\n\n";
                }
                eventMsg += "\n\n" + "Survey Code: " + eventEntity.SurveyCode + "\n\n\n\n"
                    + "Type /Event to manage your event. Thank you and have a great day ahead!";

                await context.PostAsync(eventMsg);
                bool a = SharedFunction.EmailEvent(eventEntity, context.UserData.GetValue<string>(ContextConstants.Name), survey, code1, code2);
                context.Done(this);
            }
            else
            {
                surveyList.Add(message);
                string promptText = "";
                
                promptText = "For survey question " + qCount + ", what are the options?";
                promptText += "\n1. Open Ended Question(Users can freely answer)";
                promptText += "\n2. Yes|No";
                promptText += "\n3. Too Long|Just Right|Too Short";
                promptText += "\n4. Strongly Agree|...|Neutral|...|Strongly Disagree|NA(Total 8 choices)";
                promptText += "\n5. Strongly Satisfied|...|Neutral|...|Strongly Dissatisfied|NA(Total 8 choices)";
                promptText += "\n6. Highly Relevant|...|Neutral|...|Highly Irrelevant|NA(Total 8 choices)";
                PromptDialog.Number(context, this.ResumeAfterOptionsPrompt, promptText,promptText,100,null,1,6);
            }
        }

        private async Task ResumeAfterOptionsPrompt(IDialogContext context, IAwaitable<double> result)
        {
            var message = await result;
            int choice = Convert.ToInt32(message);
            optionList.Add(choice);

            string promptText = "";

            qCount++;
            promptText = "For survey question " + qCount + ", what is the question? Enter '-' to end the creation of event.";
            PromptDialog.Text(context, this.ResumeAfterSurveyPrompt, promptText);
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