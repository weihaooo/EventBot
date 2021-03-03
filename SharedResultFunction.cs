using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Sample.ProactiveBot;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace ProactiveBot
{
    [Serializable]
    public class SharedResultFunction
    {
        private static readonly Random random = new Random();
        private static readonly object syncLock = new object();

        public IMessageActivity exportResult(IDialogContext context, IDictionary<int, string> eventDict, EventEntity oneEventEntity, List<string> responseList, List<string> questionList, List<List<string>> answerList,
                                        List<string> questionTypeList, IDictionary<string, string> uniqueAnswerDict, IDictionary<string, int> attendanceDict)
        { 
            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();

            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            var blobContainer = blobClient.GetContainerReference("tempresults");
            blobContainer.CreateIfNotExists();

            // Set the permissions so the blobs are public. //ON HOLD
            BlobContainerPermissions permissions = new BlobContainerPermissions
            {
                PublicAccess = BlobContainerPublicAccessType.Blob
            };
            blobContainer.SetPermissions(permissions);
            var newBlockBlob = blobContainer.GetBlockBlobReference("new_" + DateTime.Now.ToString("MMddyyyy-hhmmss") + ".xlsx");

            XSSFWorkbook wb = new XSSFWorkbook();

            computeResult(context.Activity.From.Id, eventDict, wb, "ResponseSummary", responseList, questionTypeList, uniqueAnswerDict, questionList, answerList);
            computeResult(context.Activity.From.Id, eventDict, wb, "ComputeSummary", responseList, questionTypeList, uniqueAnswerDict, questionList, answerList, "Percent");
            computeScore(context.Activity.From.Id, eventDict, wb, "ScoreSummary", responseList, questionTypeList, questionList, answerList);
            computeRawResult(wb, "RawResult", questionList, responseList);
            attendanceResult(oneEventEntity.SurveyCode, wb, "Attendance");

            string ExcelFileName = string.Format("{0}.xlsx", "SurveyResult");

            using (var memoryStream = new MemoryStream())
            {
                wb.Write(memoryStream);
                using (var ms = new MemoryStream(memoryStream.ToArray()))
                {
                    newBlockBlob.UploadFromStream(ms);
                    newBlockBlob.Properties.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    newBlockBlob.SetProperties();
                }
            }

            CloudTable tokenTable = tableClient.GetTableReference("Token");
            tokenTable.CreateIfNotExists();
            TableQuery<TokenEntity> query = new TableQuery<TokenEntity>().Where(TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, oneEventEntity.SurveyCode));
            string token = "";
            foreach (TokenEntity t in tokenTable.ExecuteQuery(query))
            {
                token = t.PartitionKey;
            }

            if (token.Equals(""))
            {
                var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
                var stringChars = new char[32];
                var random = new Random();

                for (int i = 0; i < stringChars.Length; i++)
                {
                    stringChars[i] = chars[random.Next(chars.Length)];
                }

                token = new String(stringChars);
                TokenEntity tokenEntity = new TokenEntity(token, oneEventEntity.SurveyCode);
                TableOperation insertOperation = TableOperation.Insert(tokenEntity);
                tokenTable.Execute(insertOperation);
            }

            var newMessage = context.MakeMessage();
            newMessage.Text = "Here is your result summary for " + oneEventEntity.EventName + "!  \n"
                            + "Date: " + oneEventEntity.EventStartDate + "  \n"
                            + "Morning Attendance: " + attendanceDict["Morning"] + "  \n"
                            + "Afternoon Attendance: " + attendanceDict["Afternoon"] + "  \n"
                            + "Survey Participants: " + attendanceDict["Survey"] + "  \n"
                            + "Link to live results: https://jtceventdashboard.azurewebsites.net/Home/Index?token=" + token + "  \n";

            newMessage.Attachments.Add(new Attachment()
            {
                ContentUrl = newBlockBlob.Uri.ToString(),
                ContentType = newBlockBlob.Properties.ContentType,
                Name = newBlockBlob.Name
            });

            return newMessage;
        }

        public bool emailResult(IDialogContext context, IDictionary<int, string> eventDict, EventEntity oneEventEntity, List<string> responseList, List<string> questionList, List<List<string>> answerList,
                                        List<string> questionTypeList, IDictionary<string, string> uniqueAnswerDict, IDictionary<string, int> attendanceDict)
        {
            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();

            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            var blobContainer = blobClient.GetContainerReference("tempresults");
            blobContainer.CreateIfNotExists();

            // Set the permissions so the blobs are public. //ON HOLD
            BlobContainerPermissions permissions = new BlobContainerPermissions
            {
                PublicAccess = BlobContainerPublicAccessType.Blob
            };
            blobContainer.SetPermissions(permissions);
            var newBlockBlob = blobContainer.GetBlockBlobReference("new_" + DateTime.Now.ToString("MMddyyyy-hhmmss") + ".xlsx");

            XSSFWorkbook wb = new XSSFWorkbook();

            computeResult(context.Activity.From.Id, eventDict, wb, "ResponseSummary", responseList, questionTypeList, uniqueAnswerDict, questionList, answerList);
            computeResult(context.Activity.From.Id, eventDict, wb, "ComputeSummary", responseList, questionTypeList, uniqueAnswerDict, questionList, answerList, "Percent");
            computeScore(context.Activity.From.Id, eventDict, wb, "ScoreSummary", responseList, questionTypeList, questionList, answerList);
            computeRawResult(wb, "RawResult", questionList, responseList);
            attendanceResult(oneEventEntity.SurveyCode, wb, "Attendance");

            string ExcelFileName = string.Format("{0}.xlsx", "SurveyResult");

            CloudTable tokenTable = tableClient.GetTableReference("Token");
            tokenTable.CreateIfNotExists();
            TableQuery<TokenEntity> query = new TableQuery<TokenEntity>().Where(TableQuery.GenerateFilterCondition("RowKey", QueryComparisons.Equal, oneEventEntity.SurveyCode));
            string token = "";
            foreach (TokenEntity t in tokenTable.ExecuteQuery(query))
            {
                token = t.PartitionKey;
            }

            if (token.Equals(""))
            {
                var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
                var stringChars = new char[32];
                var random = new Random();

                for (int i = 0; i < stringChars.Length; i++)
                {
                    stringChars[i] = chars[random.Next(chars.Length)];
                }

                token = new String(stringChars);
                TokenEntity tokenEntity = new TokenEntity(token, oneEventEntity.SurveyCode);
                TableOperation insertOperation = TableOperation.Insert(tokenEntity);
                tokenTable.Execute(insertOperation);
            }

            using (var memoryStream = new MemoryStream())
            {
                wb.Write(memoryStream);
                using (var ms = new MemoryStream(memoryStream.ToArray()))
                {
                    string SendTo = oneEventEntity.Email;
                    string Subject = "[Eventory] " + oneEventEntity.EventName + " Survey Results";
                    string MessageBody = "Hi,\n\nHere is your result summary for " + oneEventEntity.EventName + "!  \n"
                                    + "Date: " + oneEventEntity.EventStartDate + "  \n"
                                    + "Morning Attendance: " + attendanceDict["Morning"] + "  \n"
                                    + "Afternoon Attendance: " + attendanceDict["Afternoon"] + "  \n"
                                    + "Survey Participants: " + attendanceDict["Survey"] + "  \n"
                                    + "Link to live results: https://jtceventdashboard.azurewebsites.net/Home/Index?token=" + token + "  \n";

                    try
                    {
                        using (System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient(ConfigurationManager.AppSettings["EmailHost"], Convert.ToInt32(ConfigurationManager.AppSettings["EmailPort"])))
                        {
                            client.EnableSsl = true;
                            client.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["EmailId"], ConfigurationManager.AppSettings["EmailPwd"]);
                            
                            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage(
                                                     ConfigurationManager.AppSettings["EmailId"], // From field
                                                     SendTo, // Recipient field
                                                     Subject, // Subject of the email message
                                                     MessageBody // Email message body
                                                  );

                            System.Net.Mail.Attachment a = new System.Net.Mail.Attachment(ms, "results.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                            message.Attachments.Add(a);
                            client.Send(message);

                            return true;
                        }
                    }
                    catch (Exception e)
                    {
                        Debug.WriteLine("Send Email: " + e);
                        return false;
                    }
                }
            }

        }
        public void StoreResultsToList(string surveyCode, List<string> responseList, List<string> questionList, List<List<string>> answerList, 
                                        List<string> questionTypeList, IDictionary<string, string> uniqueAnswerDict, IDictionary<string, int> attendanceDict)
        { 
            attendanceDict.Clear();

            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            List<QuestionEntity> surveyEntity = new List<QuestionEntity>();

            CloudTable feedbackTable = tableClient.GetTableReference("Feedback");
            TableQuery<FeedbackEntity> feedbackQuery = new TableQuery<FeedbackEntity>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, surveyCode));
            foreach (FeedbackEntity entity in feedbackTable.ExecuteQuery(feedbackQuery))
            {
                responseList.Add(JsonConvert.SerializeObject(entity));
                surveyEntity = JsonConvert.DeserializeObject<List<QuestionEntity>>(entity.Survey);
            }

            foreach (QuestionEntity entity in surveyEntity)
            { 
                questionList.Add(entity.QuestionText);
                if (JsonConvert.DeserializeObject<List<string>>(entity.AnswerList).Count != 0)
                {
                    answerList.Add(JsonConvert.DeserializeObject<List<string>>(entity.AnswerList));
                    questionTypeList.Add("2");

                    //Collate same answers group together
                    if (uniqueAnswerDict.ContainsKey(entity.AnswerList))
                    {
                        //uniqueAnswerDict[entity.AnswerList] += 1;
                    }
                    else
                    {
                        byte rdmRed, rdmGreen, rdmBlue;
                        lock (syncLock)
                        {
                            rdmRed = Convert.ToByte(random.Next(200, 255));
                        }
                        lock (syncLock)
                        {
                            rdmGreen = Convert.ToByte(random.Next(150, 255));
                        }
                        lock (syncLock)
                        {
                            rdmBlue = Convert.ToByte(random.Next(150, 255));
                        }
                        byte[] rgb = new byte[] { rdmRed, rdmGreen, rdmBlue };
                        

                        uniqueAnswerDict.Add(entity.AnswerList, JsonConvert.SerializeObject(rgb));
                    }
                }
                else
                {
                    answerList.Add(new List<string>());
                    questionTypeList.Add("1");
                }
            }

            //For Morning / Afternoon
            storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            tableClient = storageAccount.CreateCloudTableClient();
            CloudTable attendanceTable = tableClient.GetTableReference("Attendance");
            TableQuery<AttendanceEntity> query = new TableQuery<AttendanceEntity>().Where(TableQuery.GenerateFilterCondition("SurveyCode", QueryComparisons.Equal, surveyCode));

            attendanceDict.Add("Morning", 0);
            attendanceDict.Add("Afternoon", 0);
            attendanceDict.Add("Survey", 0);

            foreach (AttendanceEntity entity in attendanceTable.ExecuteQuery(query))
            {
                if (entity.Morning)
                    attendanceDict["Morning"] += 1;

                if (entity.Afternoon)
                    attendanceDict["Afternoon"] += 1;

                if (entity.Survey)
                    attendanceDict["Survey"] += 1;
            }

        }

        public void computeResult(string uniqueID, IDictionary<int, string> eventDict, XSSFWorkbook wb, string sheetName, List<string> responseList,
                                    List<string> questionTypeList, IDictionary<string, string> uniqueAnswerDict,
                                    List<string> questionList, List<List<string>> answerList, string resultType = "Normal")
        {
            //Create bold font for excel
            XSSFCellStyle xStyle = boldFont(wb);

            int rowCount = 0;
            int totalNumChoiceLength = 0;

            if (resultType != "Percent")
            {
                XSSFCreationHelper createHelper = (XSSFCreationHelper)wb.GetCreationHelper();
            }

            XSSFSheet summarySheet = (XSSFSheet)wb.CreateSheet(sheetName);  
            //Group all the different answers list together
            foreach (KeyValuePair<string, string> entry in uniqueAnswerDict)
            { 
                // Create a row and put some cells in it. Rows are 0 based.
                XSSFRow row = (XSSFRow)summarySheet.CreateRow(rowCount);
                 
                xStyle = boldFont(wb);
                XSSFColor color = new XSSFColor();
                color.SetRgb(JsonConvert.DeserializeObject<byte[]>(entry.Value));
                xStyle.SetFillForegroundColor(color);
                xStyle.FillPattern = FillPattern.SolidForeground;

                int cellCount = 1;
                List<string> answerChoiceList = JsonConvert.DeserializeObject<List<string>>(entry.Key);

                if (answerChoiceList.Count > totalNumChoiceLength)
                    totalNumChoiceLength = answerChoiceList.Count;
                  
                foreach (string choice in answerChoiceList)
                {  
                    row.CreateCell(cellCount).SetCellValue(choice);
                    row.GetCell(cellCount).CellStyle = xStyle; 
                    cellCount += 1;
                }
                rowCount += 1; 
            }

            xStyle = boldFont(wb);

            //xStyle.FillPattern = FillPattern.NoFill;
            List<string> openEndedList = new List<string>();
            Dictionary<int, List<string>> openEndedResponseDict = new Dictionary<int, List<string>>();
            Dictionary<string, int> choiceResponseDict = new Dictionary<string, int>(); 

            //Get a list of response
            foreach (var response in responseList)
            {
                FeedbackEntity feedbackEntity = JsonConvert.DeserializeObject<FeedbackEntity>(response);
                List<string> feedbackResponse = JsonConvert.DeserializeObject<List<string>>(feedbackEntity.Response);

                int qnNo = 1;
                foreach (string feedback in feedbackResponse)
                {
                    List<string> openEndedResponse = new List<string>();

                    if (questionTypeList[qnNo - 1] == "1")   //OpenEnded Question
                    {
                        //Store it in a string so that you can write it in later
                        if (openEndedResponseDict.ContainsKey(qnNo))
                        {
                            openEndedResponseDict[qnNo].Add(feedback);
                        }
                        else
                        {
                            openEndedResponse.Add(feedback);
                            openEndedResponseDict.Add(qnNo, openEndedResponse);
                        }
                    }
                    else if (questionTypeList[qnNo - 1] == "2")      //Multiple Choice
                    {
                        if (choiceResponseDict.ContainsKey(qnNo + "||" + feedback))
                        {
                            choiceResponseDict[qnNo + "||" + feedback] += 1;
                        }
                        else
                        {
                            choiceResponseDict.Add(qnNo + "||" + feedback, 1);
                        }
                    }
                    qnNo += 1;
                }
            }
             
            //Print out all the question number and text
            int qnCount = 1;
            foreach (var qnText in questionList)
            { 
                if (questionTypeList[qnCount - 1] == "1")   //OpenEnded Question
                {
                    //Store it in a string so that you can write it in later
                    openEndedList.Add("Q" + qnCount + "." + qnText);
                }
                else if (questionTypeList[qnCount - 1] == "2")      //Multiple Choice
                {
                    xStyle = boldFont(wb);
                    XSSFColor color = new XSSFColor();
                    color.SetRgb(JsonConvert.DeserializeObject<byte[]>(uniqueAnswerDict[JsonConvert.SerializeObject(answerList[qnCount - 1])]));
                    xStyle.SetFillForegroundColor(color);
                    xStyle.FillPattern = FillPattern.SolidForeground; 

                    XSSFRow row = (XSSFRow)summarySheet.CreateRow(rowCount);
                    row.CreateCell(0).SetCellValue("Q" + qnCount + ". ");
                    row.GetCell(0).CellStyle = xStyle;
                    row.CreateCell(0 + totalNumChoiceLength + 1).SetCellValue(qnText);

                    xStyle = normalFont(wb);
                    color = new XSSFColor();
                    color.SetRgb(JsonConvert.DeserializeObject<byte[]>(uniqueAnswerDict[JsonConvert.SerializeObject(answerList[qnCount - 1])]));
                    xStyle.SetFillForegroundColor(color);
                    xStyle.FillPattern = FillPattern.SolidForeground;

                    for (int i = 0; i < answerList[qnCount - 1].Count; i++)
                    {  
                        //Print out all the question response 
                        if (choiceResponseDict.ContainsKey(qnCount + "||" + answerList[qnCount - 1][i]))
                        { 
                            double displayValue = 0;
                            if (resultType == "Percent")
                            {
                                double value = choiceResponseDict[qnCount + "||" + answerList[qnCount - 1][i]];
                                displayValue = value / responseList.Count * 100;
                                row.CreateCell(1 + i).SetCellValue(Math.Round(displayValue, 2) + "%");
                            }
                            else
                            {
                                displayValue = choiceResponseDict[qnCount + "||" + answerList[qnCount - 1][i]];
                                row.CreateCell(1 + i).SetCellValue(Math.Round(displayValue, 2));
                            } 
                            row.GetCell(1 + i).CellStyle = xStyle;
                        }
                        else
                        {
                            row.CreateCell(1 + i).SetCellValue(0);
                            row.GetCell(1 + i).CellStyle = xStyle;
                        }
                        xStyle.SetFillForegroundColor(color);   //to end off the color xStyle
                    }
                    rowCount += 1; 
                }
                qnCount++;
            }

            summarySheet.CreateRow(rowCount);
            rowCount++;

            xStyle = boldFont(wb);
            //Print out all the openEnded Questions
            foreach (var openEnded in openEndedList)
            {
                XSSFRow row = (XSSFRow)summarySheet.CreateRow(rowCount);
                row.CreateCell(0).SetCellValue(openEnded);
                row.GetCell(0).CellStyle = xStyle;
                rowCount += 1; 

                int qnNo = Convert.ToInt32(openEnded.Split('.')[0].Remove(0, 1));
                //Create rows for response answers
                if (openEndedResponseDict.ContainsKey(qnNo))
                {
                    foreach (var response in openEndedResponseDict[qnNo])
                    { 
                        xStyle = normalFont(wb);
                        row = (XSSFRow)summarySheet.CreateRow(rowCount);
                        row.CreateCell(1).SetCellValue(response);
                        row.GetCell(1).CellStyle = xStyle;
                        rowCount += 1;
                    }
                }
                else
                    rowCount += 1;
            }

            for (int i = 1; i <= 15; i++) // this will aply it form col 1 to 10
            {
                summarySheet.AutoSizeColumn(i);
            }
        }

        public void computeScore(string uniqueID, IDictionary<int, string> eventDict, XSSFWorkbook wb, string sheetName, 
                                List<string> responseList, List<string> questionTypeList, List<string> questionList, List<List<string>> answerList)
        {
            //Create bold font for excel
            XSSFCellStyle xStyle = boldFont(wb);

            int rowCount = 0;

            XSSFSheet summarySheet = (XSSFSheet)wb.CreateSheet(sheetName);

            Dictionary<int, List<string>> openEndedResponseDict = new Dictionary<int, List<string>>();
            Dictionary<string, int> choiceResponseDict = new Dictionary<string, int>();

            XSSFRow topRow = (XSSFRow)summarySheet.CreateRow(rowCount);
            topRow.CreateCell(1).SetCellValue("Asc");
            topRow.GetCell(1).CellStyle = xStyle;
            topRow.CreateCell(2).SetCellValue("Asc");
            topRow.GetCell(2).CellStyle = xStyle;
            topRow.CreateCell(3).SetCellValue("Desc");
            topRow.GetCell(3).CellStyle = xStyle;
            topRow.CreateCell(4).SetCellValue("Desc");
            topRow.GetCell(4).CellStyle = xStyle;
            rowCount++;
            //Get a list of response

            foreach (var response in responseList)
            {
                FeedbackEntity feedbackEntity = JsonConvert.DeserializeObject<FeedbackEntity>(response);
                List<string> feedbackResponse = JsonConvert.DeserializeObject<List<string>>(feedbackEntity.Response);

                int qnNo = 1;
                foreach (string feedback in feedbackResponse)
                {
                    if (questionTypeList[qnNo - 1] == "2")      //Multiple Choice
                    {
                        if (choiceResponseDict.ContainsKey(qnNo + "||" + feedback))
                        {
                            choiceResponseDict[qnNo + "||" + feedback] += 1;
                        }
                        else
                        {
                            choiceResponseDict.Add(qnNo + "||" + feedback, 1);
                        }
                    }
                    qnNo += 1;
                }
            }

            //Print out all the question number and text
            int qnCount = 1;
            foreach (var qnText in questionList)
            {
                if (questionTypeList[qnCount - 1] == "2")      //Multiple Choice
                {
                    int naCount = 0;
                    XSSFRow row = (XSSFRow)summarySheet.CreateRow(rowCount);
                    row.CreateCell(0).SetCellValue("Q" + qnCount + ". ");
                    row.GetCell(0).CellStyle = xStyle;
                    row.CreateCell(5).SetCellValue(qnText);
                    var ascValue = 0.00;
                    var descValue = 0.00;
                    var answerListcount = answerList[qnCount - 1].Count;
                    if (answerList[qnCount - 1][answerListcount - 1] == "NA" || answerList[qnCount - 1][answerListcount - 1] == "Not Applicable" || answerList[qnCount - 1][answerListcount - 1] == "N.A.")
                    {
                        if (choiceResponseDict.ContainsKey(qnCount + "||" + answerList[qnCount - 1][answerListcount - 1]))
                        {
                            naCount = choiceResponseDict[qnCount + "||" + answerList[qnCount - 1][answerListcount - 1]];
                        }
                        answerListcount -= 1;
                    }
                    for (int i = 0; i < answerListcount; i++)
                    {
                        //Print out all the question response 
                        if (choiceResponseDict.ContainsKey(qnCount + "||" + answerList[qnCount - 1][i]))
                        {
                            double displayValue = 0;

                            double value = choiceResponseDict[qnCount + "||" + answerList[qnCount - 1][i]];//No. of ppl who choose option
                            displayValue = value / (responseList.Count - naCount);//Total Response;
                            ascValue += (displayValue * (i + 1));
                            descValue += (displayValue * (answerListcount - i));

                        }
                    }
                    row.CreateCell(1).SetCellValue(Math.Round(ascValue, 2));
                    row.CreateCell(2).SetCellValue(Math.Round(ascValue / answerListcount * 100, 2) + "%");
                    row.CreateCell(3).SetCellValue(Math.Round(descValue, 2));
                    row.CreateCell(4).SetCellValue(Math.Round(descValue / answerListcount * 100, 2) + "%");
                    rowCount += 1;
                }
                qnCount++;
            }

            for (int i = 1; i <= 15; i++) // this will aply it form col 1 to 10
            {
                summarySheet.AutoSizeColumn(i);
            }
        }

        public void computeRawResult(XSSFWorkbook wb, string sheetName, List<string> questionList, List<string> responseList)
        {
            //Create bold font for excel
            XSSFCellStyle xStyle = boldFont(wb);

            int rowCount = 0;
            int fixColCount = 4;
            
            XSSFSheet summarySheet = (XSSFSheet)wb.CreateSheet(sheetName);

            // Create a row and put some cells in it. Rows are 0 based.
            XSSFRow row = (XSSFRow)summarySheet.CreateRow(rowCount);
            //Create Titles for the raw results
            row.CreateCell(0).SetCellValue("No.");
            row.GetCell(0).CellStyle = xStyle;
            row.CreateCell(1).SetCellValue("surveyCode");
            row.GetCell(1).CellStyle = xStyle;
            row.CreateCell(2).SetCellValue("Date");
            row.GetCell(2).CellStyle = xStyle;
            row.CreateCell(3).SetCellValue("Name");
            row.GetCell(3).CellStyle = xStyle;

            //Get the total number of questions
            int colCount = fixColCount;
            for (int i = 0; i < questionList.Count; i++)
            {
                row.CreateCell(colCount + i).SetCellValue("Q" + (i + 1));
                row.GetCell(colCount + i).CellStyle = xStyle;
            }
            //row.RowStyle.SetFont(font);
            rowCount += 1;
            //Get all the response and print out to the columns
            int responseNo = 1;
            foreach (var response in responseList)
            {
                row = (XSSFRow)summarySheet.CreateRow(rowCount);
                FeedbackEntity feedbackEntity = JsonConvert.DeserializeObject<FeedbackEntity>(response);
                row.CreateCell(0).SetCellValue(responseNo);
                row.CreateCell(1).SetCellValue(feedbackEntity.PartitionKey);
                row.CreateCell(2).SetCellValue(feedbackEntity.Date);
                row.CreateCell(3).SetCellValue(feedbackEntity.Name);
                //row.RowStyle.SetFont(font);

                List<string> responseAnswerList = JsonConvert.DeserializeObject<List<string>>(feedbackEntity.Response);
                colCount = fixColCount;
                foreach (var feedback in responseAnswerList)
                {
                    row.CreateCell(colCount).SetCellValue(feedback);
                    colCount += 1;
                }

                responseNo += 1;
                rowCount += 1;
            }

            for (int i = 1; i <= 15; i++) // this will aply it form col 1 to 10
            {
                summarySheet.AutoSizeColumn(i);
            }
        }

        public void attendanceResult(string surveyCode, XSSFWorkbook wb, string sheetName)
        {
            //Create bold font for excel
            XSSFCellStyle xStyle = boldFont(wb);

            int rowCount = 4;

            //XSSFCreationHelper createHelper = (XSSFCreationHelper)wb.GetCreationHelper(); 
            XSSFSheet summarySheet = (XSSFSheet)wb.CreateSheet(sheetName);

            // Create a row and put some cells in it. Rows are 0 based.
            XSSFRow row = (XSSFRow)summarySheet.CreateRow(rowCount);

            //Create Titles for the raw results
            row.CreateCell(0).SetCellValue("No.");
            row.GetCell(0).CellStyle = xStyle;
            row.CreateCell(1).SetCellValue("EventCode");
            row.GetCell(1).CellStyle = xStyle;
            row.CreateCell(2).SetCellValue("Date");
            row.GetCell(2).CellStyle = xStyle;
            row.CreateCell(3).SetCellValue("Name");
            row.GetCell(3).CellStyle = xStyle;

            row.CreateCell(4).SetCellValue("Morning");
            row.GetCell(4).CellStyle = xStyle;
            row.CreateCell(5).SetCellValue("Afternoon");
            row.GetCell(5).CellStyle = xStyle;
            row.CreateCell(6).SetCellValue("Survey");
            row.GetCell(6).CellStyle = xStyle;

            rowCount += 1;

            //Get results from attendance
            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            CloudTable attendanceTable = tableClient.GetTableReference("Attendance");
            TableQuery<AttendanceEntity> query = new TableQuery<AttendanceEntity>().Where(TableQuery.GenerateFilterCondition("SurveyCode", QueryComparisons.Equal, surveyCode));

            Dictionary<string, int> attendanceDict = new Dictionary<string, int>();
            attendanceDict.Add("Morning", 0);
            attendanceDict.Add("Afternoon", 0);
            attendanceDict.Add("Survey", 0);

            int no = 1;
            foreach (AttendanceEntity entity in attendanceTable.ExecuteQuery(query))
            {
                row = (XSSFRow)summarySheet.CreateRow(rowCount);

                row.CreateCell(0).SetCellValue(no);
                row.CreateCell(1).SetCellValue(entity.PartitionKey);
                row.CreateCell(2).SetCellValue(entity.RowKey);
                row.CreateCell(3).SetCellValue(entity.Name);
                var morning = "No";
                var afternoon = "No";
                var survey = "No";
                if (entity.Morning)
                {
                    attendanceDict["Morning"] += 1;
                    morning = "Yes";
                }

                if (entity.Afternoon)
                {
                    attendanceDict["Afternoon"] += 1;
                    afternoon = "Yes";
                }

                if (entity.Survey)
                {
                    attendanceDict["Survey"] += 1;
                    survey = "Yes";
                }
                row.CreateCell(4).SetCellValue(morning);
                row.CreateCell(5).SetCellValue(afternoon);
                row.CreateCell(6).SetCellValue(survey);

                rowCount += 1;
                no++;
            }

            //Print out the attendance  
            List<string> attendanceType = new List<string>() { "Morning", "Afternoon", "Survey" };
            XSSFRow newRow;
            XSSFCell cell;
            for (int i = 0; i < attendanceType.Count; i++)
            {
                newRow = (XSSFRow)summarySheet.GetRow(i);
                if (newRow == null) newRow = (XSSFRow)summarySheet.CreateRow(i);
                cell = (XSSFCell)newRow.GetCell(0);
                if (cell == null) newRow.CreateCell(0).SetCellValue(attendanceType[i]);
                newRow.GetCell(0).CellStyle = xStyle;

                newRow = (XSSFRow)summarySheet.GetRow(i);
                if (newRow == null) newRow = (XSSFRow)summarySheet.CreateRow(i);
                cell = (XSSFCell)newRow.GetCell(1);
                if (cell == null) newRow.CreateCell(1).SetCellValue(attendanceDict[attendanceType[i]]);
                newRow.GetCell(1).CellStyle = xStyle;
            }

            for (int i = 1; i <= 15; i++) // this will aply it form col 1 to 10
            {
                summarySheet.AutoSizeColumn(i);
            }
        }

        public XSSFCellStyle boldFont(XSSFWorkbook wb)
        {
            var font = wb.CreateFont();
            font.Boldweight = (short)FontBoldWeight.Bold;
            XSSFCellStyle xStyle = (XSSFCellStyle)wb.CreateCellStyle();
            xStyle.SetFont(font);
            return xStyle;
        }

        public XSSFCellStyle normalFont(XSSFWorkbook wb)
        {
            var font = wb.CreateFont();
            font.Boldweight = (short)FontBoldWeight.Normal;
            XSSFCellStyle xStyle = (XSSFCellStyle)wb.CreateCellStyle();
            xStyle.SetFont(font);
            return xStyle;
        } 

        public string Sha256(string randomString)
        {
            var crypt = new System.Security.Cryptography.SHA256Managed();
            var hash = new System.Text.StringBuilder();
            byte[] crypto = crypt.ComputeHash(Encoding.UTF8.GetBytes(randomString));
            foreach (byte theByte in crypto)
            {
                hash.Append(theByte.ToString("x2"));
            }
            return hash.ToString();
        }
    }
}