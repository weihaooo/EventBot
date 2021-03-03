using Microsoft.Bot.Sample.ProactiveBot;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using ZXing;

namespace ProactiveBot
{
    [Serializable]
    public class SharedFunction
    {
        public void StoreResultsToList(string eventCode, List<string> eventList, List<string> questionList, List<string> questionTypeList, List<List<string>> answerList)
        {
            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);
            var tableClient = storageAccount.CreateCloudTableClient();
            List<QuestionEntity> surveyEntity = new List<QuestionEntity>();

            CloudTable feedbackTable = tableClient.GetTableReference("Event");
            TableQuery<EventEntity> feedbackQuery = new TableQuery<EventEntity>().Where(TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, eventCode));
            foreach (EventEntity entity in feedbackTable.ExecuteQuery(feedbackQuery))
            {
                eventList.Add(JsonConvert.SerializeObject(entity));
                surveyEntity = JsonConvert.DeserializeObject<List<QuestionEntity>>(entity.Survey); 
            }

            int qnCount = 1;
            foreach (QuestionEntity entity in surveyEntity)
            {
                questionList.Add("Q" + qnCount + ". " + entity.QuestionText);
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

        public static bool EmailEvent(EventEntity eventEntity, string name, List<QuestionEntity> survey, List<string> code1,List<string> code2)
        {
            string SendTo = eventEntity.Email;
            string Subject = "[Eventory] " + eventEntity.EventName +" Successfully Created";
            string MessageBody = "Hi,\n\n'" + eventEntity.EventName + "' has been created, with a total of " + survey.Count + " questions. \n\n";
            MessageBody += "Date: " + eventEntity.EventStartDate + " to " + eventEntity.EventEndDate + ", " + eventEntity.Day + " day(s)  \n\n";
            MessageBody += "Attendance Code 1(Valid Time From " + eventEntity.AttendanceCode1StartTime +"-"+eventEntity.AttendanceCode1EndTime+"): \n\n";
            for (int i = 0; i < code1.Count; i++)
            {
                MessageBody += "Day " + (i + 1) + ": " + (code1[i] == null ? "-" : code1[i]) + "\n\n";
            }
            MessageBody += "\n\n" + "Attendance Code 2(Valid Time From " + eventEntity.AttendanceCode2StartTime + "-" + eventEntity.AttendanceCode2EndTime + "):  \n\n";
            for (int i = 0; i < code2.Count; i++)
            {
                MessageBody += "Day " + (i + 1) + ": " + (code2[i] == null ? "-" : code2[i]) + "\n\n";
            }
            MessageBody += "\n\n" + "Survey Code: " + eventEntity.SurveyCode + " (Expiring on " + eventEntity.SurveyEndDate + " " + eventEntity.SurveyEndTime + ")\n\n\n\n";
            
            try
            {
                using (SmtpClient client = new SmtpClient(ConfigurationManager.AppSettings["EmailHost"], Convert.ToInt32(ConfigurationManager.AppSettings["EmailPort"])))
                {
                    // Configure the client
                    client.EnableSsl = true;
                    client.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["EmailId"], ConfigurationManager.AppSettings["EmailPwd"]);
                    
                    MailMessage message = new MailMessage(
                                             ConfigurationManager.AppSettings["EmailId"], // From field
                                             SendTo, // Recipient field
                                             Subject, // Subject of the email message
                                             MessageBody // Email message body
                                          );
                    var barcodeWriter = new BarcodeWriter();
                    barcodeWriter.Format = BarcodeFormat.QR_CODE;
                    MemoryStream ms;
                    Bitmap imageAsBytes;

                    var msList = new List<MemoryStream>();
                    for (int i = 0; i < code1.Count; i++) {
                        //Write qrcode into memorystream
                         ms = new MemoryStream();
                        imageAsBytes = barcodeWriter.Write(code1[i]);
                        imageAsBytes.Save(ms, ImageFormat.Jpeg);
                        ms.Position = 0;
                        msList.Add(ms);

                        var a = new Attachment(ms, "Attendance_Code_1_Day_" + (i + 1) + "_" + code1[i] + ".jpg", MediaTypeNames.Image.Jpeg);
                        
                        message.Attachments.Add(a);
                    }
                    for (int i = 0; i < code2.Count; i++)
                    {
                        //Write qrcode into memorystream
                        ms = new MemoryStream();
                        imageAsBytes = barcodeWriter.Write(code2[i]);
                        imageAsBytes.Save(ms, ImageFormat.Jpeg);
                        ms.Position = 0;
                        msList.Add(ms);

                        var a = new Attachment(ms, "Attendance_Code_2_Day_" + (i + 1) + "_" + code2[i] + ".jpg", MediaTypeNames.Image.Jpeg);
                        
                        message.Attachments.Add(a);
                    }
                    //Write qrcode into memorystream
                    ms = new MemoryStream();
                    imageAsBytes = barcodeWriter.Write(eventEntity.SurveyCode);
                    imageAsBytes.Save(ms, ImageFormat.Jpeg);
                    msList.Add(ms);
                    ms.Position = 0;

                    var attachment = new Attachment(ms, "Survey_Code_" + eventEntity.SurveyCode + ".jpg", MediaTypeNames.Image.Jpeg);
                    message.Attachments.Add(attachment);
                    client.Send(message);

                    foreach (var memorystream in msList)
                    {
                        memorystream.Dispose();
                    }
                    return true;
                }
            }
            catch (Exception e)
            {
                return false;
            }
        }
    }
}