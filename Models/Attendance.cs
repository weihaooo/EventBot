using Microsoft.WindowsAzure.Storage.Table;
using System;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class AttendanceEntity : TableEntity
    {
        public AttendanceEntity(string eventCode, string userId)
        {
            this.PartitionKey = eventCode;
            this.RowKey = userId;
        }

        public AttendanceEntity() { }
        
        public string SurveyCode { get; set; }
        public string Date { get; set; }
        public string Name { get; set; }
        public string EventName { get; set; }
        public bool Survey { get; set; }
        public bool Morning { get; set; }
        public bool Afternoon { get; set; }
    }
}