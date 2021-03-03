using Microsoft.WindowsAzure.Storage.Table;
using System;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class EventEntity : TableEntity
    {
        public EventEntity(string creator,string eventId)
        {
            this.PartitionKey = creator;
            this.RowKey = eventId;
        }

        public EventEntity() { }

        public string AttendanceCode1StartTime { get; set; }
        public string AttendanceCode1EndTime { get; set; }
        public string AttendanceCode2StartTime { get; set; }
        public string AttendanceCode2EndTime { get; set; }
        public string AttendanceCode1 { get; set; }
        public string AttendanceCode2 { get; set; }
        public string SurveyCode { get; set; }
        public string EventStartDate { get; set; }
        public string EventEndDate { get; set; }
        public string SurveyEndDate { get; set; }
        public string SurveyEndTime { get; set; }
        public string EventName { get; set; } 
        public string Survey { get; set; }//Json object
        public string Description { get; set; }
        public string Anonymous { get; set; }
        public string Password { get; set; }
        public string Day { get; set; }
        public string Email { get; set; }
    }
}