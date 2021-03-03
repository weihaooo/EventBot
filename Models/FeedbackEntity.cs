using Microsoft.WindowsAzure.Storage.Table;
using System;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class FeedbackEntity : TableEntity
    {
        public FeedbackEntity(string code, string userid)
        {
            this.PartitionKey = code;
            this.RowKey = userid;
        }

        public FeedbackEntity() { }
        
        public string Date { get; set; }
        public string Name { get; set; }
        public string WorkshopName { get; set; }
        public string Response { get; set; }
        public string Survey { get; set; }
    }
}