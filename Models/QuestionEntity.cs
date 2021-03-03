using Microsoft.WindowsAzure.Storage.Table;
using System;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class QuestionEntity : TableEntity
    {
        public QuestionEntity(string code, string index)
        {
            this.PartitionKey = code;
            this.RowKey = index;
        }

        public QuestionEntity() { }

        public string QuestionText { get; set; }
        public string AnswerList { get; set; }//Json object
    }
}