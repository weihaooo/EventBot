using Microsoft.WindowsAzure.Storage.Table;
using System;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class TokenEntity : TableEntity
    {
        public TokenEntity(string token, string surveyCode)
        {
            this.PartitionKey = token;
            this.RowKey = surveyCode;
        }

        public TokenEntity() { }
    }
}