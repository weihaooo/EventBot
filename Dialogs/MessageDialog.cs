using Microsoft.Bot.Builder.Dialogs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class MessageDialog : IDialog<object>
    {
        private readonly string _messageToSend;

        public MessageDialog(string message)
        {
            _messageToSend = message;
        }

        public async Task StartAsync(IDialogContext context)
        {
            await context.PostAsync(_messageToSend);
            context.Done<object>(null);
        }
    }
}