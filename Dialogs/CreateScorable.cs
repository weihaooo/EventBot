using System;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder.Dialogs.Internals;
using Microsoft.Bot.Builder.Internals.Fibers;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Builder.Scorables.Internals;
using Microsoft.Bot.Builder.Dialogs;
using System.Text.RegularExpressions;
using ProactiveBot.Dialogs;

namespace Microsoft.Bot.Sample.ProactiveBot
{

#pragma warning disable 1998
    [Serializable]
    public class CreateScorable : ScorableBase<IActivity, string, double>
    {
        private readonly IDialogTask task;
        string abortCondition = "(?i)^(bye|cancel|exit|abort)$";
        string resultPattern = @"^\/result [a-zA-Z0-9]{8}";
        string qrPattern = @"^\/qr [a-zA-Z0-9]{8}";
        private bool resetStack = false;

        public CreateScorable(IDialogTask task)
        {
            SetField.NotNull(out this.task, nameof(task), task);
        }

        protected override async Task<string> PrepareAsync(IActivity activity, CancellationToken token)
        {
            var message = activity as IMessageActivity; 

            if (message != null && !string.IsNullOrWhiteSpace(message.Text))
            {
                if ((message.Text.Equals("/createevent", StringComparison.InvariantCultureIgnoreCase) || message.Text.Equals("/ce", StringComparison.InvariantCultureIgnoreCase)) ||
                        ((message.Text.Equals("/survey", StringComparison.InvariantCultureIgnoreCase)) || (message.Text.Equals("/event", StringComparison.InvariantCultureIgnoreCase))) ||
                        (message.Text.Equals("/result", StringComparison.InvariantCultureIgnoreCase)) || (Regex.IsMatch(message.Text, resultPattern)) || (Regex.IsMatch(message.Text, qrPattern)) ||
                         (message.Text.Equals("/help", StringComparison.InvariantCultureIgnoreCase) || message.Text.Equals("/info", StringComparison.InvariantCultureIgnoreCase)) ||
                         (Regex.IsMatch(message.Text, abortCondition)) || message.Text.Equals("/cevm", StringComparison.InvariantCultureIgnoreCase)) 
                {
                    return message.Text;
                } 
            } 
            return null;
        }

        protected override bool HasScore(IActivity item, string state)
        {
            return state != null;
        }

        protected override double GetScore(IActivity item, string state)
        {
            return 1.0;
        }

        protected override async Task PostAsync(IActivity item, string state, CancellationToken token)
        {
            var message = item as IMessageActivity;
            var messageToSend = string.Empty;

            if (message != null)
            {
                if (message.Text.Equals("/createevent", StringComparison.InvariantCultureIgnoreCase) || message.Text.Equals("/ce", StringComparison.InvariantCultureIgnoreCase))
                {
                    var createDialog = new CreateDialog();
                    var interruption = createDialog.Void<object, IMessageActivity>();
                    this.task.Call(interruption, null);
                }
                else if ((message.Text.Equals("/survey", StringComparison.InvariantCultureIgnoreCase)) || (message.Text.Equals("/event", StringComparison.InvariantCultureIgnoreCase)))
                {
                    var deleteEventDialog = new DeleteEventDialog();
                    var interruption = deleteEventDialog.Void<object, IMessageActivity>();
                    this.task.Call(interruption, null);
                }
                else if (message.Text.Equals("/result", StringComparison.InvariantCultureIgnoreCase))
                {
                    var resultDialog = new ResultDialog(null);
                    var interruption = resultDialog.Void<object, IMessageActivity>();
                    this.task.Call(interruption, null);
                }
                else if (message.Text.Equals("/cevm", StringComparison.InvariantCultureIgnoreCase))
                {
                    var cevmDialog = new CreateMobileDialog();
                    var interruption = cevmDialog.Void<object, IMessageActivity>();
                    this.task.Call(interruption, null);
                }
                else if (Regex.IsMatch(message.Text, resultPattern))
                {
                    string[] split = message.Text.Split(' ');
                    var resultDialog = new ResultDialog(split[1]);
                    var interruption = resultDialog.Void<object, IMessageActivity>();
                    this.task.Call(interruption, null);
                }
                else if (Regex.IsMatch(message.Text, qrPattern))
                {
                    string[] split = message.Text.Split(' ');
                    var resultDialog = new QRDialog(split[1]);
                    var interruption = resultDialog.Void<object, IMessageActivity>();
                    this.task.Call(interruption, null);
                }
                else if (message.Text.Equals("/help", StringComparison.InvariantCultureIgnoreCase) || message.Text.Equals("/info", StringComparison.InvariantCultureIgnoreCase))
                {
                    messageToSend = "Hi, these are commands you can use:\n\n"
                        + "1. Enter \"/event\" to manage the event created. \n\n"
                        + "2. Enter \"/result\" to view only the results of the events. \n\n"
                        + "3. Enter \"/result xxxxxxx\" (e.g. /result 58ab8a2) to view specific survey result. \n\n"
                        + "4. Enter \"/ce\" to receive a downloadable excel template to create an event. \n\n"
                        + "5. Enter \"/cevm\" create an event through question and answer style . \n\n"
                        + "6. Enter \"/qr\" with the attendance or survey code given (e.g. /qr 912ad823j) to retrieve the QR Code image. \n\n"
                        + "7. Enter the code given or upload QR code image to register a workshop attendance or do the survey. \n\n"
                        + "8. Enter \"exit\", \"bye\" or \"cancel\" to restart the conversation. \n\n"
                        + "Say something to continue! :)";

                    var newDialog = new MessageDialog(messageToSend);
                    var interruption = newDialog.Void<object, IMessageActivity>();
                    this.task.Call(interruption, null);
                    await this.task.PollAsync(token);

                } else if (Regex.IsMatch(message.Text, abortCondition))
                {
                    messageToSend = "Bye, talk to me again if you need my assistance!";
                    var newDialog = new MessageDialog(messageToSend);
                    var interruption = newDialog.Void<object, IMessageActivity>();
                    this.task.Call(interruption, null);
                    resetStack = true;
                    await this.task.PollAsync(token);
                }
            }
            await this.task.PollAsync(token); 
        } 

        protected override Task DoneAsync(IActivity item, string state, CancellationToken token)
        {
            if (resetStack)
            {
                this.task.Reset();
                resetStack = false;
            }
            return Task.CompletedTask;
        }
    }
}