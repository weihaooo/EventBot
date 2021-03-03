using System;
using System.Threading.Tasks;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.Description;
using Microsoft.Bot.Connector;
using Microsoft.Bot.Builder.Dialogs;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {
        [ResponseType(typeof(void))]
        public virtual async Task<HttpResponseMessage> Post([FromBody] Activity activity)
        {
            if (activity.GetActivityType() == ActivityTypes.Message)
            {
                await Conversation.SendAsync(activity, () => new RootDialog());
            }
            else if (activity.Type == ActivityTypes.Event)
            {
            }
            else
            {
                HandleSystemMessage(activity);
            }
            return new HttpResponseMessage(System.Net.HttpStatusCode.Accepted);
        }

        private Activity HandleSystemMessage(Activity message)
        {
            return null;
        }
    }

    public class Message
    {
        public ConversationReference RelatesTo { get; set; }
        public String Text { get; set; }
    }
}