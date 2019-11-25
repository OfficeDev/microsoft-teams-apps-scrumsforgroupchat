// <copyright file="AdapterWithErrorHandler.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Scrum
{
    using System;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Logging;
    using Microsoft.Teams.Apps.Scrum.Properties;

    /// <summary>
    /// Implements Error Handler.
    /// </summary>
    public class AdapterWithErrorHandler : BotFrameworkHttpAdapter
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AdapterWithErrorHandler"/> class.
        /// Methods handles any kind of exception.
        /// </summary>
        /// <param name="credentialProvider">credentialProvider.</param>
        /// <param name="channelProvider">channelProvider.</param>
        /// <param name="logger">logger.</param>
        /// <param name="conversationState">conversationState.</param>
        public AdapterWithErrorHandler(ICredentialProvider credentialProvider, IChannelProvider channelProvider, ILogger<BotFrameworkHttpAdapter> logger, ConversationState conversationState = null)
            : base(credentialProvider, channelProvider, logger)
        {
            this.OnTurnError = async (turnContext, exception) =>
            {
                // Log any leaked exception from the application.
                logger.LogError($"Exception caught : {exception.Message}");

                // Send a catch-all appology to the user.
                await turnContext.SendActivityAsync(Resources.GenericErrorMessage);

                if (conversationState != null)
                {
                    try
                    {
                        // Delete the conversationState for the current conversation to prevent the
                        // bot from getting stuck in a error-loop caused by being in a bad state.
                        // ConversationState should be thought of as similar to "cookie-state" in a Web pages.
                        await conversationState.DeleteAsync(turnContext);
                    }
                    catch (Exception e)
                    {
                        logger.LogError($"Exception caught on attempting to Delete ConversationState : {e.Message}");
                    }
                }
            };
        }
    }
}
