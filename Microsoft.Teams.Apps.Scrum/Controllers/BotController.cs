// <copyright file="BotController.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Scrum.Controller
{
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Integration.AspNet.Core;

    /// <summary>
    /// This is a Bot controller class includes all API's related to this Bot.
    /// </summary>
    [Route("api/messages")]
    [ApiController]
    public class BotController : ControllerBase
    {
        private readonly IBotFrameworkHttpAdapter adapter;
        private readonly IBot bot;

        /// <summary>
        /// Initializes a new instance of the <see cref="BotController"/> class.
        /// </summary>
        /// <param name="adapter">Bot adapter.</param>
        /// <param name="bot"> Bot Interface.</param>
        /// <param name="env"> HostingEnvironment.</param>
        public BotController(IBotFrameworkHttpAdapter adapter, IBot bot)
        {
            this.adapter = adapter;
            this.bot = bot;
        }

        /// <summary>
        /// Executing the Post Async method.
        /// </summary>
        /// <returns>A <see cref="Task"/> representing the result of the asynchronous operation.</returns>
        [HttpPost]
        public async Task PostAsync()
        {
            // Delegate the processing of the HTTP POST to the adapter.
            // The adapter will invoke the bot.
            await this.adapter.ProcessAsync(this.Request, this.Response, this.bot);
        }
    }
}
