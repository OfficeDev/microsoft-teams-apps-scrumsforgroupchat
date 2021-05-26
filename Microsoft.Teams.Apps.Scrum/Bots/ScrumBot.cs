// <copyright file="ScrumBot.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Scrum.Bots
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Rest;
    using Microsoft.Teams.Apps.Scrum.Cards;
    using Microsoft.Teams.Apps.Scrum.Common;
    using Microsoft.Teams.Apps.Scrum.Models;
    using Microsoft.Teams.Apps.Scrum.Properties;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Polly;
    using Polly.Contrib.WaitAndRetry;
    using Polly.Retry;

    /// <summary>
    /// Implements the core logic of the AskHR bot.
    /// </summary>
    public class ScrumBot : TeamsActivityHandler
    {
        /// <summary>
        /// Sets the team members cache key.
        /// </summary>
        private const string TeamMembersCacheKey = "teamMembersCacheKey";

        private static AsyncRetryPolicy retryPolicy = Policy.Handle<HttpOperationException>()
            .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 5));

        private static AsyncRetryPolicy memberRetryPolicy = Policy.Handle<HttpOperationException>(ex => ex.Response.StatusCode == HttpStatusCode.TooManyRequests)
            .WaitAndRetryAsync(Backoff.DecorrelatedJitterBackoffV2(TimeSpan.FromMilliseconds(1000), 5));

        private readonly string expectedTenantId;
        private readonly IConfiguration configuration;
        private readonly IScrumProvider scrumProvider;
        private readonly TelemetryClient telemetryClient;

        ///// <summary>
        ///// Sends logs to the Application Insights service.
        ///// </summary>
        // private readonly ILogger logger;

        /// <summary>
        /// Cache for storing authorization result.
        /// </summary>
        private readonly IMemoryCache memoryCache;

        /// <summary>
        /// Initializes a new instance of the <see cref="ScrumBot"/> class.
        /// </summary>
        /// <param name="conversationState">Conversation State.</param>
        /// <param name="configuration">Configuration.</param>
        /// <param name="scrumProvider">scrumProvider.</param>
        /// <param name="telemetryClient">Telemetry.</param>
        /// <param name="logger">Instance to send logs to the Application Insights service.</param>
        /// <param name="memoryCache">MemoryCache instance for caching authorization result.</param>
        public ScrumBot(IConfiguration configuration, IScrumProvider scrumProvider, TelemetryClient telemetryClient, IMemoryCache memoryCache)
        {
            this.scrumProvider = scrumProvider;
            this.telemetryClient = telemetryClient;
            this.configuration = configuration;
            this.expectedTenantId = configuration["TenantId"];
            this.memoryCache = memoryCache;
        }

        /// <inheritdoc/>
        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default(CancellationToken))
        {
            try
            {
                if (!this.IsActivityFromExpectedTenant(turnContext))
                {
                    this.telemetryClient.TrackTrace($"Unexpected tenant id {turnContext.Activity.Conversation.TenantId}", SeverityLevel.Warning);
                    await turnContext.SendActivityAsync(Resources.WarningTextForTenantFailure);
                    await Task.CompletedTask;
                }
                else
                {
                    bool isMemberCountUnsupported = await this.IsMemberCountGreaterThanAllowed(turnContext, cancellationToken);
                    if (isMemberCountUnsupported)
                    {
                        await turnContext.SendActivityAsync(Resources.MemberCountErrorMessage);
                    }
                    else
                    {
                        switch (turnContext.Activity.Type)
                        {
                            case ActivityTypes.Message:
                                await this.OnMessageActivityAsync(new DelegatingTurnContext<IMessageActivity>(turnContext), cancellationToken);
                                break;

                            default:
                                this.telemetryClient.TrackTrace($"Ignoring event from conversation type {turnContext.Activity.Conversation.ConversationType}");
                                await base.OnTurnAsync(turnContext, cancellationToken);
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                this.telemetryClient.TrackTrace($"Exception : {ex.Message} for conversation id : {turnContext.Activity.Conversation.Id}");
            }
        }

        /// <inheritdoc/>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                await this.SendTypingIndicatorAsync(turnContext);

                var conversationType = turnContext.Activity.Conversation.ConversationType;
                string conversationId = turnContext.Activity.Conversation.Id;

                if (string.Compare(conversationType, "groupChat", StringComparison.OrdinalIgnoreCase) == 0)
                {
                    if (turnContext.Activity.Type.Equals(ActivityTypes.Message))
                    {
                        turnContext.Activity.RemoveRecipientMention();

                        switch (turnContext.Activity.Text.Trim().ToLower())
                        {
                            case Constants.Start:
                                this.telemetryClient.TrackTrace($"scrum {conversationId} started by {turnContext.Activity.From.Id}");

                                var scrum = await this.scrumProvider.GetScrumAsync(conversationId);
                                if (scrum != null && scrum.IsScrumRunning)
                                {
                                    // check if member in scrum exists.
                                    // A user is added during a running scrum and tries to start a new scrum.
                                    var activityId = this.GetActivityIdToMatch(scrum.MembersActivityIdMap, turnContext.Activity.From.Id);
                                    if (activityId == null)
                                    {
                                        await turnContext.SendActivityAsync(string.Format(Resources.NoPartOfScrumStartText, turnContext.Activity.From.Name));
                                        this.telemetryClient.TrackTrace($"Member who is updating the scrum is not the part of scrum for : {conversationId}");
                                    }
                                    else
                                    {
                                        this.telemetryClient.TrackTrace($"Scrum is already running for conversation id {conversationId}");
                                        await turnContext.SendActivityAsync(Resources.RunningScrumMessage);
                                    }
                                }
                                else
                                {
                                    // start a new scrum
                                    this.telemetryClient.TrackTrace($"Scrum start for : {conversationId}");
                                    await this.StartScrumAsync(turnContext, cancellationToken);
                                }

                                break;

                            case Constants.TakeATour:
                                var tourCards = TourCard.GetTourCards(this.configuration["AppBaseURL"]);
                                await turnContext.SendActivityAsync(MessageFactory.Carousel(tourCards));
                                break;

                            case Constants.CompleteScrum:
                                var scrumInfo = await this.scrumProvider.GetScrumAsync(conversationId);
                                if (scrumInfo.IsScrumRunning)
                                {
                                    var activityId = this.GetActivityIdToMatch(scrumInfo.MembersActivityIdMap, turnContext.Activity.From.Id);

                                    // check if member in scrum exists.
                                    // A user is added during a running scrum and tries to complete the running scrum.
                                    if (activityId == null)
                                    {
                                        await turnContext.SendActivityAsync(string.Format(Resources.NoPartOfCompleteScrumText, turnContext.Activity.From.Name));
                                        this.telemetryClient.TrackTrace($"Member who is updating the scrum is not the part of scrum for : {conversationId}");
                                    }
                                    else
                                    {
                                        var cardId = scrumInfo.ScrumStartActivityId;
                                        var activity = MessageFactory.Attachment(ScrumCompleteCard.GetScrumCompleteCard());
                                        activity.Id = cardId;
                                        activity.Conversation = turnContext.Activity.Conversation;
                                        await turnContext.UpdateActivityAsync(activity, cancellationToken);

                                        // Update the trail card
                                        var dateString = string.Format("{{{{TIME({0})}}}}", DateTime.UtcNow.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'"));
                                        string cardTrailMessage = string.Format(Resources.ScrumCompletedByText, turnContext.Activity.From.Name, dateString);
                                        await this.UpdateTrailCard(cardTrailMessage, turnContext, cancellationToken);

                                        scrumInfo.IsScrumRunning = false;
                                        scrumInfo.ThreadConversationId = conversationId;
                                        var savedData = await this.scrumProvider.SaveOrUpdateScrumAsync(scrumInfo);

                                        if (!savedData)
                                        {
                                            await turnContext.SendActivityAsync(Resources.ErrorMessage);
                                            return;
                                        }

                                        this.telemetryClient.TrackTrace($"Scrum completed by: {turnContext.Activity.From.Name} for {conversationId}");
                                    }
                                }
                                else
                                {
                                    await turnContext.SendActivityAsync(Resources.CompleteScrumErrorText);
                                }

                                break;

                            default:
                                await turnContext.SendActivityAsync(MessageFactory.Attachment(HelpCard.GetHelpCard()), cancellationToken);
                                break;
                        }
                    }
                }
                else
                {
                    await turnContext.SendActivityAsync(Resources.ScopeErrorMessage);
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"For {turnContext.Activity.Conversation.Id} : Message Activity failed: {ex.Message}");
                this.telemetryClient.TrackException(ex);
            }
        }

        /// <summary>
        /// TaskModuleFetch.
        /// </summary>
        /// <param name="turnContext">turnContext.</param>
        /// <param name="taskModuleRequest">taskmoduleRequest.</param>
        /// <param name="cancellationToken">cancellationToken.</param>
        /// <returns>returns task.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            try
            {
                var activityFetch = (Activity)turnContext.Activity;
                if (activityFetch.Value == null)
                {
                    throw new ArgumentException("activity's value should not be null");
                }

                ScrumDetails scrumMemberDetails = JsonConvert.DeserializeObject<ScrumDetails>(JObject.Parse(activityFetch.Value.ToString())["data"].ToString());
                string membersId = scrumMemberDetails.MembersActivityIdMap;
                string activityIdval = this.GetActivityIdToMatch(membersId, turnContext.Activity.From.Id);

                // A user is added during a running scrum and tries to update his/her details.
                if (activityIdval == null)
                {
                    return new TaskModuleResponse
                    {
                        Task = new TaskModuleContinueResponse
                        {
                            Value = new TaskModuleTaskInfo()
                            {
                                Card = ScrumCards.NoActionScrumCard(),
                                Height = "small",
                                Width = "medium",
                                Title = Resources.NoActiveScrumTitle,
                            },
                        },
                    };
                }
                else
                {
                    return new TaskModuleResponse
                    {
                        Task = new TaskModuleContinueResponse
                        {
                            Value = new TaskModuleTaskInfo()
                            {
                                Card = ScrumCards.ScrumCard(membersId),
                                Height = "large",
                                Width = "medium",
                                Title = Resources.ScrumTaskModuleTitle,
                            },
                        },
                    };
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Invoke Activity failed: {ex.Message}");
                this.telemetryClient.TrackException(ex);
                return null;
            }
        }

        /// <summary>
        /// TaskModuleSubmit.
        /// </summary>
        /// <param name="turnContext">turnContext.</param>
        /// <param name="taskModuleRequest">taskmoduleRequest.</param>
        /// <param name="cancellationToken">cancellationToken.</param>
        /// <returns>returns task.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            try
            {
                var activity = (Activity)turnContext.Activity;
                if (activity.Value == null)
                {
                    throw new ArgumentException("activity's value should not be null");
                }

                ScrumDetails scrumDetails = JsonConvert.DeserializeObject<ScrumDetails>(JObject.Parse(activity.Value.ToString())["data"].ToString());
                if (string.IsNullOrEmpty(scrumDetails.Yesterday) || string.IsNullOrEmpty(scrumDetails.Today))
                {
                    return this.GetScrumValidation(scrumDetails, turnContext, cancellationToken);
                }

                string activityId = this.GetActivityIdToMatch(scrumDetails.MembersActivityIdMap, turnContext.Activity.From.Id);

                // check if member in scrum does not exists
                if (activityId == null)
                {
                    await turnContext.SendActivityAsync(string.Format(Resources.NotPartofScrumText, turnContext.Activity.From.Name));
                    return default;
                }

                activity.Id = activityId;
                activity.Conversation = turnContext.Activity.Conversation;

                var activityupdate = MessageFactory.Attachment(ScrumCards.GetUpdateCard(turnContext.Activity.From.Name, scrumDetails, this.configuration["AppBaseURL"]));
                activityupdate.Id = activityId;
                activityupdate.Conversation = turnContext.Activity.Conversation;
                await turnContext.UpdateActivityAsync(activityupdate, cancellationToken).ConfigureAwait(false);

                var scrumInfo = await this.scrumProvider.GetScrumAsync(turnContext.Activity.Conversation.Id);

                // Update the trail card.
                string cardTrailMesasge = string.Format(Resources.ScrumUpdatedByText, turnContext.Activity.From.Name, this.GetUtcTimeInAdaptiveTextFormat());
                var activityTrail = MessageFactory.Attachment(ScrumStartCards.ScrumTrailCardForStartScrum(cardTrailMesasge));
                activityTrail.Id = scrumInfo.TrailCardActivityId;
                activityTrail.Conversation = turnContext.Activity.Conversation;
                await turnContext.UpdateActivityAsync(activityTrail, cancellationToken).ConfigureAwait(false);
                return default;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Invoke Activity failed: {ex.Message}");
                this.telemetryClient.TrackException(ex);
                return null;
            }
        }

        /// <inheritdoc/>
        protected override async Task OnConversationUpdateActivityAsync(ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            this.telemetryClient.TrackTrace($"Received conversationUpdate activity");

            var activity = turnContext.Activity;
            if (activity.MembersAdded?.Count > 0)
            {
                await this.OnMembersAddedAsync(activity.MembersAdded, turnContext, cancellationToken);
            }
        }

        /// <inheritdoc/>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            foreach (var member in membersAdded)
            {
                if (member.Id == turnContext.Activity.Recipient.Id)
                {
                    this.telemetryClient.TrackTrace($"Bot added to groupchat {turnContext.Activity.Conversation.Id}");

                    var userWelcomeCardAttachment = WelcomeCard.GetCard(this.configuration["AppBaseURL"]);
                    await turnContext.SendActivityAsync(MessageFactory.Attachment(userWelcomeCardAttachment));
                }
            }
        }

        /// <summary>
        /// Method to show Name card and get members list with the update status and end scrum button.
        /// </summary>
        /// <param name="turnContext">turn Context.</param>
        /// <param name="cancellationToken">cancellationToken.</param>
        private async Task StartScrumAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            try
            {
                string cardMessage = string.Format(Resources.ScrumRequestedByText, turnContext.Activity.From.Name, this.GetUtcTimeInAdaptiveTextFormat());

                var scrumTrailActivity = MessageFactory.Attachment(ScrumStartCards.ScrumTrailCardForStartScrum(cardMessage));
                var scrumTrailActivityResponse = await turnContext.SendActivityAsync(scrumTrailActivity, cancellationToken);

                var membersActivityIdMap = await this.GetActivityIdOfMembersInScrum(turnContext, cancellationToken);
                if (membersActivityIdMap != null)
                {
                    var scrumStartActivity = MessageFactory.Attachment(ScrumStartCards.GetScrumStartCard(membersActivityIdMap));
                    var scrumStartActivityResponse = await turnContext.SendActivityAsync(scrumStartActivity, cancellationToken);
                    string membersList = JsonConvert.SerializeObject(membersActivityIdMap);
                    await this.CreateScrumAsync(scrumTrailActivityResponse.Id, scrumStartActivityResponse.Id, membersList, turnContext, cancellationToken);
                    this.telemetryClient.TrackTrace($"Scrum start details saved to table storage for: {turnContext.Activity.Conversation.Id}");
                }
                else
                {
                    this.telemetryClient.TrackTrace($"Id not mapped to members: {turnContext.Activity.Conversation.Id}");
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"Start scrum failed for {turnContext.Activity.Conversation.Id}: {ex.Message}");
                this.telemetryClient.TrackException(ex);
            }
        }

        /// <summary>
        /// Update the first trail card with user details.
        /// </summary>
        /// <param name="cardTrailMesasge">string message to shown on card.</param>
        /// <param name="turnContext">turnContext.</param>
        /// <param name="cancellationToken">cancellationToken.</param>
        /// <returns>void.</returns>
        private async Task UpdateTrailCard(string cardTrailMesasge, ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var scrumInfo = await this.scrumProvider.GetScrumAsync(turnContext.Activity.Conversation.Id);
            if (scrumInfo != null)
            {
                var activityTrail = MessageFactory.Attachment(ScrumStartCards.ScrumTrailCardForCompleteScrum(cardTrailMesasge));
                activityTrail.Id = scrumInfo.TrailCardActivityId;
                activityTrail.Conversation = turnContext.Activity.Conversation;
                await turnContext.UpdateActivityAsync(activityTrail, cancellationToken).ConfigureAwait(false);
                this.telemetryClient.TrackTrace($"Trail card updated for: {turnContext.Activity.Conversation.Id}");
            }
            else
            {
                this.telemetryClient.TrackTrace($"No data obtained from storage to update trail card for : {turnContext.Activity.Conversation.Id}");
            }
        }

        /// <summary>
        /// Checks the total members in group chat and returns true if members are more than allowed limit.
        /// </summary>
        /// <param name="turnContext">turn Context.</param>
        /// <param name="cancellationToken">cancellation Token.</param>
        /// <returns>true if members are more than allowed limit.</returns>
        private async Task<bool> IsMemberCountGreaterThanAllowed(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            bool isCacheEntryExists = this.memoryCache.TryGetValue(TeamMembersCacheKey, out List<TeamsChannelAccount> members);
            if (!isCacheEntryExists)
            {
                members = await this.GetTeamMembersAsync(turnContext, cancellationToken);
                this.memoryCache.Set(TeamMembersCacheKey, members, TimeSpan.FromDays(3));
                this.telemetryClient.TrackTrace($"Total Members in group chat is :{members.Count()}");
            }

            return members.Count() > Constants.MaxAllowedMembers;
        }

        private async Task<List<TeamsChannelAccount>> GetTeamMembersAsync(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            List<TeamsChannelAccount> members = new List<TeamsChannelAccount>();
            string continuationToken = null;
            do
            {
                var currentPage = await TeamsInfo.GetPagedMembersAsync(turnContext, pageSize: 500, continuationToken, cancellationToken);
                continuationToken = currentPage.ContinuationToken;
                members.AddRange(currentPage.Members);
            }
            while (continuationToken != null);
            return members;
        }

        /// <summary>
        ///  Create a new scrum from the input.
        /// </summary>
        /// <param name="trailCardId">activityId of trail card.</param>
        /// <param name="scrumCardId">activityId of scrum card.</param>
        /// <param name="membersList">JSON serialized member and activity mapping.</param>
        /// <param name="turnContext">turnContext.</param>
        /// <param name="cancellationToken">cancellationToken.</param>
        /// <returns>void.</returns>
        private async Task CreateScrumAsync(string trailCardId, string scrumCardId, string membersList, ITurnContext turnContext, CancellationToken cancellationToken)
        {
            string conversationId = turnContext.Activity.Conversation.Id;
            try
            {
                ScrumEntity scrumEntity = new ScrumEntity
                {
                    ThreadConversationId = conversationId,
                    ScrumStartActivityId = scrumCardId,
                    IsScrumRunning = true,
                    MembersActivityIdMap = membersList,
                    TrailCardActivityId = trailCardId,
                };
                var savedData = await this.scrumProvider.SaveOrUpdateScrumAsync(scrumEntity);
                if (!savedData)
                {
                    await turnContext.SendActivityAsync(Resources.ErrorMessage);
                }
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackTrace($"For {conversationId} : Saving data to table storage failed.: {ex.Message}");
                await turnContext.SendActivityAsync(Resources.ErrorMessage);
                this.telemetryClient.TrackException(ex);
            }
        }

        /// <summary>
        /// Method to show Name card and get members list.
        /// </summary>
        /// <param name="turnContext">turn Context.</param>
        /// <param name="cancellationToken">cancellationToken.</param>
        /// <returns>member ids.</returns>
        private async Task<Dictionary<string, string>> GetActivityIdOfMembersInScrum(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var membersActivityIdMap = new Dictionary<string, string>();
            bool isCacheEntryExists = this.memoryCache.TryGetValue(TeamMembersCacheKey, out List<TeamsChannelAccount> members);
            if (!isCacheEntryExists)
            {
                members = await this.GetTeamMembersAsync(turnContext, cancellationToken);
                this.memoryCache.Set(TeamMembersCacheKey, members, TimeSpan.FromDays(1));
                this.telemetryClient.TrackTrace($"Total Members in group chat is :{members.Count()}");
            }

            foreach (var member in members)
            {
                var mentionActivity = MessageFactory.Attachment(ScrumStartCards.GetNameCard(member.Name));
                mentionActivity.Entities = new List<Entity>();
                var mentionedEntity = member;
                string mentionEntityText = string.Format("<at>{0}</at>", mentionedEntity.Name);

                mentionActivity.Entities.Add(new Mention
                {
                    Text = mentionEntityText,
                    Mentioned = mentionedEntity,
                });

                // https://github.com/Polly-Contrib/Polly.Contrib.WaitAndRetry/tree/DecorrelatedJitterV2Explorer
                await retryPolicy.ExecuteAsync(async () =>
                {
                    var response = await turnContext.SendActivityAsync(mentionActivity, cancellationToken).ConfigureAwait(false);
                    membersActivityIdMap[member.Id] = response.Id;
                });
            }

            return membersActivityIdMap;
        }

        /// <summary>
        /// Get card for validations used in task module.
        /// </summary>
        /// <param name="scrum">ScrumDetails object.</param>
        /// <param name="turnContext">Turn context.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>TaskModuleResponse.</returns>
        private TaskModuleResponse GetScrumValidation(ScrumDetails scrum, ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Value = new TaskModuleTaskInfo()
                    {
                        Card = ScrumCards.ValidationCard(scrum),
                        Height = "large",
                        Width = "medium",
                        Title = Resources.ScrumTaskModuleTitle,
                    },
                },
            };
        }

        /// <summary>
        /// Get activity Id is being used to check if member in scrum exists.
        /// </summary>
        /// <param name="membersId"> membersId.</param>
        /// <param name="activityFromId"> activityFromId.</param>
        /// <returns>activityId.</returns>
        private string GetActivityIdToMatch(string membersId, string activityFromId)
        {
            Dictionary<string, string> membersDictionary = JsonConvert.DeserializeObject<Dictionary<string, string>>(membersId);
            return membersDictionary.TryGetValue(activityFromId, out string activityId) ? activityId : string.Empty;
        }

        /// <summary>
        /// Verify if the tenant Id in the message is the same tenant Id used when application was configured.
        /// </summary>
        /// <param name="turnContext">Turn context.</param>
        /// <returns>True if context is from expected tenant else false.</returns>
        private bool IsActivityFromExpectedTenant(ITurnContext turnContext)
        {
            return turnContext.Activity.Conversation.TenantId == this.expectedTenantId;
        }

        /// <summary>
        /// Convert the date time to adaptive text format for handling locale
        /// https://docs.microsoft.com/en-us/adaptive-cards/authoring-cards/text-features#datetime-example.
        /// </summary>
        /// <returns>{{{{TIME(datetime in t-z format)}}}}.</returns>
        private string GetUtcTimeInAdaptiveTextFormat()
        {
            return string.Format("{{{{TIME({0})}}}}", DateTime.UtcNow.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'"));
        }

        /// <summary>
        /// Send typing indicator to the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task SendTypingIndicatorAsync(ITurnContext turnContext)
        {
            try
            {
                var typingActivity = turnContext.Activity.CreateReply();
                typingActivity.Type = ActivityTypes.Typing;
                await turnContext.SendActivityAsync(typingActivity);
            }
            catch (Exception ex)
            {
                // Do not fail on errors sending the typing indicator
                this.telemetryClient.TrackTrace($"Failed to send a typing indicator: {ex.Message}", SeverityLevel.Warning);
                this.telemetryClient.TrackException(ex);
            }
        }
    }
}
