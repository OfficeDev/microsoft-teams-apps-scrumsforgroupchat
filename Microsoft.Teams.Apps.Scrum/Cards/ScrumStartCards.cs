// <copyright file="ScrumStartCards.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Scrum.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.Scrum.Models;
    using Microsoft.Teams.Apps.Scrum.Properties;
    using Newtonsoft.Json;

    /// <summary>
    /// Implement Start Scrum Card.
    /// </summary>
    public class ScrumStartCards
    {
        /// <summary>
        /// Start scrum card.
        /// </summary>
        /// <param name="membersId">group members id.</param>
        /// <returns>card.</returns>
        public static Attachment GetScrumStartCard(Dictionary<string, string> membersId)
        {
            AdaptiveCard card = new AdaptiveCard("1.0");

            card.Actions.Add(
                new AdaptiveSubmitAction
                {
                    Title = Resources.UpdateScrumTitle,
                    Data = new AdaptiveSubmitActionData
                    {
                        MsTeams = new TaskModuleAction(Resources.UpdateScrumTitle, new ScrumDetails { MembersActivityIdMap = JsonConvert.SerializeObject(membersId, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore }) }),
                    },
                });
            card.Actions.Add(
            new AdaptiveSubmitAction
            {
                Title = Resources.CompleteScrumTitle,
                Data = new AdaptiveSubmitActionData
                {
                    MsTeams = new CardAction
                    {
                        Type = ActionTypes.MessageBack,
                        Text = Resources.CompleteScrumText,
                        Value = membersId,
                    },
                },
            });
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Tag each user in the group.
        /// </summary>
        /// <param name="name">name of member.</param>
        /// <returns>card.</returns>
        public static Attachment GetNameCard(string name)
        {
            AdaptiveCard card = new AdaptiveCard("1.0");
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Medium,
                            Wrap = true,
                            Text = name,
                        },
                    },
            };
            card.Body.Add(container);
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Shows Scrum Trail details.
        /// </summary>
        /// <param name="cardMessage">Message to show on card.</param>
        /// <returns>card.</returns>
        public static Attachment ScrumTrailCardForCompleteScrum(string cardMessage)
        {
            AdaptiveCard card = new AdaptiveCard("1.0");
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Medium,
                            Wrap = true,
                            Text = cardMessage,
                        },
                    },
            };
            card.Body.Add(container);
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Shows Scrum Trail details.
        /// </summary>
        /// <param name="cardMessage">Message to show on card.</param>
        /// <returns>scrum trail card.</returns>
        public static Attachment ScrumTrailCardForStartScrum(string cardMessage)
        {
            AdaptiveCard card = new AdaptiveCard("1.0");
            var container = new AdaptiveContainer()
            {
                Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Medium,
                            Wrap = true,
                            Text = cardMessage,
                        },
                    },
            };
            card.Body.Add(container);
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };
            return adaptiveCardAttachment;
        }
    }
}
