// <copyright file="HelpCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Scrum.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.Scrum.Models;
    using Microsoft.Teams.Apps.Scrum.Properties;

    /// <summary>
    /// Implements Help Card.
    /// </summary>
    public class HelpCard
    {
        /// <summary>
        /// Get the help card.
        /// </summary>
        /// <returns>help card.</returns>
        public static Attachment GetHelpCard()
        {
            AdaptiveCard helpCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = Resources.ScrumHelpMessage,
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = Resources.StartText,
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new CardAction
                            {
                              Type = ActionTypes.MessageBack,
                              DisplayText = Resources.StartText,
                              Text = Constants.Start,
                            },
                        },
                    },
                    new AdaptiveSubmitAction
                    {
                        Title = Resources.CompleteScrumTitle,
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new CardAction
                            {
                              Type = ActionTypes.MessageBack,
                              DisplayText = Resources.CompleteScrumText,
                              Text = Constants.CompleteScrum,
                            },
                        },
                    },
                    new AdaptiveSubmitAction
                    {
                        Title = Resources.TakeATourButtonText,
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new CardAction
                            {
                              Type = ActionTypes.MessageBack,
                              DisplayText = Resources.TakeATourButtonText,
                              Text = Constants.TakeATour,
                            },
                        },
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = helpCard,
            };
        }
    }
}