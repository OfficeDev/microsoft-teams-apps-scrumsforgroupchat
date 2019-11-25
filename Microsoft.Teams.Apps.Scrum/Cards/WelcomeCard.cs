// <copyright file="WelcomeCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Scrum.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.Scrum.Models;
    using Microsoft.Teams.Apps.Scrum.Properties;

    /// <summary>
    /// Implements Welcome Card.
    /// </summary>
    public class WelcomeCard
    {
        /// <summary>
        /// This method will construct the user welcome card when bot is added in personal scope.
        /// </summary>
        /// <param name="appBaseUrl">application base URL.</param>
        /// <returns>User welcome card.</returns>
        public static Attachment GetCard(string appBaseUrl)
        {
            AdaptiveCard userWelcomeCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveImage
                                    {
                                        Url = new Uri(string.Format("{0}/content/appLogo.png", appBaseUrl)),
                                        Size = AdaptiveImageSize.Large,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Large,
                                        Wrap = true,
                                        Text = Resources.WelcomeText,
                                        Weight = AdaptiveTextWeight.Bolder,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Default,
                                        Wrap = true,
                                        Text = Resources.WelcomeCardContentPart1,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = $"**{Resources.StartCarouselCardTitle}**: {Resources.StartCommandTourTextPart1}",
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = $"**{Resources.CompleteScrumTitle}**: {Resources.CompleteScrumTourText}",
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = Resources.StartCommandTourTextPart2,
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
                              Title = Resources.StartText,
                              DisplayText = Resources.StartText,
                              Text = Constants.Start,
                            },
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = userWelcomeCard,
            };
        }
    }
}
