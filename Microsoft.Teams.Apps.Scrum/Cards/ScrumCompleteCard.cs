// <copyright file="ScrumCompleteCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>
namespace Microsoft.Teams.Apps.Scrum.Cards
{
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.Scrum.Properties;

    /// <summary>
    /// Implement Scrum complete card.
    /// </summary>
    public class ScrumCompleteCard
    {
        /// <summary>
        /// Card to show when scrum is complete.
        /// </summary>
        /// <returns>card.</returns>
        public static Attachment GetScrumCompleteCard()
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
                            Text = Resources.ScrumCompleteText,
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
