// <copyright file="TourCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Scrum.Cards
{
    using System.Collections.Generic;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.Scrum.Properties;

    /// <summary>
    /// Implements Welcome Tour Carousel card.
    /// </summary>
    public class TourCard
    {
        /// <summary>
        /// Start command carousel card.
        /// </summary>
        /// <param name="appBaseUrl">appBaseUrl.</param>
        /// <returns>card.</returns>
        public static Attachment StartCard(string appBaseUrl)
        {
            string imageUri = appBaseUrl + "/content/startScrum.png";
            HeroCard tourCarouselCard = new HeroCard()
            {
                Title = Resources.StartCarouselCardTitle,
                Text = string.Format("{0} <br /><br /> {1}", Resources.StartCommandTourTextPart1, Resources.StartCommandTourTextPart2),
                Images = new List<CardImage>()
                {
                    new CardImage(imageUri),
                },
                Buttons = new List<CardAction>()
                {
                    new CardAction
                    {
                        Type = ActionTypes.MessageBack,
                        Title = Resources.StartText,
                        DisplayText = Resources.StartText,
                        Text = Constants.Start,
                    },
                },
            };

            return tourCarouselCard.ToAttachment();
        }

        /// <summary>
        /// Implements Complete.
        /// </summary>
        /// <param name="appBaseUrl">appBaseUrl.</param>
        /// <returns>complete scum tour card.</returns>
        public static Attachment CompleteScrumCard(string appBaseUrl)
        {
            string imageUri = appBaseUrl + "/content/endScrum.png";
            HeroCard tourCarouselCard = new HeroCard()
            {
                Title = Resources.CompleteScrumTitle,
                Text = Resources.CompleteScrumTourText,
                Images = new List<CardImage>()
                {
                    new CardImage(imageUri),
                },
            };

            return tourCarouselCard.ToAttachment();
        }

        /// <summary>
        /// Create the set of cards that comprise tour carousel.
        /// </summary>
        /// <param name="appBaseUrl">The base URI where the app is hosted.</param>
        /// <returns>The cards that comprise the team tour.</returns>
        public static IEnumerable<Attachment> GetTourCards(string appBaseUrl)
        {
            return new List<Attachment>()
            {
                StartCard(appBaseUrl),
                CompleteScrumCard(appBaseUrl),
            };
        }
    }
}