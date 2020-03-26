// <copyright file="ScrumCards.cs" company="Microsoft">
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
    /// Implements cards.
    /// </summary>
    public class ScrumCards
    {
        /// <summary>
        /// Scrum card rendered on task module.
        /// </summary>
        /// <param name="membersId">Members id in the group.</param>
        /// <returns>card.</returns>
        public static Attachment ScrumCard(string membersId)
        {
            var variablesToValues = new Dictionary<string, string>();
            AdaptiveCard card = new AdaptiveCard("1.0")
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
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = Resources.YesterdayText,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextInput
                        {
                            Placeholder = Resources.PlaceholderText,
                            IsMultiline = true,
                            Style = AdaptiveTextInputStyle.Text,
                            Id = "yesterday",
                            MaxLength = 1000,
                        },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = Resources.TodayText,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextInput
                        {
                            Placeholder = Resources.PlaceholderText,
                            IsMultiline = true,
                            Style = AdaptiveTextInputStyle.Text,
                            Id = "today",
                            MaxLength = 1000,
                        },
                    new AdaptiveTextBlock
                        {
                            Size = AdaptiveTextSize.Medium,
                            Wrap = true,
                            Text = Resources.BlockersText,
                        },
                    new AdaptiveTextInput
                        {
                            Placeholder = Resources.PlaceholderText,
                            IsMultiline = true,
                            Style = AdaptiveTextInputStyle.Text,
                            Id = "blockers",
                            MaxLength = 1000,
                        },
                },
            };
            card.Actions.Add(
                new AdaptiveSubmitAction()
                {
                    Title = Resources.SubmitTitle,
                    Data = new AdaptiveSubmitActionData
                    {
                        MsTeams = new CardAction
                        {
                            Type = "task/submit",
                        },
                        MembersActivityIdMap = membersId,
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
        /// Card to show to user when not part of scrum.
        /// </summary>
        /// <returns>card.</returns>
        public static Attachment NoActionScrumCard()
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
                            Text = Resources.NoPartOfScrumUpdateText,
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
        /// Update scrum details to name card.
        /// </summary>
        /// <param name="name">name of group member.</param>
        /// <param name="scrumDetails">scrum details.</param>
        /// <param name="appBaseUrl">app base url.</param>
        /// <returns>card.</returns>
        public static Attachment GetUpdateCard(string name, ScrumDetails scrumDetails, string appBaseUrl)
        {
            Uri blockerImgUrl = new Uri(appBaseUrl + "/content/blocked.png");
            string updatedTimeStamp = DateTime.UtcNow.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'Z'");

            AdaptiveColumnSet columnSet = new AdaptiveColumnSet();
            var dateTimeinTextFormat = string.Format("{{{{DATE({0}, SHORT)}}}} {{{{TIME({1})}}}}", updatedTimeStamp, updatedTimeStamp);
            columnSet.Columns.Add(
                 new AdaptiveColumn
                 {
                     Width = AdaptiveColumnWidth.Auto,
                     Items = new List<AdaptiveElement>
                     {
                         new AdaptiveTextBlock
                         {
                             Weight = AdaptiveTextWeight.Bolder,
                             Text = name,
                             Wrap = true,
                             Size = AdaptiveTextSize.Default,
                         },
                         new AdaptiveTextBlock
                         {
                             Weight = AdaptiveTextWeight.Lighter,
                             Text = string.Format(Resources.UpdateScrumTimeStampText, dateTimeinTextFormat),
                             Wrap = true,
                             IsSubtle = true,
                         },
                     },
                 });

            if (!string.IsNullOrEmpty(scrumDetails.Blockers))
            {
                AdaptiveColumn column = new AdaptiveColumn();
                column.Items.Add(new AdaptiveImage
                {
                    Style = AdaptiveImageStyle.Default,
                    HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                    Url = blockerImgUrl,
                });
                columnSet.Columns.Add(column);
            }

            AdaptiveCard validationCard = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    columnSet,
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveShowCardAction()
                    {
                        Title = Resources.ShowScrumDetailsTitle,
                        Card = new AdaptiveCard("1.0")
                        {
                            Body = new List<AdaptiveElement>
                            {
                               new AdaptiveTextBlock
                               {
                                   Text = Resources.YesterdayText,
                                   Color = AdaptiveTextColor.Dark,
                                   Separator = true,
                                   IsSubtle = true,
                                   Wrap = true,
                                   Weight = AdaptiveTextWeight.Bolder,
                               },
                               new AdaptiveTextBlock
                               {
                                   Text = scrumDetails.Yesterday,
                                   Wrap = true,
                                   Weight = AdaptiveTextWeight.Lighter,
                               },
                               new AdaptiveTextBlock
                               {
                                   Text = Resources.TodayText,
                                   Color = AdaptiveTextColor.Dark,
                                   Separator = true,
                                   IsSubtle = true,
                                   Wrap = true,
                                   Weight = AdaptiveTextWeight.Bolder,
                               },
                               new AdaptiveTextBlock
                               {
                                   Text = scrumDetails.Today,
                                   Wrap = true,
                                   Weight = AdaptiveTextWeight.Lighter,
                               },
                               new AdaptiveTextBlock
                               {
                                   Text = Resources.BlockersText,
                                   Color = AdaptiveTextColor.Dark,
                                   Separator = true,
                                   IsSubtle = true,
                                   Wrap = true,
                                   Weight = AdaptiveTextWeight.Bolder,
                               },
                               new AdaptiveTextBlock
                               {
                                   Text = scrumDetails.Blockers,
                                   Wrap = true,
                                   Weight = AdaptiveTextWeight.Lighter,
                               },
                            },
                        },
                    },
                },
            };
            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = validationCard,
            };
        }

        /// <summary>
        /// validation card on task module.
        /// </summary>
        /// <param name="scrumDetails">ScrumDetails object.</param>
        /// <returns>return carad.</returns>
        public static Attachment ValidationCard(ScrumDetails scrumDetails)
        {
            string yesterdayValidationText = string.IsNullOrEmpty(scrumDetails.Yesterday) ? Resources.YesterdayValidationText : string.Empty;
            string todayValidationText = string.IsNullOrEmpty(scrumDetails.Today) ? Resources.TodayValidationText : string.Empty;

            AdaptiveCard validationCard = new AdaptiveCard("1.0")
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
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = Resources.YesterdayText,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = yesterdayValidationText,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        Color = AdaptiveTextColor.Attention,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextInput
                        {
                            IsMultiline = true,
                            Style = AdaptiveTextInputStyle.Text,
                            Id = "yesterday",
                            MaxLength = 1000,
                            Value = scrumDetails.Yesterday,
                        },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = Resources.TodayText,
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = todayValidationText,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        Color = AdaptiveTextColor.Attention,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextInput
                        {
                            IsMultiline = true,
                            Style = AdaptiveTextInputStyle.Text,
                            Id = "today",
                            MaxLength = 1000,
                            Value = scrumDetails.Today,
                        },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Wrap = true,
                                        Text = Resources.BlockersText,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveTextInput
                        {
                            IsMultiline = true,
                            Style = AdaptiveTextInputStyle.Text,
                            Id = "blockers",
                            MaxLength = 1000,
                            Value = scrumDetails.Blockers,
                        },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction()
                    {
                        Title = Resources.SubmitTitle,
                        Data = new AdaptiveSubmitActionData
                        {
                            MsTeams = new CardAction
                            {
                                Type = "task/submit",
                            },
                            MembersActivityIdMap = scrumDetails.MembersActivityIdMap,
                        },
                    },
                },
            };

            return new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = validationCard,
            };
        }
    }
}
