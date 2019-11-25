// <copyright file="ScrumEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Scrum.Models
{
    using System;
    using Microsoft.WindowsAzure.Storage.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Entity stored in table storage.
    /// </summary>
    public class ScrumEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets status of the scrum.
        /// </summary>
        [JsonProperty("ScrumStatus")]
        public bool IsScrumRunning { get; set; }

        /// <summary>
        /// Gets or sets the conversation ID of the group chat that started the scrum.
        /// </summary>
        [JsonProperty("ThreadConversationId")]
        public string ThreadConversationId { get; set; }

        /// <summary>
        /// Gets or sets the activity ID of the root scrum card.
        /// </summary>
        [JsonProperty("ScrumStartActivityId")]
        public string ScrumStartActivityId { get; set; }

        /// <summary>
        /// Gets or sets the activity ID of the root scrum card.
        /// </summary>
        [JsonProperty("TrailCardActivityId")]
        public string TrailCardActivityId { get; set; }

        /// <summary>
        /// Gets or sets the activity ID of the root scrum card.
        /// </summary>
        [JsonProperty("MembersActivityIdMap")]
        public string MembersActivityIdMap { get; set; }

        /// <summary>
        /// Gets timestamp from storage table.
        /// </summary>
        [JsonProperty("Timestamp")]
        public new DateTimeOffset Timestamp => base.Timestamp;
    }
}
