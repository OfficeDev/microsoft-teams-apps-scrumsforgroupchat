// <copyright file="ScrumDetails.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Scrum.Models
{
    /// <summary>
    /// Holds Scrum details.
    /// </summary>
    public class ScrumDetails
    {
        /// <summary>
        /// Gets or sets yesterday.
        /// </summary>
        public string Yesterday { get; set; }

        /// <summary>
        /// Gets or sets today.
        /// </summary>
        public string Today { get; set; }

        /// <summary>
        /// Gets or sets blockers.
        /// </summary>
        public string Blockers { get; set; }

        /// <summary>
        /// Gets or sets members Id in group.
        /// </summary>
        public string MembersActivityIdMap { get; set; }
    }
}
