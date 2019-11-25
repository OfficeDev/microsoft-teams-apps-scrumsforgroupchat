// <copyright file="IScrumProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Scrum.Common
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.Scrum.Models;

    /// <summary>
    /// Interface to implement Table storage methods.
    /// </summary>
    public interface IScrumProvider
    {
        /// <summary>
        /// Save or update scrum entity.
        /// </summary>
        /// <param name="scrum">Scrum received from bot based on which appropriate row will replaced or inserted in table storage.</param>
        /// <returns><see cref="Task"/> that resolves successfully if the data was saved successfully.</returns>
        Task<bool> SaveOrUpdateScrumAsync(ScrumEntity scrum);

        /// <summary>
        /// Get already saved entity detail from storage table.
        /// </summary>
        /// <param name="conversationId">scrum id received from bot based on which appropriate row data will be fetched.</param>
        /// <returns><see cref="Task"/> Already saved entity detail.</returns>
        Task<ScrumEntity> GetScrumAsync(string conversationId);
    }
}
