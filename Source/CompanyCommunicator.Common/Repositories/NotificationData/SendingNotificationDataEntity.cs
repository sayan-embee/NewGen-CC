// <copyright file="SendingNotificationDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// Sending notification entity class.
    /// This entity holds the information about the content for a notification
    /// that is either currently being sent or was previously sent.
    /// </summary>
    public class SendingNotificationDataEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets the notification id.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets the content of the notification in serialized JSON form.
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// Gets or sets the EmailBody.
        /// </summary>
        public string EmailBody { get; set; }

        /// <summary>
        /// Gets or sets the EmailBody.
        /// </summary>
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets the EmailBody.
        /// </summary>
        public string TenantId { get; set; }

        /// <summary>
        /// Gets or sets the EmailTitle.
        /// </summary>
        public string EmailTitle { get; set; }

        /// <summary>
        /// Gets or sets the SendTypeId.
        /// </summary>
        public string SendTypeId { get; set; }

        /// <summary>
        /// Gets or sets the TemplateType.
        /// </summary>
        public string TemplateType { get; set; }

        /// <summary>
        /// Gets or sets the Title.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the Title.
        /// </summary>
        public string ImageLink { get; set; }
    }
}
