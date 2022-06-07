
namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// Base QuestionAnswer model class.
    /// </summary>
    public class UserReaction
    {
        /// <summary>
        /// Gets or sets the FromId.
        /// </summary>
        public string FromId { get; set; }

        /// <summary>
        /// Gets or sets the Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the AadId.
        /// </summary>
        public string AadId { get; set; }

        /// <summary>
        /// Gets or sets the ActivityId.
        /// </summary>
        public string ActivityId { get; set; }

        /// <summary>
        /// Gets or sets the ReactionType.
        /// </summary>
        public string ReactionType { get; set; }
    }

    /// <summary>
    /// UserReactionEntity.
    /// </summary>
    public class UserReactionEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UserReactionEntity"/> class.
        /// </summary>
        public UserReactionEntity()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="UserReactionEntity"/> class.
        /// </summary>
        /// <param name="activityId">Activity Id.</param>
        /// <param name="fromId">From Id</param>
        public UserReactionEntity(string activityId, string fromId)
        {
            this.PartitionKey = activityId;
            this.RowKey = fromId;
        }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string FromId { get; set; }

        /// <summary>
        /// Gets or sets a Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets a AadId.
        /// </summary>
        public string AadId { get; set; }

        /// <summary>
        /// Gets or sets a ActivityId.
        /// </summary>
        public string ActivityId { get; set; }

        /// <summary>
        /// Gets or sets a ReactionType.
        /// </summary>
        public string ReactionType { get; set; }
    }

    /// <summary>
    /// 
    /// </summary>
    public class UserReactionExport
    {
        /// <summary>
        /// Gets or sets a LikeCount.
        /// </summary>
        public int LikeCount { get; set; }

        /// <summary>
        /// Gets or sets a HeartCount.
        /// </summary>
        public int HeartCount { get; set; }

        /// <summary>
        /// Gets or sets a LaughCount.
        /// </summary>
        public int LaughCount { get; set; }

        /// <summary>
        /// Gets or sets a SurprisedCount.
        /// </summary>
        public int SurprisedCount { get; set; }

        /// <summary>
        /// Gets or sets a SadCount.
        /// </summary>
        public int SadCount { get; set; }

        /// <summary>
        /// Gets or sets a AngryCount.
        /// </summary>
        public int AngryCount { get; set; }
    }
}
