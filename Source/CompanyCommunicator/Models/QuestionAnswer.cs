
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
    public class QuestionAnswer
    {
        /// <summary>
        /// Gets or sets FromId.
        /// </summary>
        public string FromId { get; set; }

        /// <summary>
        /// Gets or sets Name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets AadId.
        /// </summary>
        public string AadId { get; set; }

        /// <summary>
        /// Gets or sets NotificationId.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets Questions.
        /// </summary>
        public string Questions { get; set; }

        /// <summary>
        /// Gets or sets Title.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets Author.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets answer0.
        /// </summary>
        public string answer0 { get; set; }

        /// <summary>
        /// Gets or sets answer1.
        /// </summary>
        public string answer1 { get; set; }

        /// <summary>
        /// Gets or sets answer2.
        /// </summary>
        public string answer2 { get; set; }

        /// <summary>
        /// Gets or sets answer3.
        /// </summary>
        public string answer3 { get; set; }

        /// <summary>
        /// Gets or sets answer4.
        /// </summary>
        public string answer4 { get; set; }

        /// <summary>
        /// Gets or sets answer5.
        /// </summary>
        public string answer5 { get; set; }
    }

    /// <summary>
    /// Gets or sets QuestionAnswerAdaptiveCardEntity.
    /// </summary>
    public class QuestionAnswerAdaptiveCardEntity : TableEntity
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="QuestionAnswerAdaptiveCardEntity"/> class.
        /// </summary>
        public QuestionAnswerAdaptiveCardEntity() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="QuestionAnswerAdaptiveCardEntity"/> class.
        /// </summary>
        /// <param name="notificationId">Notification Id</param>
        /// <param name="notification">Notification.</param>
        public QuestionAnswerAdaptiveCardEntity(string notificationId, string notification)
        {
            this.PartitionKey = notificationId;
            this.RowKey = notification;
        }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string NotificationId { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Questions { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string FromId { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string AadId { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Question0 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Question1 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Question2 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Question3 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Question4 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Question5 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Answer0 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Answer1 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Answer2 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Answer3 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Answer4 { get; set; }

        /// <summary>
        /// Gets or sets a FromId.
        /// </summary>
        public string Answer5 { get; set; }

    }

    public class QuestionAnswerExport
    {
        /// <summary>
        /// 
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string QuestionTitle { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Question1 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Answer1 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Question2 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Answer2 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Question3 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Answer3 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Question4 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Answer4 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Question5 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Answer5 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Question6 { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Answer6 { get; set; }

    }
}
