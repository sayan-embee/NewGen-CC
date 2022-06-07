namespace Microsoft.Teams.Apps.CompanyCommunicator.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Azure.Documents;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Cloud storage helper.
    /// </summary>
    public class CloudStorageHelper
    {
        /// <summary>
        /// Merge Adaptive card data.
        /// </summary>
        /// <param name="_formData">Question answer form data.</param>
        /// <returns>json string.</returns>
        public async Task<string> MergeAdaptiveCardData(QuestionAnswer _formData)
        {
            var configuration = new ConfigurationBuilder()
               .SetBasePath(Directory.GetCurrentDirectory())
               .AddJsonFile("appsettings.json")
               .Build();
            var storageConnectionString = configuration.GetSection("StorageAccountConnectionString").Value.ToString();
            string body = "{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"" + _formData.Title + "\"," +
                      "\"size\": \"ExtraLarge\"," +
                      "\"wrap\": true," +
                      "\"weight\": \"Bolder\"" +
                    "}," +
                    "{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"" + _formData.Author + "\"," +
                      "\"size\": \"Small\"," +
                      "\"wrap\": true," +
                      "\"weight\": \"Lighter\"" +
                    "},";
            string Question0 = "", Question1 = "", Question2 = "", Question3 = "", Question4 = "", Question5 = "";
            if (_formData.Questions != "")
            {
                string[] questionlist = _formData.Questions.Split("||");
                if (questionlist.Length > 0)
                {
                    Question0 = questionlist[0];
                    body += "{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"1."+ Question0 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}," +
                    "{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"Ans: "+_formData.answer0+"\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}";
                }
                if (questionlist.Length > 1)
                {
                    Question1 = questionlist[1];
                    body += ",{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"2." + Question1 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}," +
                    "{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"Ans: " + _formData.answer1 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}";
                }
                if (questionlist.Length > 2)
                {
                    Question2 = questionlist[2];
                    body += ",{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"3." + Question2 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}," +
                    "{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"Ans: " + _formData.answer2 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}";
                }
                if (questionlist.Length > 3)
                {
                    Question3 = questionlist[3];
                    body += ",{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"4." + Question3 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}," +
                    "{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"Ans: " + _formData.answer3 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}";
                }
                if (questionlist.Length > 4)
                {
                    Question4 = questionlist[4];
                    body += ",{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"5." + Question4 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}," +
                    "{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"Ans: " + _formData.answer5 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}";
                }
                if (questionlist.Length > 5)
                {
                    body += ",{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"6." + Question5 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}," +
                    "{" +
                      "\"type\": \"TextBlock\"," +
                      "\"text\": \"Ans: " + _formData.answer5 + "\"," +
                      "\"size\": \"Medium\"," +
                      "\"wrap\": true," +
                      "\"horizontalAlignment\": \"Left\"" +
                    "}";
                }
            }
            
            QuestionAnswerAdaptiveCardEntity adaptiveCard = new QuestionAnswerAdaptiveCardEntity(_formData.NotificationId, _formData.Name)
            {
                NotificationId = _formData.NotificationId,
                Question0 = Question0,
                Question1 = Question1,
                Question2 = Question2,
                Question3 = Question3,
                Question4 = Question4,
                Question5 = Question5,
                Answer0 = _formData.answer0,
                Answer1 = _formData.answer1,
                Answer2 = _formData.answer2,
                Answer3 = _formData.answer3,
                Answer4 = _formData.answer4,
                Answer5 = _formData.answer5,
                Title =_formData.Title,
                Author=_formData.Author,
            };
            var tableName = "QuestionAnswer";
            CloudStorageAccount storageAccount;
            storageAccount = CloudStorageAccount.Parse(storageConnectionString);

            CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
            CloudTable table = tableClient.GetTableReference(tableName);
            TableOperation insertOnMergeOperation = TableOperation.InsertOrMerge(adaptiveCard);
            TableResult result = await table.ExecuteAsync(insertOnMergeOperation);
            return ReturnAdaptiveCardJSON(body);
        }

        public string ReturnAdaptiveCardJSON(string body)
        {
            string JsonString= "{" +
              "\"type\": \"AdaptiveCard\"," +
              "\"body\": [" + body +
              "]," +
              "\"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\"," +
              "\"version\": \"1.2\"" +
            "}";
            return JsonString;
        }

        public async Task<string> MergeUserReactionData(UserReaction _formData)
        {
            var configuration = new ConfigurationBuilder()
               .SetBasePath(Directory.GetCurrentDirectory())
               .AddJsonFile("appsettings.json")
               .Build();
            var storageConnectionString = configuration.GetSection("StorageAccountConnectionString").Value.ToString();
            var tableName = "UserReaction";
            CloudStorageAccount storageAccount;
            storageAccount = CloudStorageAccount.Parse(storageConnectionString);
            UserReactionEntity entity = new UserReactionEntity(_formData.ActivityId, _formData.AadId)
            {
                ActivityId = _formData.ActivityId,
                FromId=_formData.FromId,
                Name=_formData.Name,
                AadId=_formData.AadId,
                ReactionType=_formData.ReactionType,
            };
            CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
            CloudTable table = tableClient.GetTableReference(tableName);
            TableOperation insertOnMergeOperation = TableOperation.InsertOrMerge(entity);
            TableResult result = await table.ExecuteAsync(insertOnMergeOperation);
            return "OK";
        }

        public async Task<IEnumerable<QuestionAnswerExport>> GetSurveryList(string NotificationId)
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();
            var storageConnectionString = configuration.GetSection("StorageAccountConnectionString").Value.ToString();
            var tableName = "QuestionAnswer";
            CloudStorageAccount storageAccount;
            storageAccount = CloudStorageAccount.Parse(storageConnectionString);

            CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
            CloudTable table = tableClient.GetTableReference(tableName);
            var entities = table.ExecuteQuery(new TableQuery<QuestionAnswerAdaptiveCardEntity>()).ToList();
            var dataentity = from s in entities
                       where s.PartitionKey == NotificationId
                       select new QuestionAnswerExport
                       {
                           QuestionTitle = s.Title,
                           Author = s.Author,
                           Name = s.RowKey,
                           Question1 = s.Question0,
                           Answer1 = Convert.ToString(s.Answer0),
                           Question2 = s.Question1,
                           Answer2 = Convert.ToString(s.Answer1),
                           Question3 = s.Question2,
                           Answer3 = Convert.ToString(s.Answer2),
                           Question4 = s.Question3,
                           Answer4 = Convert.ToString(s.Answer3),
                           Question5 = s.Question4,
                           Answer5 = Convert.ToString(s.Answer4),
                           Question6 = s.Question5,
                           Answer6 = Convert.ToString(s.Answer5),
                       };
            return dataentity;
        }

        /// <summary>
        /// Get reaction list.
        /// </summary>
        /// <param name="notificationId">Notification Id.</param>
        /// <returns>Reaction list data.</returns>
        public async Task<UserReactionExport> GetReactionList(string notificationId)
        {
            int reactionLike = 0;
            int reactionHeart = 0;
            int reactionLaugh = 0;
            int reactionSurprised = 0;
            int reactionSad = 0;
            int reactionAngry = 0;
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json")
                .Build();
            var storageConnectionString = configuration.GetSection("StorageAccountConnectionString").Value.ToString();
            var tableName = "SentNotificationData";
            CloudStorageAccount storageAccount;
            storageAccount = CloudStorageAccount.Parse(storageConnectionString);

            CloudTableClient tableClient = storageAccount.CreateCloudTableClient(new TableClientConfiguration());
            CloudTable table = tableClient.GetTableReference(tableName);
            var entities = table.ExecuteQuery(new TableQuery<SentNotificationDataEntity>()).ToList().Where(x => x.PartitionKey == notificationId);
            foreach (var entity in entities)
            {
                // int like = 0, heart = 0,laugh=0,surprised=0,
                int sad = 0, angry = 0;
                CloudTable table1 = tableClient.GetTableReference("UserReaction");
                var subentities = table1.ExecuteQuery(new TableQuery<UserReactionEntity>()).ToList().Where(x => x.ActivityId == entity.ActivityId);
                foreach (var subentity in subentities)
                {
                    if (subentity.ReactionType.ToLower() == "like")
                    {
                        reactionLike = reactionLike + 1;
                    }
                    else if (subentity.ReactionType.ToLower() == "heart")
                    {
                        reactionHeart = reactionHeart + 1;
                    }
                    else if (subentity.ReactionType.ToLower() == "laugh")
                    {
                        reactionLaugh = reactionLaugh + 1;
                    }
                    else if (subentity.ReactionType.ToLower() == "surprised")
                    {
                        reactionSurprised = reactionSurprised + 1;
                    }
                    else if (subentity.ReactionType.ToLower() == "sad")
                    {
                        reactionSad = reactionSad + sad;
                    }
                    else if (subentity.ReactionType.ToLower() == "angry")
                    {
                        reactionAngry = reactionAngry + angry;
                    }
                }

                // ReactionLike = ReactionLike + like;
                // ReactionHeart = ReactionHeart + heart;
                // ReactionLaugh = ReactionHeart + laugh;
                // ReactionSurprised = ReactionHeart + surprised;
                // ReactionSad = ReactionHeart + sad;
                // ReactionAngry = ReactionHeart + angry;
            }

            UserReactionExport exportData = new UserReactionExport();
            exportData.LikeCount = reactionLike;
            exportData.HeartCount = reactionHeart;
            exportData.LaughCount = reactionLaugh;
            exportData.SurprisedCount = reactionSurprised;
            exportData.SadCount = reactionSad;
            exportData.AngryCount = reactionAngry;

            await Task.Delay(0);

            return exportData;
        }
    }
}
