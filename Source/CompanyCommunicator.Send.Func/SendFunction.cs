// <copyright file="SendFunction.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Send.Func
{
    using System;
    using System.Threading.Tasks;
    using System.Web;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Rest;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.SendQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams;
    using Microsoft.Teams.Apps.CompanyCommunicator.Send.Func.Services;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Blob;
    using Newtonsoft.Json;

    /// <summary>
    /// Azure Function App triggered by messages from a Service Bus queue
    /// Used for sending messages from the bot.
    /// </summary>
    public class SendFunction
    {
        /// <summary>
        /// This is set to 10 because the default maximum delivery count from the service bus
        /// message queue before the service bus will automatically put the message in the Dead Letter
        /// Queue is 10.
        /// </summary>
        private static readonly int MaxDeliveryCountForDeadLetter = 10;
        private static readonly string AdaptiveCardContentType = "application/vnd.microsoft.card.adaptive";

        private readonly int maxNumberOfAttempts;
        private readonly double sendRetryDelayNumberOfSeconds;
        private readonly INotificationService notificationService;
        private readonly ISendingNotificationDataRepository notificationRepo;
        private readonly IMessageService messageService;
        private readonly ISendQueue sendQueue;
        private readonly IStringLocalizer<Strings> localizer;
        private readonly IUsersService usersService;
        private readonly IOptions<RepositoryOptions> optionsRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="SendFunction"/> class.
        /// </summary>
        /// <param name="options">Send function options.</param>
        /// <param name="notificationService">The service to precheck and determine if the queue message should be processed.</param>
        /// <param name="messageService">Message service.</param>
        /// <param name="notificationRepo">Notification repository.</param>
        /// <param name="sendQueue">The send queue.</param>
        /// <param name="localizer">Localization service.</param>
        /// <param name="usersService">User service.</param>
        /// <param name="optionsRepository">Storage Account Options.</param>
        public SendFunction(
            IOptions<SendFunctionOptions> options,
            INotificationService notificationService,
            IMessageService messageService,
            ISendingNotificationDataRepository notificationRepo,
            ISendQueue sendQueue,
            IStringLocalizer<Strings> localizer,
            IUsersService usersService,
            IOptions<RepositoryOptions> optionsRepository
            )
        {
            if (options is null)
            {
                throw new ArgumentNullException(nameof(options));
            }

            if (optionsRepository is null)
            {
                throw new ArgumentNullException(nameof(optionsRepository));
            }

            this.maxNumberOfAttempts = options.Value.MaxNumberOfAttempts;
            this.sendRetryDelayNumberOfSeconds = options.Value.SendRetryDelayNumberOfSeconds;

            this.notificationService = notificationService ?? throw new ArgumentNullException(nameof(notificationService));
            this.messageService = messageService ?? throw new ArgumentNullException(nameof(messageService));
            this.notificationRepo = notificationRepo ?? throw new ArgumentNullException(nameof(notificationRepo));
            this.sendQueue = sendQueue ?? throw new ArgumentNullException(nameof(sendQueue));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
            this.usersService = usersService ?? throw new ArgumentNullException(nameof(usersService));

            this.optionsRepository = optionsRepository;
        }

        /// <summary>
        /// Azure Function App triggered by messages from a Service Bus queue
        /// Used for sending messages from the bot.
        /// </summary>
        /// <param name="myQueueItem">The Service Bus queue item.</param>
        /// <param name="deliveryCount">The deliver count.</param>
        /// <param name="enqueuedTimeUtc">The enqueued time.</param>
        /// <param name="messageId">The message ID.</param>
        /// <param name="log">The logger.</param>
        /// <param name="context">The execution context.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [FunctionName("SendMessageFunction")]
        public async Task Run(
            [ServiceBusTrigger(
                SendQueue.QueueName,
                Connection = SendQueue.ServiceBusConnectionConfigurationKey)]
            string myQueueItem,
            int deliveryCount,
            DateTime enqueuedTimeUtc,
            string messageId,
            ILogger log,
            ExecutionContext context)
        {
            log.LogInformation($"C# ServiceBus queue trigger function processed message: {myQueueItem}");

            var messageContent = JsonConvert.DeserializeObject<SendQueueMessageContent>(myQueueItem);

            try
            {
                // Check if notification is pending.
                var isPending = await this.notificationService.IsPendingNotification(messageContent);
                if (!isPending)
                {
                    // Notification is either already sent or failed and shouldn't be retried.
                    return;
                }

                // Check if conversationId is set to send message.
                if (string.IsNullOrWhiteSpace(messageContent.GetConversationId()))
                {
                    await this.notificationService.UpdateSentNotification(
                        notificationId: messageContent.NotificationId,
                        recipientId: messageContent.RecipientData.RecipientId,
                        totalNumberOfSendThrottles: 0,
                        statusCode: SentNotificationDataEntity.FinalFaultedStatusCode,
                        allSendStatusCodes: $"{SentNotificationDataEntity.FinalFaultedStatusCode},",
                        errorMessage: this.localizer.GetString("AppNotInstalled"),
                        activityId: "");
                    return;
                }

                // Check if the system is throttled.
                var isThrottled = await this.notificationService.IsSendNotificationThrottled();
                if (isThrottled)
                {
                    // Re-Queue with delay.
                    await this.sendQueue.SendDelayedAsync(messageContent, this.sendRetryDelayNumberOfSeconds);
                    return;
                }

                // Send message.
                var messageActivity = await this.GetMessageActivity(messageContent);

                // If the message is important, we need to notify the user in Teams
                if (messageContent.IsImportant)
                {
                    messageActivity.TeamsNotifyUser();
                }

                var response = await this.messageService.SendMessageAsync(
                    message: messageActivity,
                    serviceUrl: messageContent.GetServiceUrl(),
                    conversationId: messageContent.GetConversationId(),
                    maxAttempts: this.maxNumberOfAttempts,
                    logger: log);

                // Process response.
                await this.ProcessResponseAsync(messageContent, response, log);

                await this.SendEmailTeamplate(messageContent, log);

            }
            catch (InvalidOperationException exception)
            {
                // Bad message shouldn't be requeued.
                log.LogError(exception, $"InvalidOperationException thrown. Error message: {exception.Message}");
            }
            catch (Exception e)
            {
                var errorMessage = $"{e.GetType()}: {e.Message}";
                log.LogError(e, $"Failed to send message. ErrorMessage: {errorMessage}");

                // Update status code depending on delivery count.
                var statusCode = SentNotificationDataEntity.FaultedAndRetryingStatusCode;
                if (deliveryCount >= SendFunction.MaxDeliveryCountForDeadLetter)
                {
                    // Max deliveries attempted. No further retries.
                    statusCode = SentNotificationDataEntity.FinalFaultedStatusCode;
                }

                // Update sent notification table.
                await this.notificationService.UpdateSentNotification(
                    notificationId: messageContent.NotificationId,
                    recipientId: messageContent.RecipientData.RecipientId,
                    totalNumberOfSendThrottles: 0,
                    statusCode: statusCode,
                    allSendStatusCodes: $"{statusCode},",
                    errorMessage: errorMessage,
                    activityId: "");

                throw;
            }
        }

        /// <summary>
        /// Process send notification response.
        /// </summary>
        /// <param name="messageContent">Message content.</param>
        /// <param name="sendMessageResponse">Send notification response.</param>
        /// <param name="log">Logger.</param>
        private async Task ProcessResponseAsync(
            SendQueueMessageContent messageContent,
            SendMessageResponse sendMessageResponse,
            ILogger log)
        {
            if (sendMessageResponse.ResultType == SendMessageResult.Succeeded)
            {
                log.LogInformation($"Successfully sent the message." +
                    $"\nRecipient Id: {messageContent.RecipientData.RecipientId}");
            }
            else
            {
                log.LogError($"Failed to send message." +
                    $"\nRecipient Id: {messageContent.RecipientData.RecipientId}" +
                    $"\nResult: {sendMessageResponse.ResultType}." +
                    $"\nErrorMessage: {sendMessageResponse.ErrorMessage}.");
            }

            await this.notificationService.UpdateSentNotification(
                    notificationId: messageContent.NotificationId,
                    recipientId: messageContent.RecipientData.RecipientId,
                    totalNumberOfSendThrottles: sendMessageResponse.TotalNumberOfSendThrottles,
                    statusCode: sendMessageResponse.StatusCode,
                    allSendStatusCodes: sendMessageResponse.AllSendStatusCodes,
                    errorMessage: sendMessageResponse.ErrorMessage,
                    activityId: sendMessageResponse.ActivityId);

            // Throttled
            if (sendMessageResponse.ResultType == SendMessageResult.Throttled)
            {
                // Set send function throttled.
                await this.notificationService.SetSendNotificationThrottled(this.sendRetryDelayNumberOfSeconds);

                // Requeue.
                await this.sendQueue.SendDelayedAsync(messageContent, this.sendRetryDelayNumberOfSeconds);
                return;
            }
        }

        private async Task<IMessageActivity> GetMessageActivity(SendQueueMessageContent message)
        {
            var notification = await this.notificationRepo.GetAsync(
                NotificationDataTableNames.SendingNotificationsPartition,
                message.NotificationId);

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCardContentType,
                Content = JsonConvert.DeserializeObject(notification.Content),
            };

            return MessageFactory.Attachment(adaptiveCardAttachment);
        }

        private async Task<bool> SendEmailTeamplate(SendQueueMessageContent messageContent, ILogger log)
        {
            // Only email template, and primary tenant
            try
            {
                var notification = await this.notificationRepo.GetAsync(
               NotificationDataTableNames.SendingNotificationsPartition,
               messageContent.NotificationId);
                if (notification.TemplateType.ToUpper() == "UPLOAD EMAIL TEMPLATE" && (notification.SendTypeId == "2" || notification.SendTypeId == "3" || notification.SendTypeId == "4"))
                {

                    var user = await this.usersService.GetUserAsync(messageContent.RecipientData.UserData.AadId);
                    log.LogInformation($"Mail Sending User Name :{user.DisplayName} -> UPN : {user.UserPrincipalName} -> User Email : {user.Mail}");

                    var htmlContent=await this.GetEmailContentFromHtml(notification, log);
                    if (!string.IsNullOrEmpty(htmlContent))
                    {
                        // var sentResult = await this.usersService.SendMailToUserAsync(user.UserPrincipalName, "Company Communicator V3 Notification", user.Mail, notification.Content, true);
                        var sentResult = await this.usersService.SendMailToUserAsync(notification.CreatedBy, notification.Title, user.Mail, htmlContent, true);

                        log.LogInformation($"Mail sent status :{sentResult} ->{notification.NotificationId}");
                    }
                    else
                    {
                        log.LogInformation($"Email content is empty :{notification.NotificationId}");
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                // Bad message shouldn't be requeued.
                log.LogError(e, $"SendMailToUserAsync thrown. Error message: {e.Message}");
                return false;
            }
        }

        private async Task<string> GetEmailContentFromHtml(SendingNotificationDataEntity notification, ILogger log)
        {
            string contents = string.Empty;
            try
            {
                if (!string.IsNullOrEmpty(notification.ImageLink))
                {
                    
                    var imageLinkArray=notification.ImageLink.Split("/");
                    var fileName= imageLinkArray[imageLinkArray.Length - 1];
                    if (!string.IsNullOrEmpty(fileName))
                    {
                        // Decode the encoded string.
                        var decodedFileName=HttpUtility.UrlDecode(fileName);
                        CloudStorageAccount storageAccount = Microsoft.WindowsAzure.Storage.CloudStorageAccount.Parse(this.optionsRepository.Value.StorageAccountConnectionString);

                        // Connect to the blob storage
                        CloudBlobClient serviceClient = storageAccount.CreateCloudBlobClient();

                        // Connect to the blob container
                        CloudBlobContainer container = serviceClient.GetContainerReference("pdffiles");

                        // Connect to the blob file
                        CloudBlockBlob blob = container.GetBlockBlobReference(decodedFileName);

                        // Get the blob file as text
                        contents = await blob.DownloadTextAsync();
                        log.LogInformation($"Content downloaded from html for notification  :{notification}");
                    }
                }
            }
            catch (Exception e)
            {
                // Bad message shouldn't be requeued.
                log.LogError(e, $"SendMailToUserAsync thrown. Error message: {e.Message}");
            }

            return contents;
        }
    }
}
