using Microsoft.Bot.Connector;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using AdaptiveCards;
using Microsoft.Bot.Builder;
using Newtonsoft.Json;
using Xunit;
using Microsoft.Extensions.Configuration;

namespace TeamsBotTest
{
    public class TeamsBotIntegrationTests
    {
        private static readonly BotTestSettings Settings = LoadSettings();

        /// <summary>
        /// Verifies the bot can authenticate and send a plain text message to an existing Teams conversation.
        /// </summary>
        [Fact()]
        public async Task SendToExistingConversation_Test()
        {
            var connectorClient = new ConnectorClient(new Uri(Settings.ServiceUrl), new MicrosoftAppCredentials(Settings.AppId, Settings.AppPassword));

            // Create a simple text message
            var message = new Activity(ActivityTypes.Message)
            {
                Text = $"Hello from Teams Bot! This message follows the TeamsBotUserService pattern. By {nameof(SendToExistingConversation_Test)}",
                From = new ChannelAccount { Id = Settings.AppId },
                Recipient = new ChannelAccount { Id = Settings.RecipientId },
                Conversation = new ConversationAccount
                {
                    Id = Settings.ConversationId,
                    TenantId = Settings.TenantId
                },
                ChannelId = "msteams"
            };

            var response = await connectorClient.Conversations.SendToConversationAsync(message);
        }

        /// <summary>
        /// Verifies the bot can proactively create a new Teams conversation and post a message in it.
        /// </summary>
        [Fact()]
        public async Task CreateProactiveConversation_Test()
        {
            var connectorClient = new ConnectorClient(new Uri(Settings.ServiceUrl), new MicrosoftAppCredentials(Settings.AppId, Settings.AppPassword));

            // Create conversation parameters
            var conversationParameters = new ConversationParameters
            {
                Bot = new ChannelAccount { Id = Settings.AppId },
                Members = new List<ChannelAccount>
                {
                    new ChannelAccount { Id = Settings.RecipientId }
                },
                TenantId = Settings.TenantId,
                ChannelData = new { tenant = new { id = Settings.TenantId } }
            };

            // Create conversation
            var conversationResponse = await connectorClient.Conversations.CreateConversationAsync(conversationParameters);


            // Send message to the new conversation
            var message = new Activity(ActivityTypes.Message)
            {
                Text = "This is a proactive message in a new conversation!",
                From = new ChannelAccount { Id = Settings.AppId },
                Conversation = new ConversationAccount
                {
                    Id = conversationResponse.Id,
                    TenantId = Settings.TenantId
                }
            };

            var response = await connectorClient.Conversations.SendToConversationAsync(message);
        }

        /// <summary>
        /// Verifies the bot can send an Adaptive Card attachment to an existing Teams conversation.
        /// </summary>
        [Fact()]
        public async Task SendAdaptiveCardExample_Test()
        {
            var credentials = new MicrosoftAppCredentials(Settings.AppId, Settings.AppPassword);
            var connectorClient = new ConnectorClient(new Uri(Settings.ServiceUrl), credentials);

            // Create adaptive card (matching the pattern in TeamsBotUserService)
            var adaptiveCard = new AdaptiveCard("1.4")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = "Task Notification",
                        Size = AdaptiveTextSize.Large,
                        Weight = AdaptiveTextWeight.Bolder
                    },
                    new AdaptiveFactSet
                    {
                        Facts = new List<AdaptiveFact>
                        {
                            new AdaptiveFact("Summary:", "Test Task"),
                            new AdaptiveFact("Type:", "RenewGroup"),
                            new AdaptiveFact("Status:", "Pending")
                        }
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
                                        Text = "Expiration Date:",
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Size = AdaptiveTextSize.Medium
                                    }
                                }
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = DateTime.Now.AddDays(30).ToString("yyyy-MM-dd"),
                                        IsSubtle = true
                                    }
                                }
                            }
                        }
                    }
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveOpenUrlAction
                    {
                        Title = "View Details",
                        Url = new Uri("https://teams.microsoft.com/l/entity/YOUR_APP_ID/teams-app-home")
                    }
                }
            };

            var message = new Activity(ActivityTypes.Message)
            {
                Attachments = new List<Attachment>
                {
                    new Attachment
                    {
                        ContentType = "application/vnd.microsoft.card.adaptive",
                        Content = JsonConvert.DeserializeObject(adaptiveCard.ToJson())
                    }
                },
                From = new ChannelAccount { Id = Settings.AppId },
                Recipient = new ChannelAccount { Id = Settings.RecipientId },
                Conversation = new ConversationAccount
                {
                    Id = Settings.ConversationId,
                    TenantId = Settings.TenantId
                },
                ChannelId = "msteams"
            };

            var response = await connectorClient.Conversations.SendToConversationAsync(message);
        }

        /// <summary>
        /// Verifies markdown, adaptive card, and hero card formatting samples can be sent successfully.
        /// </summary>
        [Fact()]
        public async Task SendHeroCardMessage_Test()
        {
            var connectorClient = new ConnectorClient(new Uri(Settings.ServiceUrl), new MicrosoftAppCredentials(Settings.AppId, Settings.AppPassword));

            // Create a Hero Card
            var heroCard = new HeroCard
            {
                Title = "📋 Task Assignment Notification",
                Subtitle = "Microsoft 365 Group Renewal Required",
                Text = "You have been assigned to review and approve the renewal request for **Marketing Team** workspace. " +
                       "This task is marked as *high priority* and requires action within 7 days.\n\n" +
                       "Please review the details and take appropriate action.",

                Images = new List<CardImage>
                {
                    new CardImage
                    {
                        Url = "https://upload.wikimedia.org/wikipedia/commons/thumb/4/44/Microsoft_logo.svg/200px-Microsoft_logo.svg.png",
                        Alt = "Microsoft Logo"
                    }
                },

                Buttons = new List<CardAction>
                {
                    // Open URL action
                    new CardAction
                    {
                        Type = ActionTypes.OpenUrl,
                        Title = "View Task Details",
                        Value = "https://teams.microsoft.com/l/entity/YOUR_APP_ID/teams-app-home"
                    },

                    // Message back action
                    new CardAction
                    {
                        Type = ActionTypes.MessageBack,
                        Title = "Approve",
                        Text = "approve_task",
                        DisplayText = "Task Approved",
                        Value = JsonConvert.SerializeObject(new { action = "approve", taskId = "12345" })
                    },

                    // Message back action
                    new CardAction
                    {
                        Type = ActionTypes.MessageBack,
                        Title = "Reject",
                        Text = "reject_task",
                        DisplayText = "Task Rejected",
                        Value = JsonConvert.SerializeObject(new { action = "reject", taskId = "12345" })
                    },

                    // IM back action (start chat)
                    new CardAction
                    {
                        Type = ActionTypes.ImBack,
                        Title = "Ask Question",
                        Value = "I have a question about this task"
                    }
                }
            };

            var message = MessageFactory.Attachment(heroCard.ToAttachment());
            message.From = new ChannelAccount { Id = Settings.AppId };
            message.Recipient = new ChannelAccount { Id = Settings.RecipientId };
            message.Conversation = new ConversationAccount
            {
                Id = Settings.ConversationId,
                TenantId = Settings.TenantId
            };
            message.ChannelId = "msteams";

            var response = await connectorClient.Conversations.SendToConversationAsync((Activity)message);
        }

        [Fact()]
        public async Task MessageFormattingSamples_Test()
        {
            var connectorClient = new ConnectorClient(new Uri(Settings.ServiceUrl), new MicrosoftAppCredentials(Settings.AppId, Settings.AppPassword));

            // Markdown text with various formatting
            string markdownText = @"
## Task Assignment Notification

**Important:** You have been assigned a new task!

### Task Details:
- **Task Name:** Renew Microsoft 365 Group
- **Priority:** High
- **Due Date:** *March 15, 2026*

### Description:
The workspace *'Marketing Team'* requires renewal approval. Please review and take action.

### Next Steps:
1. Review the workspace details
2. Verify the business justification
3. Approve or reject the request

[Click here to view task details](https://teams.microsoft.com/...)

---
**Note:** This task will expire in 7 days.
";

            var message = new Activity(ActivityTypes.Message)
            {
                Text = markdownText,
                TextFormat = "markdown", // Enable markdown rendering
                From = new ChannelAccount { Id = Settings.AppId },
                Recipient = new ChannelAccount { Id = Settings.RecipientId },
                Conversation = new ConversationAccount
                {
                    Id = Settings.ConversationId,
                    TenantId = Settings.TenantId
                },
                ChannelId = "msteams"
            };

            var response = await connectorClient.Conversations.SendToConversationAsync(message);
        }

        [Fact()]
        public async Task SendAdaptiveCardMessage_Test()
        {
            var credentials = new MicrosoftAppCredentials(Settings.AppId, Settings.AppPassword);
            var connectorClient = new ConnectorClient(new Uri(Settings.ServiceUrl), credentials);

            // Create an Adaptive Card
            var adaptiveCard = new AdaptiveCard("1.4")
            {
                Body = new List<AdaptiveElement>
                {
                    // Header with icon emoji
                    new AdaptiveTextBlock
                    {
                        Text = "📋 Task Assignment",
                        Size = AdaptiveTextSize.ExtraLarge,
                        Weight = AdaptiveTextWeight.Bolder,
                        Color = AdaptiveTextColor.Accent
                    },

                    // Subtitle
                    new AdaptiveTextBlock
                    {
                        Text = "You have been assigned a new task",
                        Size = AdaptiveTextSize.Medium,
                        Weight = AdaptiveTextWeight.Lighter,
                        IsSubtle = true,
                        Spacing = AdaptiveSpacing.None
                    },

                    // Separator line
                    new AdaptiveContainer
                    {
                        Separator = true,
                        Items = new List<AdaptiveElement>()
                    },

                    // Task details using FactSet
                    new AdaptiveFactSet
                    {
                        Facts = new List<AdaptiveFact>
                        {
                            new AdaptiveFact("Task:", "Renew Microsoft 365 Group"),
                            new AdaptiveFact("Workspace:", "Marketing Team"),
                            new AdaptiveFact("Type:", "RenewGroup"),
                            new AdaptiveFact("Priority:", "High"),
                            new AdaptiveFact("Requested By:", "john.doe@company.com")
                        },
                        Spacing = AdaptiveSpacing.Medium
                    },

                    // Expiration date with special formatting
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
                                        Text = "⏰ Expiration Date:",
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Color = AdaptiveTextColor.Warning
                                    }
                                }
                            },
                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Stretch,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = "March 15, 2026",
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Color = AdaptiveTextColor.Warning
                                    }
                                }
                            }
                        },
                        Separator = true,
                        Spacing = AdaptiveSpacing.Medium
                    },

                    // Description with markdown
                    new AdaptiveTextBlock
                    {
                        Text = "**Description:**\n\nThe workspace requires renewal approval. Please **review** and take action within the next *7 days*.",
                        Wrap = true,
                        Spacing = AdaptiveSpacing.Medium
                    },

                    // Image (optional)
                    new AdaptiveImage
                    {
                        Url = new Uri("https://upload.wikimedia.org/wikipedia/commons/thumb/4/44/Microsoft_logo.svg/200px-Microsoft_logo.svg.png"),
                        Size = AdaptiveImageSize.Small,
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Center,
                        Spacing = AdaptiveSpacing.Medium
                    }
                },

                Actions = new List<AdaptiveAction>
                {
                    // Primary action button
                    new AdaptiveOpenUrlAction
                    {
                        Title = "✓ View Details",
                        Url = new Uri("https://teams.microsoft.com/l/entity/YOUR_APP_ID/teams-app-home"),
                        Style = "positive"
                    },

                    // Secondary action
                    new AdaptiveOpenUrlAction
                    {
                        Title = "📧 Contact Requester",
                        Url = new Uri("https://teams.microsoft.com/l/chat/0/0?users=john.doe@company.com")
                    }
                }
            };

            // Convert to attachment
            var attachment = new Attachment
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = JsonConvert.DeserializeObject(adaptiveCard.ToJson())
            };

            var message = new Activity(ActivityTypes.Message)
            {
                Attachments = new List<Attachment> { attachment },
                From = new ChannelAccount { Id = Settings.AppId },
                Recipient = new ChannelAccount { Id = Settings.RecipientId },
                Conversation = new ConversationAccount
                {
                    Id = Settings.ConversationId,
                    TenantId = Settings.TenantId
                },
                ChannelId = "msteams"
            };

            var response = await connectorClient.Conversations.SendToConversationAsync(message);
        }

        private static BotTestSettings LoadSettings()
        {
            var configuration = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
                .AddEnvironmentVariables()
                .Build();

            return new BotTestSettings
            {
                AppId = GetRequiredConfig(configuration, "BotConfiguration:AppId"),
                AppPassword = GetRequiredConfig(configuration, "BotConfiguration:AppPassword"),
                ServiceUrl = GetRequiredConfig(configuration, "BotConfiguration:ServiceUrl"),
                TenantId = GetRequiredConfig(configuration, "BotConfiguration:TenantId"),
                ConversationId = GetRequiredConfig(configuration, "TestUser:ConversationId"),
                RecipientId = GetRequiredConfig(configuration, "TestUser:RecipientId")
            };
        }

        private static string GetRequiredConfig(IConfiguration config, string key)
        {
            var value = config[key];
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidOperationException($"Missing required configuration key: {key}");
            }

            return value;
        }
    }
}