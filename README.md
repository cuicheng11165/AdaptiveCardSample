# Teams Bot Test Project

This C# project demonstrates how to use **Microsoft.Bot.Schema** to send messages to Microsoft Teams users, based on patterns from **TeamsBotUserService**.

## Overview

This project provides comprehensive examples of:
- Sending messages to existing conversations
- Creating proactive conversations
- Sending Adaptive Cards with Teams deep links
- Task notifications matching production patterns
- Renewal reminders and approval requests

## Prerequisites

- .NET 8.0 SDK or later
- A registered Bot in Azure Bot Service
- Microsoft Teams
- App ID and App Password from your bot registration

## Setup

1. **Register a Bot in Azure**
   - Go to Azure Portal → Create a resource → Azure Bot
   - Note your App ID and generate an App Password (Client Secret)
   - Add Microsoft Teams channel

2. **Configure the Application**
   - Open `appsettings.json`
   - Replace the placeholder values:
     - `AppId`: Your bot's Application (client) ID
     - `AppPassword`: Your bot's client secret
     - `TenantId`: Your Microsoft 365 tenant ID
     - `ConversationId`: The conversation ID (can be obtained from bot interactions)
     - `RecipientId`: The user's Azure AD object ID

3. **Install Dependencies**
   ```bash
   dotnet restore
   ```

## Project Structure

- **TeamsBotTest.csproj** - Project file with NuGet package references
- **Program.cs** - Main entry point demonstrating three messaging patterns
- **TeamsBotMessenger.cs** - Helper class with comprehensive messaging methods
- **AdaptiveCardExamples.cs** - Reusable adaptive card templates
- **appsettings.json** - Configuration file for bot credentials

## Key NuGet Packages

- **Microsoft.Bot.Schema** (4.22.0) - Bot Framework schema definitions
- **Microsoft.Bot.Connector** (4.22.0) - Connector client for sending messages
- **Microsoft.Bot.Builder** (4.22.0) - Bot Builder SDK
- **AdaptiveCards** (3.1.0) - Adaptive card library
- **Newtonsoft.Json** (13.0.3) - JSON serialization
- **Microsoft.Rest.ClientRuntime** (2.3.24) - REST client infrastructure

## Usage Examples

### Method 1: Send to Existing Conversation (TeamsBotUserService Pattern)

```csharp
var credentials = new MicrosoftAppCredentials(appId, appPassword);
var connectorClient = new ConnectorClient(new Uri(serviceUrl), credentials);

var message = new Activity(ActivityTypes.Message)
{
    Text = "Hello from Teams Bot!",
    From = new ChannelAccount { Id = appId },
    Recipient = new ChannelAccount { Id = recipientId },
    Conversation = new ConversationAccount 
    { 
        Id = conversationId,
        TenantId = tenantId
    },
    ChannelId = "msteams"
};

await connectorClient.Conversations.SendToConversationAsync(message);
```

### Method 2: Create Proactive Conversation

```csharp
var credentials = new MicrosoftAppCredentials(appId, appPassword);
var connectorClient = new ConnectorClient(new Uri(serviceUrl), credentials);

var conversationParameters = new ConversationParameters
{
    Bot = new ChannelAccount { Id = appId },
    Members = new List<ChannelAccount> 
    { 
        new ChannelAccount { Id = recipientId } 
    },
    TenantId = tenantId,
    ChannelData = new { tenant = new { id = tenantId } }
};

var conversationResponse = await connectorClient.Conversations.CreateConversationAsync(conversationParameters);
```

### Method 3: Send Task Notification Adaptive Card

```csharp
using AdaptiveCards;

var adaptiveCard = AdaptiveCardExamples.CreateTaskNotificationCard(
    taskName: "Renew Group Workspace",
    objectName: "Marketing Team",
    serviceType: "RenewGroup",
    expirationDate: "2026-02-28",
    requesterName: "John Doe",
    requesterEmail: "john.doe@company.com",
    viewDetailUrl: "https://teams.microsoft.com/l/entity/YOUR_APP_ID/teams-app-home"
);

var messenger = new TeamsBotMessenger(serviceUrl, appId, appPassword, tenantId);
await messenger.SendAdaptiveCardAsync(conversationId, recipientId, adaptiveCard);
```

### Using TeamsBotMessenger Helper Class

```csharp
var messenger = new TeamsBotMessenger(serviceUrl, appId, appPassword, tenantId);

// Simple text message
await messenger.SendTextMessageAsync(conversationId, recipientId, "Hello!");

// Task notification card
await messenger.SendTaskNotificationCardAsync(
    conversationId, 
    recipientId,
    "Renew Site Collection", 
    "RenewSite", 
    "2026-03-15",
    "https://teams.microsoft.com"
);

// Hero card
var buttons = new List<CardAction>
{
    new CardAction
    {
        Type = ActionTypes.OpenUrl,
        Title = "Take Action",
        Value = "https://example.com"
    }
};
await messenger.SendHeroCardAsync(conversationId, recipientId, "Title", "Subtitle", "Description", buttons);
```

## Adaptive Card Examples

The project includes several pre-built adaptive card templates:

### 1. Task Notification Card
```csharp
var card = AdaptiveCardExamples.CreateTaskNotificationCard(
    taskName: "Renew Microsoft 365 Group",
    objectName: "Sales Department",
    serviceType: "RenewGroup",
    expirationDate: "2026-03-01",
    requesterName: "Admin User",
    requesterEmail: "admin@company.com",
    viewDetailUrl: "https://teams.microsoft.com/..."
);
```

### 2. Renewal Reminder Card
```csharp
var card = AdaptiveCardExamples.CreateRenewalReminderCard(
    workspaceName: "Project Alpha",
    workspaceType: "Teams",
    daysUntilExpiration: 7,
    renewUrl: "https://teams.microsoft.com/..."
);
```

### 3. Approval Request Card
```csharp
var card = AdaptiveCardExamples.CreateApprovalRequestCard(
    requestTitle: "New Site Collection Request",
    requesterName: "Jane Smith",
    requestDetails: "Requesting approval for new SharePoint site...",
    approveUrl: "https://teams.microsoft.com/.../approve",
    rejectUrl: "https://teams.microsoft.com/.../reject"
);
```

### 4. Teams Deep Link Card
```csharp
var card = AdaptiveCardExamples.CreateTeamsDeepLinkCard(
    title: "Action Required",
    message: "Click below to view details in Teams",
    teamsAppId: "YOUR_TEAMS_APP_ID",
    contextUrl: "/#/mytask?objectid=123",
    buttonText: "View Details"
);
```

### 5. Task Completion Card
```csharp
var card = AdaptiveCardExamples.CreateTaskCompletedCard(
    taskName: "Site Renewal",
    completedBy: "john.doe@company.com",
    completedDate: DateTime.Now
);
```

## Teams Deep Links

The project demonstrates creating Teams deep links matching the TeamsBotUserService pattern:

```csharp
// Format: https://teams.microsoft.com/l/entity/{appId}/teams-app-home?webUrl={encodedUrl}&context={encodedContext}

string teamsAppId = "YOUR_TEAMS_APP_ID";
string contextUrl = "/#/mytask/renewal/mytask/group/123?email=true";
string baseUrl = "https://yourapp.com";

var urlObj = new { url = $"{baseUrl}/teams{contextUrl}", type = "viewtaskdetails" };
var urlBase64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(urlObj)));
var encodeContext = HttpUtility.UrlEncode("{\"subEntityId\":\"" + urlBase64 + "\"}");
var webUrl = baseUrl + contextUrl;

string deepLink = $"https://teams.microsoft.com/l/entity/{teamsAppId}/teams-app-home?webUrl={HttpUtility.UrlEncode(webUrl)}&context={encodeContext}";
```

### Service-Specific Deep Links (from TeamsBotUserService)

```csharp
// Renewal tasks
"/#/mytask/renewal/mytask/site/{objectId}?email=true"
"/#/mytask/renewal/mytask/group/{objectId}?email=true"
"/#/mytask/renewal/mytask/teams/{objectId}?email=true"

// Confirmation tasks
"/#/mytask/confirm/mytask/site/{objectId}?email=true"
"/#/mytask/confirm/mytask/group/{objectId}?email=true"

// Filter views
"/#/mytask/filters?batchId={batchId}"
"/#/tasks/filters?ObjectId={objectId}&Type={serviceType}"
```

## Running the Project

```bash
# Restore packages
dotnet restore

# Build the project
dotnet build

# Run the application
dotnet run
```

The program will execute three examples:
1. Send message to existing conversation
2. Create proactive conversation
3. Send adaptive card with task notification

## Getting Conversation and User IDs

To obtain the `conversationId` and `recipientId`:

1. **During Bot Interaction**: When a user messages your bot, the incoming activity contains:
   - `activity.Conversation.Id` - conversation ID
   - `activity.From.Id` - user's ID
   - `activity.ChannelData.tenant.id` - tenant ID

2. **Using Microsoft Graph API**: 
   ```csharp
   // Get user's Azure AD Object ID
   GET https://graph.microsoft.com/v1.0/users/{userPrincipalName}
   ```

3. **Teams Developer Tools**: Enable developer mode in Teams (Settings → About → Version, click 6 times) and inspect activities

## Key Patterns from TeamsBotUserService

### 1. Channel ID Constant
```csharp
private const String ChannelId = "msteams";
```

### 2. Activity Creation Pattern
```csharp
var message = new Activity(ActivityTypes.Message)
{
    Text = messageText,
    From = new ChannelAccount { Id = appId },
    Recipient = new ChannelAccount { Id = recipientId },
    Conversation = new ConversationAccount 
    { 
        Id = conversationId,
        TenantId = tenantId
    },
    ChannelId = "msteams"
};
```

### 3. Adaptive Card Structure
```csharp
// Title
new AdaptiveTextBlock
{
    Text = "Title",
    Size = AdaptiveTextSize.Large,
    Weight = AdaptiveTextWeight.Bolder
}

// Facts section
new AdaptiveFactSet
{
    Facts = new List<AdaptiveFact>
    {
        new AdaptiveFact("Label:", "Value")
    },
    Separator = true
}

// Column layout for expiration
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
                    Text = "Label",
                    Weight = AdaptiveTextWeight.Bolder
                }
            }
        },
        new AdaptiveColumn
        {
            Width = AdaptiveColumnWidth.Stretch,
            Spacing = AdaptiveSpacing.Small,
            Items = new List<AdaptiveElement>
            {
                new AdaptiveTextBlock
                {
                    Text = "Value",
                    IsSubtle = true
                }
            }
        }
    }
}
```

## Important Notes

- **Authentication**: Uses `MicrosoftAppCredentials` for bot authentication
- **Service URL Regions**: 
  - Commercial: `https://smba.trafficmanager.net/amer/`
  - GCC High: `https://smba.infra.gcc.teams.microsoft.com/gcc/`
  - DoD: `https://smba.infra.dod.teams.microsoft.us/dod/`
  - China: `https://smba.infra.partner.teams.microsoft.cn/`
- **Permissions**: Bot must be installed in the tenant to message users
- **Rate Limits**: Be aware of Teams API rate limits
- **Tenant Restrictions**: Bot must have appropriate permissions

## Message Types Supported

✅ Text messages  
✅ Adaptive cards (v1.4)  
✅ Hero cards  
✅ Thumbnail cards  
✅ Carousel (multiple attachments)  
✅ Typing indicators  
✅ Message updates  
✅ Message deletions  
✅ Proactive messages  
✅ Teams deep links  

## Troubleshooting

**401 Unauthorized**
- Verify App ID and Password are correct
- Check if bot is registered properly in Azure
- Ensure credentials haven't expired

**403 Forbidden**
- Ensure bot is installed in the Teams tenant
- Verify the bot has permission to message the user
- Check tenant allows bot messaging

**404 Not Found**
- Check if conversation ID is valid
- Ensure service URL is correct for your region
- Verify recipient ID exists

**Invalid Conversation ID**
- Conversation IDs are specific to a user-bot interaction
- You may need to create a new conversation using `CreateConversationAsync`
- Ensure tenant ID is correct

## Reference Files from TeamsBotUserService

The test code is based on these key patterns:

1. **Message Sending**: `TeamsBotUserService.SendMessages()` method
2. **Adaptive Cards**: Card generation methods with facts, columns, and actions
3. **Deep Links**: `GetTeamsDeepLink()` and `ConstructTeamsDeepTaskLink()` methods
4. **Proactive Messaging**: `CreateConversationAsync` pattern with tenant channel data
5. **Service Types**: Support for various service types (RenewGroup, RenewSite, etc.)

## Additional Resources

- [Bot Framework Documentation](https://docs.microsoft.com/en-us/azure/bot-service/)
- [Teams Bot Documentation](https://docs.microsoft.com/en-us/microsoftteams/platform/bots/what-are-bots)
- [Adaptive Cards Designer](https://adaptivecards.io/designer/)
- [Bot Schema Reference](https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.schema)
- [Teams App Manifest](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema)

## Project Files Summary

| File | Purpose |
|------|---------|
| **Program.cs** | Main entry point with 3 core examples |
| **TeamsBotMessenger.cs** | Reusable helper class for all message types |
| **AdaptiveCardExamples.cs** | 5 pre-built adaptive card templates |
| **TeamsDeepLinkHelper.cs** | Deep link construction utilities |
| **UsageExamples.cs** | 10 comprehensive usage examples |
| **appsettings.json** | Configuration for credentials |

## Quick Start Examples

### Using UsageExamples Class

```csharp
var examples = new UsageExamples(appId, appPassword, serviceUrl, tenantId);

// Send a simple message
await examples.SendSimpleTextMessage(conversationId, recipientId);

// Send task notification
await examples.SendTaskNotification(conversationId, recipientId);

// Send with Teams deep link
await examples.SendTaskWithDeepLink(conversationId, recipientId, teamsAppId, baseAppUrl);

// Run all examples
await examples.RunAllExamples(conversationId, recipientId, teamsAppId, baseAppUrl);
```

### Using TeamsDeepLinkHelper

```csharp
// Create renewal task deep link
var deepLink = TeamsDeepLinkHelper.CreateRenewalDeepLink(
    teamsAppId: "YOUR_TEAMS_APP_ID",
    baseAppUrl: "https://yourapp.com",
    objectId: workspaceId,
    serviceType: "RenewGroup"
);

// Create batch filter deep link
var batchLink = TeamsDeepLinkHelper.CreateBatchFilterDeepLink(
    teamsAppId: "YOUR_TEAMS_APP_ID",
    baseAppUrl: "https://yourapp.com",
    batchId: "batch-12345",
    isMyTask: true
);
```

## License

This is a test project for educational purposes based on TeamsBotUserService patterns.
