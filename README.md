# TeamsBotTest

`TeamsBotTest` is a .NET 8 xUnit integration test project for validating Microsoft Teams bot messaging behaviors using Bot Framework SDK.

It focuses on practical end-to-end scenarios:
- Send a plain text message to an existing Teams conversation
- Create a proactive conversation and send a message
- Send an Adaptive Card message
- Validate message-formatting samples (markdown, adaptive, hero)

## Tech Stack

- .NET 8
- xUnit
- Microsoft.Bot.Connector / Microsoft.Bot.Schema / Microsoft.Bot.Builder
- AdaptiveCards
- Newtonsoft.Json
- Microsoft.Extensions.Configuration

## Project Layout

- `Program.cs`: xUnit integration tests and configuration loading
- `BotTestSettings.cs`: strongly typed test configuration model
- `TeamsBotMessenger.cs`: reusable messaging helpers
- `AdaptiveCardExamples.cs`: adaptive card templates
- `MessageFormattingSamples.cs`: formatting sample senders
- `TeamsDeepLinkHelper.cs`: Teams deep link utilities
- `UsageExamples.cs`: extra usage examples
- `appsettings.json`: mock/default config values (safe for source control)
- `appsettings.Development.json`: local real values for development

## Configuration

The tests load config from:
1. `appsettings.json`
2. Environment variables

Current required keys:
- `BotConfiguration:AppId`
- `BotConfiguration:AppPassword`
- `BotConfiguration:ServiceUrl`
- `BotConfiguration:TenantId`
- `TestUser:ConversationId`
- `TestUser:RecipientId`

### Recommended local setup

1. Keep `appsettings.json` with mock values only.
2. Put real credentials in `appsettings.Development.json` (git-ignored).
3. Or override with environment variables.

Example PowerShell environment override:

```powershell
$env:BotConfiguration__AppId = "<your-app-id>"
$env:BotConfiguration__AppPassword = "<your-client-secret>"
$env:BotConfiguration__ServiceUrl = "https://smba.trafficmanager.net/amer/"
$env:BotConfiguration__TenantId = "<your-tenant-id>"
$env:TestUser__ConversationId = "<your-conversation-id>"
$env:TestUser__RecipientId = "<your-recipient-id>"
```

## Run

```bash
dotnet restore
dotnet build
dotnet test
```

## Quick Start (5 Minutes)

Run a single integration test from PowerShell:

1. Set environment variables with real values:

```powershell
$env:BotConfiguration__AppId = "<your-app-id>"
$env:BotConfiguration__AppPassword = "<your-client-secret-value>"
$env:BotConfiguration__ServiceUrl = "https://smba.trafficmanager.net/amer/"
$env:BotConfiguration__TenantId = "<your-tenant-id>"
$env:TestUser__ConversationId = "<your-conversation-id>"
$env:TestUser__RecipientId = "<your-recipient-id>"
```

2. Restore and build:

```powershell
dotnet restore
dotnet build
```

3. Run only one test method:

```powershell
dotnet test --filter "FullyQualifiedName~TeamsBotTest.TeamsBotIntegrationTests.SendToExistingConversation_Test"
```

4. If the test is skipped, remove `Skip = "..."` from that test's `[Fact]` attribute in `Program.cs`.

## Test Notes

- These are integration tests, not isolated unit tests.
- They require valid Azure Bot credentials and reachable Teams endpoints.
- If credentials are invalid, authentication fails with errors like `AADSTS7000215` (invalid client secret).
- Some tests are intentionally marked with `[Fact(Skip = ...)]` to avoid running all live scenarios on every test run.

## Common Troubleshooting

- `AADSTS7000215`: `AppPassword` is wrong/expired or secret ID was used instead of secret value.
- `401/403`: bot app registration, permissions, or tenant/channel setup issue.
- `404`: incorrect conversation ID, recipient ID, or service URL region mismatch.

## Service URL Regions

- Commercial: `https://smba.trafficmanager.net/amer/`
- GCC High: `https://smba.infra.gcc.teams.microsoft.com/gcc/`
- DoD: `https://smba.infra.dod.teams.microsoft.us/dod/`
- China: `https://smba.infra.partner.teams.microsoft.cn/`

## Security

- Never commit production secrets.
- Rotate credentials immediately if they were exposed.
