using AdaptiveCards;
using Microsoft.Bot.Schema;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Xunit;

namespace TeamsBotTest;

public class AdaptiveCardUsageTests
{
    [Fact]
    public void BuildCardWithAdaptiveElements_SerializesExpectedBodyElements()
    {
        var card = new AdaptiveCard("1.4")
        {
            Body = new List<AdaptiveElement>
            {
                new AdaptiveTextBlock
                {
                    Text = "Task Assignment",
                    Size = AdaptiveTextSize.Large,
                    Weight = AdaptiveTextWeight.Bolder
                },
                new AdaptiveFactSet
                {
                    Facts = new List<AdaptiveFact>
                    {
                        new AdaptiveFact("Task", "Renew Microsoft 365 Group"),
                        new AdaptiveFact("Priority", "High")
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
                                    Text = "Owner:",
                                    Weight = AdaptiveTextWeight.Bolder
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
                                    Text = "john.doe@company.com"
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
                    Title = "View details",
                    Url = new Uri("https://contoso.example/tasks/123")
                }
            }
        };

        var json = JObject.Parse(card.ToJson());

        Assert.Equal("AdaptiveCard", (string?)json["type"]);
        Assert.Equal("1.4", (string?)json["version"]);

        var body = Assert.IsType<JArray>(json["body"]);
        Assert.Collection(
            body,
            item => Assert.Equal("TextBlock", (string?)item["type"]),
            item => Assert.Equal("FactSet", (string?)item["type"]),
            item => Assert.Equal("ColumnSet", (string?)item["type"]));

        var facts = Assert.IsType<JArray>(body[1]?["facts"]);
        Assert.Equal("Task", (string?)facts[0]?["title"]);
        Assert.Equal("Renew Microsoft 365 Group", (string?)facts[0]?["value"]);

        var actions = Assert.IsType<JArray>(json["actions"]);
        Assert.Equal("Action.OpenUrl", (string?)actions[0]?["type"]);
    }

    [Fact]
    public void BuildInputCard_WithChoiceSetAndSubmitAction_SerializesInputs()
    {
        var card = new AdaptiveCard("1.4")
        {
            Body = new List<AdaptiveElement>
            {
                new AdaptiveTextBlock
                {
                    Text = "Approve request?",
                    Wrap = true
                },
                new AdaptiveChoiceSetInput
                {
                    Id = "approval",
                    Style = AdaptiveChoiceInputStyle.Expanded,
                    Choices = new List<AdaptiveChoice>
                    {
                        new AdaptiveChoice { Title = "Approve", Value = "approve" },
                        new AdaptiveChoice { Title = "Reject", Value = "reject" }
                    }
                },
                new AdaptiveTextInput
                {
                    Id = "comment",
                    Placeholder = "Add comment",
                    IsMultiline = true
                }
            },
            Actions = new List<AdaptiveAction>
            {
                new AdaptiveSubmitAction
                {
                    Title = "Submit decision",
                    Data = new { source = "adaptive-card-test", taskId = "123" }
                }
            }
        };

        var json = JObject.Parse(card.ToJson());

        var body = Assert.IsType<JArray>(json["body"]);
        Assert.Equal("Input.ChoiceSet", (string?)body[1]?["type"]);
        Assert.Equal("approval", (string?)body[1]?["id"]);
        Assert.Equal("expanded", (string?)body[1]?["style"]);

        Assert.Equal("Input.Text", (string?)body[2]?["type"]);
        Assert.True((bool?)body[2]?["isMultiline"]);

        var actions = Assert.IsType<JArray>(json["actions"]);
        Assert.Equal("Action.Submit", (string?)actions[0]?["type"]);
        Assert.Equal("adaptive-card-test", (string?)actions[0]?["data"]?["source"]);
        Assert.Equal("123", (string?)actions[0]?["data"]?["taskId"]);
    }

    [Fact]
    public void ConvertAdaptiveCardToAttachment_ProducesTeamsCompatibleAttachment()
    {
        var card = new AdaptiveCard("1.4")
        {
            Body = new List<AdaptiveElement>
            {
                new AdaptiveTextBlock { Text = "Hello from adaptive card test" }
            }
        };

        var attachment = new Attachment
        {
            ContentType = "application/vnd.microsoft.card.adaptive",
            Content = JsonConvert.DeserializeObject(card.ToJson())
        };

        Assert.Equal("application/vnd.microsoft.card.adaptive", attachment.ContentType);

        var payload = Assert.IsType<JObject>(attachment.Content);
        Assert.Equal("AdaptiveCard", (string?)payload["type"]);
        Assert.Equal("1.4", (string?)payload["version"]);
    }
}
