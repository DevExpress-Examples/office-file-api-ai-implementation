using Azure.AI.OpenAI;
using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using Microsoft.Extensions.AI;
using OpenAI.Chat;
using System.Text.Json.Serialization;

string openAIApiKey = Environment.GetEnvironmentVariable("AZURE_OPENAI_KEY");
string openAIEndpoint = Environment.GetEnvironmentVariable("AZURE_OPENAI_ENDPOINT");
string openAIModel = Environment.GetEnvironmentVariable("AZURE_OPENAI_MODEL_NAME");

var builder = WebApplication.CreateBuilder(args);

// Create an Azure OpenAI client with endpoint and API key from helper.
var azureOpenAIClient = new AzureOpenAIClient(
    new Uri(openAIEndpoint),
    new System.ClientModel.ApiKeyCredential(openAIApiKey));

// Get a model-specific chat client and adapt it to IChatClient.
IChatClient chatClient = azureOpenAIClient
    .GetChatClient(openAIModel)
    .AsIChatClient();

// Register the chat client as a singleton in the dependency injection container.
builder.Services.AddSingleton(chatClient);
builder.Services.AddChatClient(chatClient);

// Add DevExpress AI services and register the document-processing extensions.
builder.Services.AddDevExpressAIConsole((config) => {
    config.RegisterAIDocProcessingService();
});

// Add services to the container.

builder.Services.AddControllers().AddJsonOptions(options => {

    /*
    .. Other config
    */
    options.JsonSerializerOptions.Converters.Add(new JsonStringEnumConverter());
});
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment()) {
    app.UseSwagger();
    app.UseSwaggerUI();
    //app.DescribeAllEnumsAsStrings();
}

app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();
app.Run();
