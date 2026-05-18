using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using OneNoteStudyPlanner.Models;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

var configuration = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json")
    .Build();

string clientId = configuration["AzureAd:ClientId"]!;

string[] scopes =
[
    "User.Read",
    "Notes.ReadWrite"
];

var options = new InteractiveBrowserCredentialOptions
{
    TenantId = "common",
    ClientId = clientId,
    RedirectUri = new Uri("http://localhost")
};

var interactiveCredential =
    new InteractiveBrowserCredential(options);

var graphClient =
    new GraphServiceClient(interactiveCredential, scopes);

Console.WriteLine("Authenticating with Microsoft...");
Console.WriteLine();

var notebooks =
    await graphClient.Me.Onenote.Notebooks.GetAsync();

if (notebooks?.Value == null || notebooks.Value.Count == 0)
{
    Console.WriteLine("No notebooks found.");
    return;
}

Console.WriteLine("Your OneNote Notebooks:");
Console.WriteLine();

foreach (var notebook in notebooks.Value)
{
    Console.WriteLine($"- {notebook.DisplayName}");
}

Console.WriteLine();

Console.Write("Enter notebook name: ");
string notebookName = Console.ReadLine()!;

var selectedNotebook = notebooks.Value
    .FirstOrDefault(n =>
        n.DisplayName!.Equals(
            notebookName,
            StringComparison.OrdinalIgnoreCase));

if (selectedNotebook == null)
{
    Console.WriteLine("Notebook not found.");
    return;
}

var sections = await graphClient
    .Me
    .Onenote
    .Notebooks[selectedNotebook.Id]
    .Sections
    .GetAsync();

if (sections?.Value == null || sections.Value.Count == 0)
{
    Console.WriteLine("No sections found.");
    return;
}

Console.WriteLine();
Console.WriteLine("Sections:");
Console.WriteLine();

foreach (var section in sections.Value)
{
    Console.WriteLine($"- {section.DisplayName}");
}

Console.WriteLine();

Console.Write("Enter section name: ");
string sectionName = Console.ReadLine()!;

var selectedSection = sections.Value
    .FirstOrDefault(s =>
        s.DisplayName!.Equals(
            sectionName,
            StringComparison.OrdinalIgnoreCase));

if (selectedSection == null)
{
    Console.WriteLine("Section not found.");
    return;
}

Console.WriteLine();
Console.WriteLine("Reading roadmap.json...");

if (!File.Exists("roadmap.json"))
{
    Console.WriteLine("roadmap.json file not found.");
    return;
}

string roadmapJson =
    await File.ReadAllTextAsync("roadmap.json");

var studyDays =
    JsonSerializer.Deserialize<List<StudyDay>>(
        roadmapJson,
        new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

if (studyDays == null || studyDays.Count == 0)
{
    Console.WriteLine("No study plan found.");
    return;
}

Console.WriteLine();
Console.WriteLine($"Found {studyDays.Count} study days.");
Console.WriteLine();

using var httpClient = new HttpClient();

var token = await interactiveCredential.GetTokenAsync(
    new TokenRequestContext(
        new[]
        {
            "https://graph.microsoft.com/.default"
        }));

httpClient.DefaultRequestHeaders.Authorization =
    new AuthenticationHeaderValue(
        "Bearer",
        token.Token);

// Get ALL existing pages in the selected section
var existingPageTitles =
    new HashSet<string>(
        StringComparer.OrdinalIgnoreCase);

var pagesResponse = await graphClient
    .Me
    .Onenote
    .Sections[selectedSection.Id]
    .Pages
    .GetAsync();

while (pagesResponse != null)
{
    if (pagesResponse.Value != null)
    {
        foreach (var page in pagesResponse.Value)
        {
            if (!string.IsNullOrWhiteSpace(page.Title))
            {
                existingPageTitles.Add(page.Title);
            }
        }
    }

    // Stop if no next page exists
    if (string.IsNullOrWhiteSpace(
        pagesResponse.OdataNextLink))
    {
        break;
    }

    // Fetch next batch of pages
    pagesResponse = await graphClient
        .Me
        .Onenote
        .Sections[selectedSection.Id]
        .Pages
        .WithUrl(pagesResponse.OdataNextLink)
        .GetAsync();
}

Console.WriteLine();
Console.WriteLine("Existing Pages:");
Console.WriteLine();

foreach (var title in existingPageTitles)
{
    Console.WriteLine($"- {title}");
}

Console.WriteLine();
Console.WriteLine(
    $"Existing pages in section: {existingPageTitles.Count}");

foreach (var studyDay in studyDays)
{
    try
    {
        string pageTitle =
            $"Day {studyDay.Day} - {studyDay.Title}";

        // Skip if page already exists
        if (existingPageTitles.Contains(pageTitle))
        {
            Console.WriteLine(
                $"Skipped Day {studyDay.Day} - Already exists");

            continue;
        }

        string html = GenerateHtmlPage(studyDay);

        using var content = new StringContent(html);

        content.Headers.ContentType =
            new MediaTypeHeaderValue("text/html");

        string endpoint =
            $"https://graph.microsoft.com/v1.0/me/onenote/sections/{selectedSection.Id}/pages";

        var response =
            await httpClient.PostAsync(endpoint, content);

        response.EnsureSuccessStatusCode();

        // Add newly created page to in-memory cache
        existingPageTitles.Add(pageTitle);

        Console.WriteLine(
            $"Created Day {studyDay.Day} - {studyDay.Title}");
    }
    catch (Exception ex)
    {
        Console.WriteLine(
            $"Failed to create Day {studyDay.Day}: {ex.Message}");
    }
}

Console.WriteLine();
Console.WriteLine("All pages processed.");

static string GenerateHtmlPage(StudyDay studyDay)
{
    StringBuilder topicsHtml = new();

    foreach (var topic in studyDay.Topics)
    {
        topicsHtml.AppendLine(
            $"<li>{WebUtility.HtmlEncode(topic)}</li>");
    }

    return $@"
<!DOCTYPE html>
<html>
<head>
    <title>Day {studyDay.Day} - {WebUtility.HtmlEncode(studyDay.Title)}</title>
    <meta charset='utf-8'>
</head>
<body>

    <h1>
        Day {studyDay.Day} - 
        {WebUtility.HtmlEncode(studyDay.Title)}
    </h1>

    <h2>Topics</h2>

    <ul>
        {topicsHtml}
    </ul>

    <h2>Tasks</h2>

    <p>☐ Complete learning</p>
    <p>☐ Practice coding</p>
    <p>☐ Revise notes</p>

    <h2>Notes</h2>

    <p></p>

</body>
</html>";
}