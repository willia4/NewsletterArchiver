using System.Collections.Immutable;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.JavaScript;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;
using MimeKit;
using QuickEPUB;

using var host = Host
    .CreateApplicationBuilder(args)
    .AddConfiguration()
    .ConfigureServices((builder, services) =>
    {
        services.Configure<MailConfig>(builder.Configuration);
        services.Configure<ArchiveConfig>(builder.Configuration);
    })
    .Build();

var mailSettings = host.Services.GetRequiredService<IOptions<MailConfig>>().Value;
var archiveSettings = host.Services.GetRequiredService<IOptions<ArchiveConfig>>().Value;

await using var styleStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("NewsletterArchiver.style.css") ?? throw new InvalidOperationException("Could not load style sheet resource");

using var client = new ImapClient();
await client.ConnectAsync(host: mailSettings.MailServer, port: 993, useSsl: true);
await client.AuthenticateAsync(userName: mailSettings.MaiLUser, password: mailSettings.MailPassword);

var inbox = client.Inbox;
await inbox.OpenAsync(FolderAccess.ReadOnly);

var doc = new Epub(archiveSettings.BookTitle, archiveSettings.SearchFrom);

var earliestDate = DateTimeOffset.MaxValue;
var latestDate = DateTimeOffset.MinValue;

BinarySearchQuery query = SearchQuery.FromContains(archiveSettings.SearchFrom).And(SearchQuery.NotSeen);
var subjectRegex = string.IsNullOrWhiteSpace(archiveSettings.SubjectRegex)
    ? null
    : new Regex(archiveSettings.SubjectRegex);


await foreach (var msg in inbox.SearchMessagesAsync(query))
{
    earliestDate = Min(earliestDate, msg.Date);
    latestDate = Max(latestDate, msg.Date);

    var subject =
        ((subjectRegex is not null)
            ? subjectRegex.Replace(msg.Subject, "")
            : msg.Subject)
        .Trim();

    var title = $"{Fmt(msg.Date)} - {subject}";
    doc.AddSection(title, CleanHtml(msg.HtmlBody, title), "style.css");
}

doc.AddResource("style.css", EpubResourceType.CSS, styleStream);
doc.Title = $"{archiveSettings.BookTitle} ({Fmt(earliestDate)} - {Fmt(latestDate)})";

await using var fs = new FileStream(Path.Join(Helpers.GetSourceDirectory(), "output.epub"), FileMode.Create);
doc.Export(fs);

var runTask = host.RunAsync();
await host.StopAsync();
await runTask;

return 0;

static string CleanHtml(string html, string? title = default)
{
    var nodesToRemove = new HashSet<string>(new string[] {"img", "#comment", "custom", "table"}, StringComparer.OrdinalIgnoreCase);
    
    var reader = SmartReader.Reader.ParseArticle("https://localhost", html, null);
    var htmlDock = new HtmlDocument();
    htmlDock.LoadHtml(reader.Content);
    
    var allNodes = htmlDock.DocumentNode.SelectNodes("//*").ToImmutableList();
    foreach (var node in allNodes.Where(node => nodesToRemove.Contains(node.Name)))
    {
        node.Remove();
    }

    if (!string.IsNullOrWhiteSpace(title) && htmlDock.DocumentNode.FirstChild is {Name: "div"} topDiv)
    {
        var titleNode = htmlDock.CreateElement("h1");
        titleNode.AppendChild(htmlDock.CreateTextNode(title.Trim()));
        
        topDiv.InsertBefore(titleNode, topDiv.FirstChild);
    }
    
    return htmlDock.DocumentNode.OuterHtml;
}

static DateTimeOffset Min(DateTimeOffset a, DateTimeOffset b)
{
    return a < b ? a : b;
}

static DateTimeOffset Max(DateTimeOffset a, DateTimeOffset b)
{
    return a > b ? a : b;
}

static string Fmt(DateTimeOffset d) => $"{d.LocalDateTime:d}";

public static class IMAPExtensions
{
    public static async IAsyncEnumerable<MimeMessage> SearchMessagesAsync(this IMailFolder mailFolder,
        SearchQuery query, [EnumeratorCancellation] CancellationToken ct = default)
    {
        var ids = await mailFolder.SearchAsync(query, ct);
        if (ids is null) yield break;
        
        await foreach (var item in mailFolder.SearchMessagesAsync(ids, ct))
        {
            yield return item;
        }
    }
    
    public static async IAsyncEnumerable<MimeMessage> SearchMessagesAsync(this IMailFolder mailFolder,
        IEnumerable<UniqueId> ids, [EnumeratorCancellation] CancellationToken ct = default)
    {
        foreach (var id in ids)
        {
            var msg = await mailFolder.GetMessageAsync(id, ct);
            if (msg is not null)
                yield return msg;
        }
    }
}
public static class HostExtensions
{
    public static HostApplicationBuilder ConfigureServices(this HostApplicationBuilder builder,
        Action<IHostApplicationBuilder, IServiceCollection> cb)
    {
        cb(builder, builder.Services);
        return builder;
    }

    public static HostApplicationBuilder AddConfiguration(this HostApplicationBuilder builder)
    {
        builder.Configuration
            .AddJsonFile(Path.Join(Helpers.GetSourceDirectory(), "appsettings.secrets.json"), optional: false)
            .AddInMemoryCollection(new Dictionary<string, string?>()
            {
                {"Logging:LogLevel:Microsoft.Hosting.Lifetime", "Error"}
            });
        return builder;
    }
}

public static class Helpers
{
    public static string GetSourceDirectory()
    {
        return GetSourceDirectoryInternal();
    }
    
    private static string GetSourceDirectoryInternal([CallerFilePath] string? callerPath = default)
    {
        return Path.GetDirectoryName(callerPath ?? "") ?? "";
    }
}

public class MailConfig
{
    public string MailServer { get; set; } = "";
    public string MaiLUser { get; set; } = "";
    public string MailPassword { get; set; } = "";
}

public class ArchiveConfig
{
    public string SearchFrom { get; set; } = "";
    public string BookTitle { get; set; } = "";
    public string SubjectRegex { get; set; } = "";
}