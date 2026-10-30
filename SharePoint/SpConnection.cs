using System.Net;
using Microsoft.SharePoint.Client;

namespace ContentMigrator.SharePoint;

/// <summary>Which tenant a connection belongs to. Same-tenant runs use Source for both sides.</summary>
public enum TenantSide { Source, Destination }

/// <summary>
/// One authenticated CSOM connection to a site (the former PnP connection object), plus
/// per-connection caches: the web's server-relative URL and resolved users.
/// </summary>
public sealed class SpConnection : IDisposable
{
    private readonly Dictionary<string, int?> _userIdCache = new(StringComparer.OrdinalIgnoreCase);

    public SpConnection(string siteUrl, ClientContext context, int maxRetries, int baseDelaySeconds,
        TenantSide side, SpoAuthenticator auth)
    {
        SiteUrl = siteUrl.TrimEnd('/');
        Context = context;
        Side = side;
        Auth = auth;
        MaxRetries = maxRetries;
        BaseDelaySeconds = baseDelaySeconds;
        var uri = new Uri(SiteUrl);
        HostRoot = $"{uri.Scheme}://{uri.Authority}";
    }

    public string SiteUrl { get; }
    public ClientContext Context { get; }
    public TenantSide Side { get; }

    /// <summary>The sign-in this connection uses (also used for REST file downloads).</summary>
    public SpoAuthenticator Auth { get; }
    public Web Web => Context.Web;
    public int MaxRetries { get; }
    public int BaseDelaySeconds { get; }

    /// <summary>https://tenant.sharepoint.com (no trailing slash).</summary>
    public string HostRoot { get; }

    /// <summary>/sites/Foo (no trailing slash; empty for the tenant root site).</summary>
    public string WebServerRelativeUrl { get; private set; }

    public string CurrentUserEmail { get; private set; }
    public string CurrentUserLogin { get; private set; }

    /// <summary>True when the signed-in account is a Site Collection Admin of this site.</summary>
    public bool CurrentUserIsSiteAdmin { get; private set; }

    /// <summary>
    /// Email if SharePoint has one, else the login name without the claims prefix
    /// ("i:0#.f|membership|admin@contoso.com" -> "admin@contoso.com").
    /// </summary>
    public string CurrentUserName =>
        !string.IsNullOrWhiteSpace(CurrentUserEmail) ? CurrentUserEmail
        : (CurrentUserLogin ?? "").Contains('|') ? CurrentUserLogin.Substring(CurrentUserLogin.LastIndexOf('|') + 1)
        : CurrentUserLogin ?? "";

    /// <summary>Validates the connection by reading the web and the signed-in user.</summary>
    public void Open()
    {
        Context.Load(Web, w => w.ServerRelativeUrl, w => w.Url, w => w.Title);
        Context.Load(Web.CurrentUser, u => u.Email, u => u.LoginName, u => u.IsSiteAdmin);
        ExecuteQuery();
        WebServerRelativeUrl = Web.ServerRelativeUrl.TrimEnd('/');
        CurrentUserEmail = Web.CurrentUser.Email;
        CurrentUserLogin = Web.CurrentUser.LoginName;
        CurrentUserIsSiteAdmin = Web.CurrentUser.IsSiteAdmin;
    }

    /// <summary>Server-relative path -> absolute, still decoded (for ResourcePath.FromDecodedUrl).</summary>
    public string ToAbsoluteUrl(string serverRelativeUrl) => HostRoot + serverRelativeUrl;

    /// <summary>
    /// ExecuteQuery with retry on throttling (HTTP 429/503) and timeouts, honouring
    /// Retry-After. The failed request is replayed with RetryQuery, so nothing pending is lost.
    /// This is what PnP did automatically under the cmdlets.
    /// </summary>
    public void ExecuteQuery()
    {
        ExecuteQueryAsync().GetAwaiter().GetResult();
    }

    private async Task ExecuteQueryAsync()
    {
        var attempt = 0;
        var delaySeconds = BaseDelaySeconds;
        ClientRequestWrapper wrapper = null;
        while (true)
        {
            try
            {
                if (wrapper?.Value != null)
                {
                    await Context.RetryQueryAsync(wrapper.Value).ConfigureAwait(false);
                }
                else
                {
                    await Context.ExecuteQueryAsync().ConfigureAwait(false);
                }
                return;
            }
            catch (WebException wex) when (attempt < MaxRetries && IsTransient(wex))
            {
                attempt++;
                wrapper = wex.Data["ClientRequest"] as ClientRequestWrapper;
                if (wrapper?.Value == null)
                {
                    // The request could not be captured for replay; surface the throttle instead
                    // of silently dropping queued work.
                    throw;
                }
                var wait = RetryAfterSeconds(wex) ?? delaySeconds;
                await Task.Delay(TimeSpan.FromSeconds(wait)).ConfigureAwait(false);
                delaySeconds = Math.Min(delaySeconds * 2, 300);
            }
        }
    }

    private static bool IsTransient(WebException wex)
    {
        if (wex.Status == WebExceptionStatus.Timeout) return true;
        return wex.Response is HttpWebResponse r && ((int)r.StatusCode == 429 || (int)r.StatusCode == 503);
    }

    private static int? RetryAfterSeconds(WebException wex)
    {
        if (wex.Response is HttpWebResponse r && int.TryParse(r.Headers["Retry-After"], out var s) && s > 0)
        {
            return s;
        }
        return null;
    }

    /// <summary>
    /// Resolves an email/login/display name to a user ID on this web (Web.EnsureUser).
    /// Returns null when the user cannot be resolved. Results are cached.
    /// </summary>
    public int? ResolveUserId(string identity, out string error)
    {
        error = null;
        if (string.IsNullOrWhiteSpace(identity)) { error = "Empty user value."; return null; }
        if (_userIdCache.TryGetValue(identity, out var cached))
        {
            if (cached == null) error = "User could not be resolved (cached).";
            return cached;
        }
        try
        {
            var user = Web.EnsureUser(identity);
            Context.Load(user, u => u.Id);
            ExecuteQuery();
            _userIdCache[identity] = user.Id;
            return user.Id;
        }
        catch (Exception ex)
        {
            _userIdCache[identity] = null;
            error = ex.Message;
            return null;
        }
    }

    public void Dispose() => Context.Dispose();
}
