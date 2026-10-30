using System.Collections.Concurrent;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using ContentMigrator.Config;

namespace ContentMigrator.SharePoint;

/// <summary>
/// Interactive Entra ID sign-in with MSAL (replaces Connect-PnPOnline -Interactive -ClientId).
/// One instance per tenant lives for the whole app session, so the operator signs in once per
/// tenant and every later connection is silent. A cross-tenant run uses two instances (source
/// and destination), each with its own app, account and in-memory token cache.
/// </summary>
public sealed class SpoAuthenticator
{
    private readonly IPublicClientApplication _app;
    private readonly ConcurrentDictionary<string, AuthenticationResult> _tokens = new(StringComparer.OrdinalIgnoreCase);
    private readonly SemaphoreSlim _interactiveGate = new(1, 1);

    public SpoAuthenticator(TenantSignIn settings)
    {
        ClientId = (settings.ClientId ?? "").Trim();
        TenantName = (settings.TenantName ?? "").Trim();
        RedirectUri = (settings.RedirectUri ?? "").Trim();
        LoginHint = (settings.LoginHint ?? "").Trim();
        _app = PublicClientApplicationBuilder
            .Create(ClientId)
            .WithAuthority($"https://login.microsoftonline.com/{AuthorityTenant(TenantName)}")
            .WithRedirectUri(RedirectUri)
            .Build();
    }

    public string ClientId { get; }
    public string TenantName { get; }
    public string RedirectUri { get; }
    public string LoginHint { get; }

    /// <summary>True when the settings that shaped this authenticator have changed.</summary>
    public bool Matches(TenantSignIn s) =>
        string.Equals(ClientId, (s.ClientId ?? "").Trim(), StringComparison.OrdinalIgnoreCase) &&
        string.Equals(TenantName, (s.TenantName ?? "").Trim(), StringComparison.OrdinalIgnoreCase) &&
        string.Equals(RedirectUri, (s.RedirectUri ?? "").Trim(), StringComparison.OrdinalIgnoreCase) &&
        string.Equals(LoginHint, (s.LoginHint ?? "").Trim(), StringComparison.OrdinalIgnoreCase);

    /// <summary>contoso | contoso.onmicrosoft.com | contoso.sharepoint.com | GUID -> authority tenant.</summary>
    public static string AuthorityTenant(string tenantName)
    {
        var t = (tenantName ?? "").Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(t)) return "organizations";
        if (Guid.TryParse(t, out _)) return t;
        if (t.EndsWith(".sharepoint.com")) return t.Split('.')[0] + ".onmicrosoft.com";
        if (t.Contains('.')) return t;
        return t + ".onmicrosoft.com";
    }

    /// <summary>The tenant's SharePoint root (https://contoso.sharepoint.com), or null for a GUID tenant.</summary>
    public static string SharePointRoot(string tenantName)
    {
        var t = (tenantName ?? "").Trim().ToLowerInvariant();
        if (string.IsNullOrEmpty(t) || Guid.TryParse(t, out _)) return null;
        var shortName = t.Split('.')[0];
        return $"https://{shortName}.sharepoint.com";
    }

    /// <summary>Expands "/sites/Foo" to "https://tenant.sharepoint.com/sites/Foo"; trims a trailing slash.</summary>
    public static string ResolveSiteUrl(string input, string tenantName)
    {
        var url = (input ?? "").Trim();
        if (url.StartsWith("/"))
        {
            var root = SharePointRoot(tenantName);
            if (root != null) url = root + url;
        }
        return url.TrimEnd('/');
    }

    public string GetAccessToken(Uri siteUri)
    {
        var resource = $"{siteUri.Scheme}://{siteUri.Host}";
        if (_tokens.TryGetValue(resource, out var cached) && cached.ExpiresOn > DateTimeOffset.UtcNow.AddMinutes(5))
        {
            return cached.AccessToken;
        }

        var result = AcquireAsync(resource).GetAwaiter().GetResult();
        _tokens[resource] = result;
        return result.AccessToken;
    }

    private async Task<AuthenticationResult> AcquireAsync(string resource)
    {
        var scopes = new[] { $"{resource}/.default" };
        var accounts = await _app.GetAccountsAsync().ConfigureAwait(false);
        var account = PickAccount(accounts);
        if (account != null)
        {
            try
            {
                return await _app.AcquireTokenSilent(scopes, account).ExecuteAsync().ConfigureAwait(false);
            }
            catch (MsalUiRequiredException)
            {
                // fall through to interactive
            }
        }

        // Only one browser prompt at a time, even if two connections race.
        await _interactiveGate.WaitAsync().ConfigureAwait(false);
        try
        {
            accounts = await _app.GetAccountsAsync().ConfigureAwait(false);
            account = PickAccount(accounts);
            if (account != null)
            {
                try
                {
                    return await _app.AcquireTokenSilent(scopes, account).ExecuteAsync().ConfigureAwait(false);
                }
                catch (MsalUiRequiredException) { }
            }

            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
            var interactive = _app.AcquireTokenInteractive(scopes)
                .WithUseEmbeddedWebView(false);       // system browser, like PnP -Interactive
            interactive = string.IsNullOrEmpty(LoginHint)
                ? interactive.WithPrompt(Prompt.SelectAccount)
                : interactive.WithLoginHint(LoginHint);  // pre-fills the configured account
            return await interactive.ExecuteAsync(cts.Token).ConfigureAwait(false);
        }
        finally
        {
            _interactiveGate.Release();
        }
    }

    /// <summary>The cached account to reuse silently: the configured one if set, else the first.</summary>
    private IAccount PickAccount(IEnumerable<IAccount> accounts)
    {
        var list = accounts.ToList();
        if (string.IsNullOrEmpty(LoginHint)) return list.FirstOrDefault();
        return list.FirstOrDefault(a => string.Equals(a.Username, LoginHint, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Creates a CSOM context that sends a bearer token on every request.</summary>
    public ClientContext CreateContext(string siteUrl)
    {
        var uri = new Uri(siteUrl);
        var ctx = new ClientContext(uri) { ClientTag = "ContentMigrator-T2T-CE/1.1" };
        ctx.ExecutingWebRequest += (_, e) =>
        {
            e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + GetAccessToken(uri);
        };
        return ctx;
    }
}
