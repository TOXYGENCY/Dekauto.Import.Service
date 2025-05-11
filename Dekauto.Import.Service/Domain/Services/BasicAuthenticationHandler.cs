using Microsoft.AspNetCore.Authentication;
using Microsoft.Extensions.Options;
using System.Security.Claims;
using System.Text.Encodings.Web;
using System.Text;

namespace Dekauto.Import.Service.Domain.Services
{
    public class BasicAuthenticationHandler : AuthenticationHandler<AuthenticationSchemeOptions>
    {
        private readonly IConfiguration _configuration;

        public BasicAuthenticationHandler(
            IOptionsMonitor<AuthenticationSchemeOptions> options,
            ILoggerFactory logger,
            UrlEncoder encoder,
            IConfiguration configuration)
            : base(options, logger, encoder)
        {
            _configuration = configuration;
        }

        protected override async Task<AuthenticateResult> HandleAuthenticateAsync()
        {
            if (!Request.Headers.ContainsKey("Authorization"))
                return AuthenticateResult.Fail("Authorization header is missing");

            try
            {
                var authHeader = Request.Headers.Authorization.ToString();
                if (!authHeader.StartsWith("Basic ", StringComparison.OrdinalIgnoreCase))
                    return AuthenticateResult.Fail("Invalid scheme");

                var encodedCredentials = authHeader["Basic ".Length..].Trim();
                var decodedBytes = Convert.FromBase64String(encodedCredentials);
                var decodedCredentials = Encoding.UTF8.GetString(decodedBytes);
                var separatorIndex = decodedCredentials.IndexOf(':');

                if (separatorIndex < 0)
                    return AuthenticateResult.Fail("Invalid credentials format");

                var clientId = decodedCredentials[..separatorIndex];
                var clientSecret = decodedCredentials[(separatorIndex + 1)..];

                if (!ValidateCredentials(clientId, clientSecret))
                    return AuthenticateResult.Fail("Invalid credentials");

                var claims = new[] { new Claim(ClaimTypes.Name, clientId) };
                var identity = new ClaimsIdentity(claims, Scheme.Name);
                var principal = new ClaimsPrincipal(identity);
                var ticket = new AuthenticationTicket(principal, Scheme.Name);

                return AuthenticateResult.Success(ticket);
            }
            catch
            {
                return AuthenticateResult.Fail("Invalid authorization header");
            }
        }

        private bool ValidateCredentials(string clientId, string clientSecret)
        {
            var validClientId = Environment.GetEnvironmentVariable("ServiceAuth__ClientId");
            var validClientSecret = Environment.GetEnvironmentVariable("ServiceAuth__ClientSecret");

            return clientId == validClientId && clientSecret == validClientSecret;
        }
    }
}
