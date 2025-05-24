using Dekauto.Import.Service.Domain.Interfaces;

namespace Dekauto.Import.Service.Domain.Services.Metric
{
    public class MetricsMiddleware
    {
        private readonly RequestDelegate next;
        private readonly IRequestMetricsService requestMetricsService;

        public MetricsMiddleware(RequestDelegate next, IRequestMetricsService requestMetricsService)
        {
            this.next = next;
            this.requestMetricsService = requestMetricsService;
        }

        public Task Invoke(HttpContext httpContext)
        {
            requestMetricsService.Increment();
            return next(httpContext);
        }
    }

    // Extension method used to add the middleware to the HTTP request pipeline.
    public static class MetricsMiddlewareExtensions
    {
        public static IApplicationBuilder UseMetricsMiddleware(this IApplicationBuilder builder)
        {
            return builder.UseMiddleware<MetricsMiddleware>();
        }
    }
}

