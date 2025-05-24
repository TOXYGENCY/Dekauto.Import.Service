using Dekauto.Import.Service.Domain.Interfaces;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Dekauto.Import.Service.API.Controllers
{
    public class MetricsController : ControllerBase
    {
        private readonly IRequestMetricsService requestMetricsService;

        public MetricsController(IRequestMetricsService requestMetricsService)
        {
            this.requestMetricsService = requestMetricsService;
        }
        [Route("healthcheck")]
        [HttpGet]
        public async Task<IActionResult> HealthCheckAsync()
        {
            try
            {
                return Ok(true);
            }
            catch (Exception ex)
            {
                return BadRequest(false);
            }
        }

        [Route("requests")]
        [Authorize]
        [HttpGet]
        public async Task<IActionResult> RequestsPerPeriod()
        {
            try
            {
                return Ok(requestMetricsService.GetRecentCounters());
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, "Не удалось получить количество запросов.");
            }
        }

    }
}

