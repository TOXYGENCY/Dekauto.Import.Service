namespace Dekauto.Import.Service.Domain.Interfaces
{
    public interface IRequestMetricsService
    {
        void Increment();
        List<int> GetRecentCounters();
    }
}
