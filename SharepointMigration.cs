using Microsoft.SharePoint.Client;
using System.Threading.Tasks;

namespace SharepointMigrations
{
    public interface SharepointMigration
    {
        string Id { get; }
        Task ExecuteAsync(ClientContext sharepoint);
    }
}
