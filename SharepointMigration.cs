using Microsoft.SharePoint.Client;
using System.Threading.Tasks;

namespace SharepointMigrations
{
    public interface SharepointMigration
    {
        Task ExecuteAsync(ClientContext sharepoint);
    }
}
