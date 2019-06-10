using Microsoft.SharePoint.Client;

namespace SharepointMigrations
{
    public interface SharepointMigration
    {
        string Id { get; }
        void Execute(ClientContext sharepoint);
    }
}
