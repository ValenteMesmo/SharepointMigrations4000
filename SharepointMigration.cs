namespace SharepointMigrations
{
    public interface SharepointMigration
    {
        string Id { get; }
        void Execute(SharepointWrapper sharepoint);
    }
}
