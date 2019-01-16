using System;
using System.Linq;

namespace SharepointMigrations
{
    public class SharepointMigrationsExecutor
    {
        private readonly SharepointWrapper sharepoint;

        public SharepointMigrationsExecutor(string sharepointUrl, string user, string password)
        {
            sharepoint = new SharepointWrapper(sharepointUrl, user, password);
        }

        public void Execute()
        {
            sharepoint.CreateList("Migrations", false, true);

            var type = typeof(SharepointMigration);
            var types = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => type.IsAssignableFrom(p) && type != p);

            var executed = sharepoint.GetAllListItens("Migrations");

            foreach (var migrationType in types)
            {
                SharepointMigration migrationInstance = (SharepointMigration)Activator.CreateInstance(migrationType);
                if (executed.Contains(migrationInstance.Id))
                    continue;

                migrationInstance.Execute(sharepoint);
                sharepoint.AddItem("Migrations", migrationInstance.Id);
            }
        }
    }
}
