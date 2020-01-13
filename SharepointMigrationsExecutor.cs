using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SharepointMigrations
{

    public class SharepointMigrationsExecutor
    {
        private readonly ClientContext sharepoint;

        public SharepointMigrationsExecutor(
            string sharepointUrl
            , string user
            , string password
        ) : this(new ClientContext(sharepointUrl)
        {
            Credentials = new SharePointOnlineCredentials(
                user
                , password.ToSecureString()
            )
        })
        { }

        public SharepointMigrationsExecutor(ClientContext ClientContext)
        {
            sharepoint = ClientContext;
        }

        public async Task ExecuteAsync()
        {
            const string migrationsListName = "Migrations";

            if (await sharepoint.ListExists(migrationsListName) == false)
                await sharepoint.CreateHiddenList(
                    internalName: migrationsListName
                    , displayName: migrationsListName);

            var type = typeof(SharepointMigration);
            var types = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => type.IsAssignableFrom(p) && type != p);

            var existentList = await sharepoint.GetList(migrationsListName);
            var items = existentList.GetItems(new CamlQuery());
            sharepoint.Load(items);
            await sharepoint.ExecuteQueryAsync();

            var executed = items
                .OfType<ListItem>()
                .Select(f => f["Title"].ToString());

            var migrations = new List<SharepointMigration>();
            foreach (var migrationType in types)
            {
                var migrationInstance = (SharepointMigration)Activator.CreateInstance(migrationType);
                migrations.Add(migrationInstance);
            }

            foreach (var migration in migrations.OrderBy(f => f.Id))
            {
                if (executed.Contains(migration.Id))
                    continue;

                await migration.ExecuteAsync(sharepoint);

                await existentList.AddItem(new { Title = migration.Id });
            }

        }
    }
}
