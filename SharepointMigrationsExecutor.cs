using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

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

        public void Execute()
        {
            sharepoint.CreateList(
                internalName: "Migrations"
                , displayName: "Migrations"
                , documentLibrary: false
                , hidden: true);

            var type = typeof(SharepointMigration);
            var types = AppDomain.CurrentDomain.GetAssemblies()
                .SelectMany(s => s.GetTypes())
                .Where(p => type.IsAssignableFrom(p) && type != p);

            sharepoint.Web.Lists.RefreshLoad();
            var existentList = sharepoint.Web.Lists.GetByTitle("Migrations");
            var items = existentList.GetItems(new CamlQuery());
            sharepoint.Load(items);
            sharepoint.ExecuteQuery();
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

                migration.Execute(sharepoint);
                sharepoint.AddItem(
                    "Migrations"
                    , new { Title = migration.Id }
                );
            }

        }
    }
}
