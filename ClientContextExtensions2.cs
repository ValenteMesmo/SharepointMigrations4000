using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;

namespace SharepointMigrations
{
    public static class ClientContextExtensions2
    {
        internal static async Task CreateHiddenList(this ClientContext clientContext, string internalName, string displayName)
        {
            if (await clientContext.ListExists(displayName))
                throw new Exception($@"""{displayName}"" list already exists!");

            ListCreationInformation listCreationInfo = new ListCreationInformation();
            listCreationInfo.Title = displayName;
            listCreationInfo.TemplateType = (int)ListTemplateType.GenericList;
            listCreationInfo.Url = internalName;

            List list = clientContext.Web.Lists.Add(listCreationInfo);

            list.ImageUrl = "/_layouts/15/images/itgen.gif?rev=45";
            list.Hidden = true;
            list.EnableAttachments = false;
            list.EnableFolderCreation = false;
            list.EnableMinorVersions = false;
            list.EnableVersioning = false;
            list.AllowDeletion = false;
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }
    }
}
