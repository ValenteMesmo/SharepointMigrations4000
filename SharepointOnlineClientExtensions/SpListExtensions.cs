using System;

namespace Microsoft.SharePoint.Client
{
    public static class SpListExtensions
    {
        public static void CreateList(this ClientContext context, string internalName, string displayName) =>
            CreateList(context, internalName, displayName, documentLibrary: false, hidden: false);

        public static void CreateLibrary(this ClientContext context, string internalName, string displayName) =>
            CreateList(context, internalName, displayName, documentLibrary: true, hidden: false);

        internal static void CreateList(this ClientContext clientContext, string internalName, string displayName, bool documentLibrary, bool hidden)
        {
            var nameAvailable = true;
            try
            {
                //TODO: refreshLoad?
                var lists = clientContext.Web.Lists;
                clientContext.Load(lists);
                clientContext.ExecuteQuery();

                var existentList = clientContext.Web.Lists.GetByTitle(internalName);
                clientContext.Load(existentList);
                clientContext.ExecuteQuery();
                nameAvailable = false;
            }
            catch { }

            if (nameAvailable)
            {
                ListCreationInformation listCreationInfo = new ListCreationInformation();
                listCreationInfo.Title = displayName;
                listCreationInfo.TemplateType = (int)(documentLibrary ? ListTemplateType.DocumentLibrary : ListTemplateType.GenericList);
                listCreationInfo.Url = internalName;

                List list = clientContext.Web.Lists.Add(listCreationInfo);
                list.EnableAttachments = false;
                list.Hidden = hidden;

                if (documentLibrary)
                    list.ImageUrl = "/_layouts/15/images/itdl.gif?rev=45";
                else
                    list.ImageUrl = "/_layouts/15/images/itgen.gif?rev=45";

                list.EnableFolderCreation = false;
                list.EnableMinorVersions = false;
                list.EnableVersioning = false;
                list.AllowDeletion = false;
                list.Update();
                clientContext.ExecuteQuery();
            }
        }

        public static void RenameList(this ClientContext clientContext, string currentDisplayName, string newDisplayName)
        {
            try
            {
                clientContext.Web.Lists.RefreshLoad();
                var list = clientContext.Web.Lists.GetByTitle(currentDisplayName);
                clientContext.Load(list);
                list.Title = newDisplayName;
                list.Update();
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                throw new Exception($"Nao foi possivel renomear a lista '{currentDisplayName}'", ex);
            }
        }

        public static void DeleteList(this ClientContext clientContext, string listDisplayName)
        {
            try
            {
                clientContext.Web.Lists.RefreshLoad();
                var list = clientContext.Web.Lists.GetByTitle(listDisplayName);
                clientContext.Load(list);
                list.AllowDeletion = true;
                list.Update();
                list.DeleteObject();
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                throw new Exception($"Nao foi possivel excluir a lista '{listDisplayName}'", ex);
            }
        }
    }
}
