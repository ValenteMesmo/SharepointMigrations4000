using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;

namespace SharepointMigrations
{
    public class SharepointWrapper
    {
        private readonly string url;
        private readonly string user;
        private readonly SecureString password;

        public SharepointWrapper(string url, string user, string password)
        {
            this.url = url;
            this.user = user;
            this.password = password.ToSecureString();
        }

        private void CreateContext(Action<ClientContext> handler)
        {
            var uri = new Uri(url);
            var cred = new SharePointOnlineCredentials(user, password);

            using (var clientContext = new ClientContext(uri) { Credentials = cred })
                handler(clientContext);
        }

        public void RenameList(string oldName, string newName)
        {
            try
            {
                CreateContext(clientContext =>
                {
                    clientContext.Web.Lists.RefreshLoad();
                    var list = clientContext.Web.Lists.GetByTitle(oldName);
                    clientContext.Load(list);
                    list.Title = newName;
                    list.Update();
                    clientContext.ExecuteQuery();
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Nao foi possivel renomear a lista '{oldName}'", ex);
            }
        }

        public void DeleteList(string name)
        {
            try
            {
                CreateContext(clientContext =>
                {
                    clientContext.Web.Lists.RefreshLoad();
                    var list = clientContext.Web.Lists.GetByTitle(name);
                    clientContext.Load(list);
                    list.AllowDeletion = true;
                    list.Update();
                    list.DeleteObject();
                    clientContext.ExecuteQuery();
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Nao foi possivel excluir a lista '{name}'", ex);
            }
        }

        public IEnumerable<string> GetAllListItens(string listName)
        {
            IEnumerable<string> result = null;

            CreateContext(clientContext =>
            {

                try
                {
                    clientContext.Web.Lists.RefreshLoad();
                    var existentList = clientContext.Web.Lists.GetByTitle(listName);
                    var items = existentList.GetItems(new CamlQuery());
                    clientContext.Load(items);
                    clientContext.ExecuteQuery();

                    result = items.OfType<ListItem>().Select(f => f["Title"].ToString());
                }
                catch
                {
                }
            });

            return result;
        }

        public void AddItem(string listName, params KeyValuePair<string,object>[] values)
        {
            CreateContext(clientContext =>
            {

                try
                {
                    clientContext.Web.Lists.RefreshLoad();
                    var existentList = clientContext.Web.Lists.GetByTitle(listName);
                    ListItem newItem = existentList.AddItem(new ListItemCreationInformation());
                    foreach (var pair in values)
                    {
                        newItem[pair.Key] = pair.Value;
                    }
                    
                    newItem.Update();
                    clientContext.ExecuteQuery();
                }
                catch
                {
                }
            });
        }

        public void CreateList(string name, bool documentLibrary = false)
        {
            CreateList(name, documentLibrary, false);
        }

        internal void CreateList(string name, bool documentLibrary, bool hidden)
        {
            CreateContext(clientContext =>
            {
                var nameAvailable = true;
                try
                {
                    clientContext.Web.Lists.RefreshLoad();
                    var existentList = clientContext.Web.Lists.GetByTitle(name);
                    clientContext.Load(existentList);
                    clientContext.ExecuteQuery();
                    nameAvailable = false;
                }
                catch { }

                if (nameAvailable)
                {
                    ListCreationInformation listCreationInfo = new ListCreationInformation();
                    listCreationInfo.Title = name;
                    listCreationInfo.TemplateType = (int)(documentLibrary ? ListTemplateType.DocumentLibrary : ListTemplateType.GenericList);
                    listCreationInfo.Url = name
                        .RemoveAccents()
                        .RemoveWhiteSpaces();

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
            });
        }

        public void AddColumnUserMulti(string listName, string columnName)
        {
            CreateContext(clientContext =>
            {
                //"<Field Type='Geolocation' DisplayName='Location'/>"
                var list = clientContext.Web.Lists.GetByTitle(listName);
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                CheckIfColumnNameIsAvailable(listName, columnName, clientContext, list);

                try
                {
                    var field = list.Fields.AddFieldAsXml(
                            $"<Field Type='UserMulti' DisplayName='{columnName.RemoveAccents().RemoveWhiteSpaces()}'/>"
                            , true
                            , AddFieldOptions.AddFieldToDefaultView
                        );
                    var userField = clientContext.CastTo<FieldUser>(field);
                    //textField.MaxLength =
                    userField.Title = columnName;
                    userField.Update();
                    userField.SelectionMode = FieldUserSelectionMode.PeopleOnly;
                    userField.AllowMultipleValues = true;
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnName}' na lista '{listName}'", ex);
                }
            });
        }


        public void AddColumnSingleLineOfText(string listName, string columnName)
        {
            CreateContext(clientContext =>
            {
                //"<Field Type='Geolocation' DisplayName='Location'/>"
                var list = clientContext.Web.Lists.GetByTitle(listName);
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                CheckIfColumnNameIsAvailable(listName, columnName, clientContext, list);

                try
                {
                    var field = list.Fields.AddFieldAsXml(
                            $"<Field Type='Text' DisplayName='{columnName.RemoveAccents().RemoveWhiteSpaces()}'/>"
                            , true
                            , AddFieldOptions.AddFieldToDefaultView
                        );
                    var textField = clientContext.CastTo<FieldText>(field);
                    //textField.MaxLength =
                    textField.Title = columnName;
                    textField.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnName}' na lista '{listName}'", ex);
                }
            });
        }

        public void AddColumnRichText(string listName, string columnName)
        {
            CreateContext(clientContext =>
            {
                //"<Field Type='Geolocation' DisplayName='Location'/>"
                var list = clientContext.Web.Lists.GetByTitle(listName);
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                CheckIfColumnNameIsAvailable(listName, columnName, clientContext, list);

                try
                {
                    var field = list.Fields.AddFieldAsXml(
                            $"<Field Type='Note' DisplayName='{columnName.RemoveAccents().RemoveWhiteSpaces()}'/>"
                            , true
                            , AddFieldOptions.AddFieldToDefaultView
                        );
                    var textField = clientContext.CastTo<FieldMultiLineText>(field);
                    textField.Title = columnName;
                    textField.RichText = true;
                    textField.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnName}' na lista '{listName}'", ex);
                }
            });
        }

        public void AddColumnBoolean(string listName, string columnName)
        {
            CreateContext(clientContext =>
            {
                var list = clientContext.Web.Lists.GetByTitle(listName);
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                CheckIfColumnNameIsAvailable(listName, columnName, clientContext, list);

                try
                {
                    var field = list.Fields.AddFieldAsXml(
                            $"<Field Type='Boolean' DisplayName='{columnName.RemoveAccents().RemoveWhiteSpaces()}'/>"
                            , true
                            , AddFieldOptions.AddFieldToDefaultView
                        );
                    field.Title = columnName;
                    field.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnName}' na lista '{listName}'", ex);
                }
            });
        }

        public void AddColumnDateTime(string listName, string columnName)
        {
            CreateContext(clientContext =>
            {
                var list = clientContext.Web.Lists.GetByTitle(listName);
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                CheckIfColumnNameIsAvailable(listName, columnName, clientContext, list);

                try
                {
                    var field = list.Fields.AddFieldAsXml(
                            $"<Field Type='DateTime' DisplayName='{columnName.RemoveAccents().RemoveWhiteSpaces()}'/>"
                            , true
                            , AddFieldOptions.AddFieldToDefaultView
                        );
                    field.Title = columnName;
                    field.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnName}' na lista '{listName}'", ex);
                }
            });
        }

        public void AddColumnNumber(string listName, string columnName)
        {
            CreateContext(clientContext =>
            {
                var list = clientContext.Web.Lists.GetByTitle(listName);
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                CheckIfColumnNameIsAvailable(listName, columnName, clientContext, list);

                try
                {
                    var field = list.Fields.AddFieldAsXml(
                        $"<Field Type='Number' DisplayName='{columnName.RemoveAccents().RemoveWhiteSpaces()}' />"
                        , true
                        , AddFieldOptions.AddFieldToDefaultView
                    );
                    var numberField = clientContext.CastTo<FieldNumber>(field);
                    //textField.MaxLength =
                    numberField.Title = columnName;
                    numberField.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnName}' na lista '{listName}'", ex);
                }
            });
        }

        private bool CheckIfColumnNameIsAvailable(string listName, string columnName, ClientContext clientContext, List list)
        {
            //TODO: check type
            var columnExists = false;
            try
            {
                var existingField = list.Fields.GetByTitle(columnName);
                clientContext.Load(existingField);

                clientContext.ExecuteQuery();
                columnExists = true;
            }
            catch { }

            return columnExists;
            //if (columnExists)
            //    throw new Exception($"Ja existe a coluna '{columnName}' na lista '{listName}'");
        }
    }
}
