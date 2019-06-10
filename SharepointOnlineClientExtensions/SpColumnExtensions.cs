using System;

namespace Microsoft.SharePoint.Client
{
    public static class SpColumnExtensions
    {
        public static void AddColumnPeoplePicker(
            this ClientContext clientContext
            , string listDisplayName
            , string columnInternalName
            , string columnDisplayName
            , bool allowMultipleValues = false
            , FieldUserSelectionMode mode = FieldUserSelectionMode.PeopleOnly)
        {
            var list = clientContext.Web.Lists.GetByTitle(listDisplayName);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            IsColumnNameAvailable(
                clientContext
                , columnDisplayName
                , list
            );

            try
            {
                var field = list.Fields.AddFieldAsXml(
                        $"<Field Type='UserMulti' DisplayName='{columnInternalName}'/>"
                        , true
                        , AddFieldOptions.AddFieldToDefaultView
                    );
                var userField = clientContext.CastTo<FieldUser>(field);

                userField.Title = columnDisplayName;
                userField.Update();
                userField.SelectionMode = mode;
                userField.AllowMultipleValues = allowMultipleValues;
                list.Update();
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                throw new Exception($"Nao foi possivel criar a coluna '{columnDisplayName}' na lista '{listDisplayName}'", ex);
            }
        }

        private static bool IsColumnNameAvailable(
            this ClientContext clientContext
            , string columnName
            , List list
        )
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

        public static void AddColumnSingleLineText(
            this ClientContext clientContext
            , string listName
            , string columnInternalName
            , string columnDisplayName
            )
        {
            var list = clientContext.Web.Lists.GetByTitle(listName);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            if (IsColumnNameAvailable(clientContext, columnDisplayName, list))
                try
                {
                    var field = list.Fields.AddFieldAsXml(
                            $"<Field Type='Text' DisplayName='{columnInternalName}'/>"
                            , true
                            , AddFieldOptions.AddFieldToDefaultView
                        );
                    var textField = clientContext.CastTo<FieldText>(field);

                    textField.Title = columnDisplayName;
                    textField.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnDisplayName}' na lista '{listName}'", ex);
                }
        }

        public static void AddColumnRichText(
            this ClientContext clientContext
            , string listName
            , string columnInternalName
            , string columnDisplayName
            )
        {
            //"<Field Type='Geolocation' DisplayName='Location'/>"
            var list = clientContext.Web.Lists.GetByTitle(listName);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            if (IsColumnNameAvailable(clientContext, columnDisplayName, list))
                try
                {
                    var field = list.Fields.AddFieldAsXml(
                            $"<Field Type='Note' DisplayName='{columnInternalName}'/>"
                            , true
                            , AddFieldOptions.AddFieldToDefaultView
                        );
                    var textField = clientContext.CastTo<FieldMultiLineText>(field);
                    textField.Title = columnDisplayName;
                    textField.RichText = true;
                    textField.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnDisplayName}' na lista '{listName}'", ex);
                }
        }

        public static void AddColumnMultiLineText(
            this ClientContext clientContext
            , string listName
            , string columnInternalName
            , string columnDisplayName
            )
        {
            var list = clientContext.Web.Lists.GetByTitle(listName);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            if (IsColumnNameAvailable(clientContext, columnDisplayName, list))
                try
                {
                    var field = list.Fields.AddFieldAsXml(
                            $"<Field Type='Note' DisplayName='{columnInternalName}'/>"
                            , true
                            , AddFieldOptions.AddFieldToDefaultView
                        );
                    var textField = clientContext.CastTo<FieldMultiLineText>(field);
                    textField.Title = columnDisplayName;
                    textField.RichText = false;
                    textField.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnDisplayName}' na lista '{listName}'", ex);
                }
        }

        public static void AddColumnBoolean(
            this ClientContext clientContext
            , string listName
            , string columnInternalName
            , string columnDisplayName
            )
        {
            var list = clientContext.Web.Lists.GetByTitle(listName);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            if (IsColumnNameAvailable(clientContext, columnDisplayName, list))
                try
                {
                    var field = list.Fields.AddFieldAsXml(
                            $"<Field Type='Boolean' DisplayName='{columnInternalName}'/>"
                            , true
                            , AddFieldOptions.AddFieldToDefaultView
                        );
                    field.Title = columnDisplayName;
                    field.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnDisplayName}' na lista '{listName}'", ex);
                }
        }

        public static void AddColumnDateTime(
            this ClientContext clientContext
            , string listName
            , string columnInternalName
            , string columnDisplayName
        )
        {
            var list = clientContext.Web.Lists.GetByTitle(listName);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            if (IsColumnNameAvailable(clientContext, columnDisplayName, list))

                try
                {
                    var field = list.Fields.AddFieldAsXml(
                            $"<Field Type='DateTime' DisplayName='{columnInternalName}'/>"
                            , true
                            , AddFieldOptions.AddFieldToDefaultView
                        );
                    field.Title = columnDisplayName;
                    field.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnDisplayName}' na lista '{listName}'", ex);
                }
        }

        public static void AddColumnNumber(
            this ClientContext clientContext
            , string listName
            , string columnInternalName
            , string columnDisplayName
        )
        {
            var list = clientContext.Web.Lists.GetByTitle(listName);
            clientContext.Load(list);
            clientContext.ExecuteQuery();

            if (IsColumnNameAvailable(clientContext, columnDisplayName, list))

                try
                {
                    var field = list.Fields.AddFieldAsXml(
                        $"<Field Type='Number' DisplayName='{columnInternalName}' />"
                        , true
                        , AddFieldOptions.AddFieldToDefaultView
                    );
                    var numberField = clientContext.CastTo<FieldNumber>(field);
                    //textField.MaxLength =
                    numberField.Title = columnDisplayName;
                    numberField.Update();
                    list.Update();
                    clientContext.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Nao foi possivel criar a coluna '{columnDisplayName}' na lista '{listName}'", ex);
                }
        }
    }
}
