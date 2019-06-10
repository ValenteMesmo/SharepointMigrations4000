using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class SpListItemExtensions
    {
        /// <summary>
        /// Helper method that simplifies the creation of a new item.
        /// </summary>
        /// <param name="clientContext">Context that will be used on the operation</param>
        /// <param name="listDisplayName">Name of the list, where the item will be created</param>
        /// <param name="itemProperties">
        /// <para>
        /// anonymous type object containing all values to the item's columns
        /// </para>
        /// <para>
        /// Example: <code>new { Title= "Example", Test = true }</code>
        /// </para>
        /// </param>
        public static void AddItem(
            this ClientContext clientContext
            , string listDisplayName
            , dynamic itemProperties)
        {
            try
            {
                var props = itemProperties?.GetType().GetProperties();
                clientContext.Web.Lists.RefreshLoad();
                var existentList = clientContext.Web.Lists.GetByTitle(listDisplayName);
                ListItem newItem = existentList.AddItem(new ListItemCreationInformation());
                foreach (var pair in props)
                    newItem[pair.Name] = pair.GetValue(itemProperties);

                newItem.Update();
                clientContext.ExecuteQuery();
            }
            catch
            {
            }
        }

    }
}
