using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class ListItemExtensions
    {
        /// <summary>
        /// Helper method that simplifies the creation of a new item.
        /// </summary>
        /// <param name="itemProperties">
        /// <para>
        /// anonymous type object containing all values to the item's columns
        /// </para>
        /// <para>
        /// Example: <code>new { Title= "Example", Test = true }</code>
        /// </para>
        /// </param>
        public static async Task AddItem(
            this List list
            , dynamic itemProperties)
        {
            var props = itemProperties?.GetType().GetProperties();


            ListItem newItem = list.AddItem(new ListItemCreationInformation());
            foreach (var pair in props)
                newItem[pair.Name] = pair.GetValue(itemProperties);

            newItem.Update();
            await list.Context.AsClientContext().ExecuteQueryAsync();
        }

        public static async Task<IList<ListItem>> GetAllItems(this List list)
        {
            var context = list.Context.AsClientContext();

            ListItemCollectionPosition itemPosition = null;
            var result = new List<ListItem>();

            while (true)
            {
                var camlQuery = new CamlQuery
                {
                    ListItemCollectionPosition = itemPosition,

                    ViewXml =
                    "<View>"
                        + "<ViewFields>"
                            + "<FieldRef Name='ID' />"
                            + "<FieldRef Name='Title' />"
                        + "</ViewFields>"
                    + "<RowLimit>5000</RowLimit>"
                   + "</View>"
                };

                var itemCollection = list.GetItems(camlQuery);
                context.Load(itemCollection);
                await context.ExecuteQueryAsync();

                itemPosition = itemCollection.ListItemCollectionPosition;

                foreach (ListItem item in itemCollection)
                    result.Add(item);

                if (itemPosition == null)
                    break;
            }

            return result;
        }

        public static async Task DeleteAllItems(this List list, IProgress<int> progress)
        {
            var clientContext = list.Context.AsClientContext();

            clientContext.Load(list, f=> f.ItemCount);
            await clientContext.ExecuteQueryAsync();

            var queryLimit = 4000;
            var batchLimit = 100;
            var moreItems = true;

            var camlQuery = new CamlQuery
            {
                ViewXml = string.Format(@"
                <View>
                    <Query><Where></Where></Query>
                    <ViewFields>
                        <FieldRef Name='ID' />
                    </ViewFields>
                    <RowLimit>{0}</RowLimit>
                </View>", queryLimit)
            };

            var deletedItems = 0;
            while (moreItems)
            {
                var listItems = list.GetItems(camlQuery);

                clientContext.Load(listItems,
                    eachItem => eachItem.Include(
                        item => item,
                        item => item["ID"]));

                await clientContext.ExecuteQueryAsync();

                var totalListItems = listItems.Count;
                if (totalListItems > 0)
                {
                    for (var i = totalListItems - 1; i > -1; i--)
                    {
                        listItems[i].DeleteObject();
                        if (i % batchLimit == 0)
                            await clientContext.ExecuteQueryAsync();
                        deletedItems++;
                        progress.Report((deletedItems / list.ItemCount) * 100);
                    }
                    //why???
                    await clientContext.ExecuteQueryAsync();
                }
                else
                    moreItems = false;
            }
        }
    }
}
