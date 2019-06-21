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

        private static int OneIfZero(this int value)
        {
            if (value == 0)
                return 1;
            return value;
        }

        public static async Task<IList<ListItem>> GetAllItems(this List list, IProgress<int> progress = null)
        {
            if (progress == null)
                progress = new Progress<int>();

            var context = list.Context.AsClientContext();
            context.Load(list, f => f.ItemCount);
            await context.ExecuteQueryAsync();

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
                
                progress.Report((result.Count / list.ItemCount.OneIfZero()) * 100);

                if (itemPosition == null)
                    break;
            }

            progress.Report(100);

            return result;
        }

        public static async Task DeleteAllItems(this List list, IProgress<int> progress = null)
        {
            if (progress == null)
                progress = new Progress<int>();

            var clientContext = list.Context.AsClientContext();

            clientContext.Load(list, f => f.ItemCount);
            await clientContext.ExecuteQueryAsync();

            var batchLimit = 100;

            var deletedItems = 0;
            var listItems = await list.GetAllItems(new Progress<int>(f => progress.Report(f / 2)));

            if (listItems.Count > 0)
            {
                for (var i = listItems.Count - 1; i > -1; i--)
                {
                    listItems[i].DeleteObject();
                    if (i % batchLimit == 0)
                        await clientContext.ExecuteQueryAsync();
                    deletedItems++;
                    
                    progress.Report((100 + ((deletedItems / list.ItemCount.OneIfZero()) * 100)) / 2);
                }
                await clientContext.ExecuteQueryAsync();
            }

            progress.Report(100);
        }
    }
}
