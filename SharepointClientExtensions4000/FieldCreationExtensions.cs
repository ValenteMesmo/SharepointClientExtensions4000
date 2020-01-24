using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    public static class SecureStringExtensions
    {
        public static SecureString ToSecureString(this string value)
        {
            var secure = new SecureString();

            foreach (char c in value)
                secure.AppendChar(c);

            return secure;
        }
    }

    public static class FieldCreationExtensions
    {

        public static async Task CreateTextField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName
        )
        {
            if (await list.ContainsField(fieldDisplayName))
                throw new Exception(
                    $@"""{list.Title}"" list already have a ""{fieldDisplayName}"" field!");

            var clientContext = list.Context.AsClientContext();

            var field = list.Fields.AddFieldAsXml(
                    $"<Field Type='Text' DisplayName='{fieldInternalName}'/>"
                    , true
                    , AddFieldOptions.AddFieldToDefaultView
                );
            var textField = clientContext.CastTo<FieldText>(field);

            textField.Title = fieldDisplayName;
            textField.Update();
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task CreateDateField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName
        )
        {
            var clientContext = list.Context.AsClientContext();

            if (await list.ContainsField(fieldDisplayName))
                throw new Exception(
                    $@"""{list.Title}"" list already have a ""{fieldDisplayName}"" field!");

            var field = list.Fields.AddFieldAsXml(
                $@"<Field Type='DateTime' Format='DateTime' DisplayName='{fieldInternalName}'></Field>"
                , true
                , AddFieldOptions.AddFieldInternalNameHint
            );

            field.Title = fieldDisplayName;
            field.Update();
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task SetAttachmentEnabled(this List list, bool value)
        {
            list.EnableAttachments = value;
            list.Update();
            await list.Context.ExecuteQueryAsync();
        }

        public static async Task CreateChoiceField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName
            , params string[] choices)
        {
            if (await list.ContainsField(fieldDisplayName))
                throw new Exception(
                    $@"""{list.Title}"" list already have a ""{fieldDisplayName}"" field!");

            if (choices?.Length == 0)
                throw new Exception($@"""{fieldDisplayName}"" field needs choices to be created!");

            var choicelist = "";
            foreach (var choice in choices)
                choicelist += $"<CHOICE>{choice}</CHOICE>";

            var schemaChoiceField = $@"<Field Type='Choice' DisplayName='{fieldInternalName}' Format='Dropdown'><CHOICES>{choicelist}</CHOICES></Field>";

            var clientContext = list.Context.AsClientContext();
            var field = list.Fields.AddFieldAsXml(
                schemaChoiceField
                , true
                , AddFieldOptions.AddFieldInternalNameHint);

            field.Title = fieldDisplayName;
            field.Update();
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task CreateRichTextField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName
            )
        {
            if (await list.ContainsField(fieldDisplayName))
                throw new Exception(
                    $@"""{list.Title}"" list already have a ""{fieldDisplayName}"" field!");

            var clientContext = list.Context.AsClientContext();

            var field = list.Fields.AddFieldAsXml(
                    $"<Field Type='Note' DisplayName='{fieldInternalName}'/>"
                    , true
                    , AddFieldOptions.AddFieldToDefaultView
                );
            var textField = clientContext.CastTo<FieldMultiLineText>(field);
            textField.Title = fieldDisplayName;
            textField.RichText = true;
            textField.Update();
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task CreateNoteField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName
        )
        {
            if (await list.ContainsField(fieldDisplayName))
                throw new Exception(
                    $@"""{list.Title}"" list already have a ""{fieldDisplayName}"" field!");

            var clientContext = list.Context.AsClientContext();

            var field = list.Fields.AddFieldAsXml(
                    $"<Field Type='Note' DisplayName='{fieldInternalName}'/>"
                    , true
                    , AddFieldOptions.AddFieldToDefaultView
                );
            var textField = clientContext.CastTo<FieldMultiLineText>(field);
            textField.Title = fieldDisplayName;
            textField.RichText = false;
            textField.Update();
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task CreateBooleanField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName
        )
        {
            if (await list.ContainsField(fieldDisplayName))
                throw new Exception(
                    $@"""{list.Title}"" list already have a ""{fieldDisplayName}"" field!");

            var clientContext = list.Context.AsClientContext();
            var field = list.Fields.AddFieldAsXml(
                $"<Field Type='Boolean' DisplayName='{fieldInternalName}'/>"
                , true
                , AddFieldOptions.AddFieldToDefaultView
            );
            field.Title = fieldDisplayName;
            field.Update();
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }

        //TODO: reaname to field
        public static async Task AddColumnDateTime(
            this List list
            , string fieldInternalName
            , string fieldDisplayName
        )
        {
            if (await list.ContainsField(fieldDisplayName))
                throw new Exception(
                    $@"""{list.Title}"" list already have a ""{fieldDisplayName}"" field!");

            var clientContext = list.Context.AsClientContext();

            var field = list.Fields.AddFieldAsXml(
                $"<Field Type='DateTime' DisplayName='{fieldInternalName}'/>"
                , true
                , AddFieldOptions.AddFieldToDefaultView
            );
            field.Title = fieldDisplayName;
            field.Update();
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }

        //TODO: reaname to field
        public static async Task AddColumnNumber(
            this List list
            , string fieldInternalName
            , string fieldDisplayName
        )
        {
            if (await list.ContainsField(fieldDisplayName))
                throw new Exception(
                    $@"""{list.Title}"" list already have a ""{fieldDisplayName}"" field!");

            var clientContext = list.Context.AsClientContext();
            var field = list.Fields.AddFieldAsXml(
                $"<Field Type='Number' DisplayName='{fieldInternalName}' />"
                , true
                , AddFieldOptions.AddFieldToDefaultView
            );
            var numberField = clientContext.CastTo<FieldNumber>(field);
            //textField.MaxLength =
            numberField.Title = fieldDisplayName;
            numberField.Update();
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }

        internal static async Task CreateLookupField(
            this List list
            , string targetListDisplayName
            , string internalFieldName
            , string displayFieldName
            , bool AllowMultipleValues)
        {
            var clientContext = list.Context.AsClientContext();
            if (await list.ContainsField(displayFieldName))
                throw new Exception(
                    $@"""{list.Title}"" list already have a ""{displayFieldName}"" field!");

            var lookupFieldXml = $"<Field DisplayName=\"{internalFieldName}\" Type=\"Lookup\" />";
            var field = list.Fields.AddFieldAsXml(lookupFieldXml, true, AddFieldOptions.AddToAllContentTypes);
            var lookupField = list.Context.CastTo<FieldLookup>(field);
            var targetList = await clientContext.GetList(targetListDisplayName);

            clientContext.Load(targetList, f => f.Id);
            await clientContext.ExecuteQueryAsync();

            lookupField.LookupList = targetList.Id.ToString();
            lookupField.Title = displayFieldName;

            lookupField.LookupField = "Title";
            lookupField.AllowMultipleValues = AllowMultipleValues;
            field.Update();
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task CreateLookupField(
            this List list
            , string targetListDisplayName
            , string internalFieldName
            , string displayFieldName) =>
            await CreateLookupField(list, targetListDisplayName, internalFieldName, displayFieldName, AllowMultipleValues: false);

        public static async Task CreateLookupMultiField(
            this List list
            , string targetListDisplayName
            , string internalFieldName
            , string displayFieldName) =>
         await CreateLookupField(list, targetListDisplayName, internalFieldName, displayFieldName, AllowMultipleValues: true);

        public static async Task<bool> ContainsField(this List list, string fieldName)
        {
            var ctx = list.Context;
            var result = ctx.LoadQuery(list.Fields.Where(f => f.InternalName == fieldName));
            await ctx.ExecuteQueryAsync();
            return result.Any();
        }

        internal static async Task CreatePeoplePickerField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName
            , bool allowMultipleValues
            , FieldUserSelectionMode mode)
        {
            if (await list.ContainsField(fieldDisplayName))
                throw new Exception(
                    $@"""{list.Title}"" list already have a ""{fieldDisplayName}"" field!");

            var clientContext = list.Context.AsClientContext();

            var field = list.Fields.AddFieldAsXml(
                    $"<Field Type='UserMulti' DisplayName='{fieldInternalName}'/>"
                    , true
                    , AddFieldOptions.AddFieldToDefaultView
                );
            var userField = clientContext.CastTo<FieldUser>(field);

            userField.Title = fieldDisplayName;
            userField.Update();
            userField.SelectionMode = mode;
            userField.AllowMultipleValues = allowMultipleValues;
            list.Update();
            await clientContext.ExecuteQueryAsync();
        }

        public static async Task CreatePeopleOnlyMultiField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await CreatePeoplePickerField(list, fieldInternalName, fieldDisplayName, true, FieldUserSelectionMode.PeopleOnly);

        public static async Task CreatePeopleOnlyField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await CreatePeoplePickerField(list, fieldInternalName, fieldDisplayName, false, FieldUserSelectionMode.PeopleOnly);

        public static async Task CreateGroupOnlyField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await CreatePeoplePickerField(list, fieldInternalName, fieldDisplayName, false, FieldUserSelectionMode.GroupsOnly);

        public static async Task CreateGroupOnlyMultiField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await CreatePeoplePickerField(list, fieldInternalName, fieldDisplayName, true, FieldUserSelectionMode.GroupsOnly);

        public static async Task CreatePeopleAndGroupMultiField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await CreatePeoplePickerField(list, fieldInternalName, fieldDisplayName, true, FieldUserSelectionMode.PeopleAndGroups);

        public static async Task CreatePeopleAndGroupField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await CreatePeoplePickerField(list, fieldInternalName, fieldDisplayName, false, FieldUserSelectionMode.PeopleAndGroups);

    }
}
