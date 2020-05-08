using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.SharePoint.Client
{
    //https://docs.microsoft.com/en-us/sharepoint/dev/schema/field-element-field
    public static class FieldCreationExtensions
    {
        public static async Task<FieldText> AddTextField(this List list, string fieldDisplayName) =>
            await AddTextField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<FieldText> AddTextField(
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

            return textField;
        }

        public static async Task<Field> AddDateField(this List list, string fieldDisplayName) =>
            await AddDateField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<Field> AddDateField(
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
                $@"<Field Type='DateTime' Format='DateOnly' DisplayName='{fieldInternalName}'></Field>"
                , true
                , AddFieldOptions.AddFieldInternalNameHint
            );
            field.Title = fieldDisplayName;            
            field.Update();
            list.Update();
            await clientContext.ExecuteQueryAsync();

            return field;
        }

        public static async Task SetAttachmentEnabled(this List list, bool value)
        {
            list.EnableAttachments = value;
            list.Update();
            await list.Context.ExecuteQueryAsync();
        }

        public static async Task<Field> AddChoiceField(this List list, string fieldDisplayName, params string[] choices) =>
            await AddChoiceField(list, fieldDisplayName, fieldDisplayName, choices);

        public static async Task<Field> AddChoiceField(
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

            return field;
        }

        public static async Task<FieldMultiLineText> AddRichTextField(this List list, string fieldDisplayName) =>
            await AddRichTextField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<FieldMultiLineText> AddRichTextField(
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

            return textField;
        }

        public static async Task<FieldMultiLineText> AddNoteField(this List list, string fieldDisplayName) =>
            await AddNoteField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<FieldMultiLineText> AddNoteField(
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

            return textField;
        }

        public static async Task<Field> AddBooleanField(this List list, string fieldDisplayName) =>
            await AddBooleanField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<Field> AddBooleanField(
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

            return field;
        }

        public static async Task<Field> AddDateTimeField(this List list, string fieldDisplayName) =>
            await AddDateTimeField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<Field> AddDateTimeField(
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

            return field;
        }

        public static async Task<FieldNumber> AddNumberField(this List list, string fieldDisplayName) =>
            await AddNumberField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<FieldNumber> AddNumberField(
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

            return numberField;
        }

        private static async Task<FieldLookup> AddLookupField(
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

            return lookupField;
        }

        public static async Task<FieldLookup> AddLookupField(this List list, string targetListDisplayName, string displayFieldName) =>
            await AddLookupField(list, targetListDisplayName, displayFieldName, displayFieldName);

        public static async Task<FieldLookup> AddLookupField(
            this List list
            , string targetListDisplayName
            , string internalFieldName
            , string displayFieldName) =>
            await AddLookupField(list, targetListDisplayName, internalFieldName, displayFieldName, AllowMultipleValues: false);

        public static async Task<FieldLookup> AddLookupMultiField(this List list, string targetListDisplayName, string displayFieldName) =>
            await AddLookupMultiField(list, targetListDisplayName, displayFieldName, displayFieldName);

        public static async Task<FieldLookup> AddLookupMultiField(
            this List list
            , string targetListDisplayName
            , string internalFieldName
            , string displayFieldName) =>
         await AddLookupField(list, targetListDisplayName, internalFieldName, displayFieldName, AllowMultipleValues: true);

        public static async Task<bool> ContainsField(this List list, string fieldName)
        {
            var ctx = list.Context;
            var result = ctx.LoadQuery(list.Fields.Where(f => f.InternalName == fieldName));
            await ctx.ExecuteQueryAsync();
            return result.Any();
        }

        private static async Task<FieldUser> AddPeoplePickerField(
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

            return userField;
        }

        public static async Task<FieldUser> AddPeopleOnlyMultiField(this List list, string fieldDisplayName) =>
            await AddPeopleOnlyMultiField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<FieldUser> AddPeopleOnlyMultiField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await AddPeoplePickerField(list, fieldInternalName, fieldDisplayName, true, FieldUserSelectionMode.PeopleOnly);

        public static async Task<FieldUser> AddPeopleOnlyField(
           this List list
           , string fieldDisplayName) =>
           await AddPeopleOnlyField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<FieldUser> AddPeopleOnlyField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await AddPeoplePickerField(list, fieldInternalName, fieldDisplayName, false, FieldUserSelectionMode.PeopleOnly);

        public static async Task<FieldUser> AddGroupOnlyField(this List list, string fieldDisplayName) =>
            await AddGroupOnlyField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<FieldUser> AddGroupOnlyField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await AddPeoplePickerField(list, fieldInternalName, fieldDisplayName, false, FieldUserSelectionMode.GroupsOnly);

        public static async Task<FieldUser> AddGroupOnlyMultiField(
            this List list
            , string fieldDisplayName) =>
            await AddGroupOnlyMultiField(list, fieldDisplayName, fieldDisplayName);


        public static async Task<FieldUser> AddGroupOnlyMultiField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await AddPeoplePickerField(list, fieldInternalName, fieldDisplayName, true, FieldUserSelectionMode.GroupsOnly);

        public static async Task<FieldUser> AddPeopleAndGroupMultiField(
            this List list
            , string fieldDisplayName) =>
            await AddPeopleAndGroupMultiField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<FieldUser> AddPeopleAndGroupMultiField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await AddPeoplePickerField(list, fieldInternalName, fieldDisplayName, true, FieldUserSelectionMode.PeopleAndGroups);

        public static async Task<FieldUser> AddPeopleAndGroupField(
            this List list
            , string fieldDisplayName) =>
            await AddPeopleAndGroupField(list, fieldDisplayName, fieldDisplayName);

        public static async Task<FieldUser> AddPeopleAndGroupField(
            this List list
            , string fieldInternalName
            , string fieldDisplayName) =>
            await AddPeoplePickerField(list, fieldInternalName, fieldDisplayName, false, FieldUserSelectionMode.PeopleAndGroups);

    }
}
