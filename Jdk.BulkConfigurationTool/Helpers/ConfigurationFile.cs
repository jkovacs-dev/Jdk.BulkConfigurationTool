using System;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Metadata;

namespace Jdk.BulkConfigurationTool.Helpers
{
    public class ConfigurationFile
    {

        public ConfigurationFile()
        {
            Worksheets = new Dictionary<WorkSheets, Worksheet>();
            Worksheets.Add(WorkSheets.Entities, new Worksheet
            {
                Columns = GetColumns<EntityFields>()
            });
            Worksheets.Add(WorkSheets.Attributes, new Worksheet
            {
                Columns = GetColumns<AttributeFields>()
            });
            Worksheets.Add(WorkSheets.OneToManyRelationships, new Worksheet
            {
                Columns = GetColumns<OneToManyFields>()
            });
            Worksheets.Add(WorkSheets.ManyToManyRelationships, new Worksheet
            {
                Columns = GetColumns<ManyToManyFields>()
            });
            Worksheets.Add(WorkSheets.OptionSets, new Worksheet
            {
                Columns = GetColumns<OptionSetFields>()
            });

        }

        public enum AttributeDataTypes
        {
            Boolean,
            DateTime,
            Decimal,
            Integer,
            Memo,
            Money,
            Picklist,
            String,
            Image
        }

        public enum AttributeFields
        {
            [Optional]
            [Label("Solution Unique Name")]
            SolutionUniqueName,

            [Label("Entity Logical Name (eg. new_entity)")]
            EntityLogicalName,

            [Label("Attribute Schema Name (eg. new_Field)")]
            SchemaName,

            [Label("Display Name")]
            DisplayName,

            [Optional]
            [Label("Field Requirement")]
            [DefaultOption("Optional", AttributeRequiredLevel.None)]
            [Option("Business Recommended", AttributeRequiredLevel.Recommended)]
            [Option("Business Required", AttributeRequiredLevel.ApplicationRequired)]
            RequiredLevel,

            [Optional]
            [Label("Searchable?")]
            [DefaultOption("Yes", true)]
            [Option("No", false)]
            IsValidForAdvancedFind,

            [Optional]
            [Label("Field Security?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsSecured,

            [Optional]
            [Label("Auditing?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsAuditEnabled,


            [Optional]
            [Label("Description")]
            Description,

            [Label("Data Type")]
            [Option("Single Line of Text", AttributeDataTypes.String)]
            [Option("Option Set", AttributeDataTypes.Picklist)]
            [Option("Two Options", AttributeDataTypes.Boolean)]
            [Option("Image", AttributeDataTypes.Image)]
            [Option("Whole Number", AttributeDataTypes.Integer)]
            [Option("Decimal Number", AttributeDataTypes.Decimal)]
            [Option("Currency", AttributeDataTypes.Money)]
            [Option("Multiple Lines of Text", AttributeDataTypes.Memo)]
            [Option("Date and Time", AttributeDataTypes.DateTime)]
            DataType,

            [Optional]
            [Label("Text - Maximum Length")]
            MaxLength,

            [Optional]
            [Label("Single Line of Text - Format")]
            [Option("Email", 0)]
            [DefaultOption("Text", 1)]
            [Option("Text Area", 2)]
            [Option("URL", 3)]
            [Option("Ticker Symbol", 4)]
            [Option("Phone", 7)]
            StringFormat,

            [Optional]
            [Label("Whole Number - Format")]
            [DefaultOption("None", Microsoft.Xrm.Sdk.Metadata.IntegerFormat.None)]
            [Option("Duration", Microsoft.Xrm.Sdk.Metadata.IntegerFormat.Duration)]
            [Option("Time Zone", Microsoft.Xrm.Sdk.Metadata.IntegerFormat.TimeZone)]
            [Option("Language", Microsoft.Xrm.Sdk.Metadata.IntegerFormat.Language)]
            IntegerFormat,

            [Optional]
            [Label("Date and Time - Format")]
            [DefaultOption("Date Only", Microsoft.Xrm.Sdk.Metadata.DateTimeFormat.DateOnly)]
            [Option("Date and Time", Microsoft.Xrm.Sdk.Metadata.DateTimeFormat.DateAndTime)]
            DateTimeFormat,

            [Optional]
            [Label("Date and Time - Behavior")]
            [DefaultOption("User Local", 1)] // Microsoft.Xrm.Sdk.Metadata.DateTimeBehavior.UserLocal
            [Option("Date Only", 2)] // Microsoft.Xrm.Sdk.Metadata.DateTimeBehavior.DateOnly
            [Option("Time-Zone Independent", 3)] // Microsoft.Xrm.Sdk.Metadata.DateTimeBehavior.TimeZoneIndependent
            DateTimeBehavior,

            [Optional]
            [Label("Numbers - Minimum Value")]
            MinimumValue,

            [Optional]
            [Label("Numbers - Maximum Value")]
            MaximumValue,

            [Optional]
            [Label("Decimal / Money - Precision")]
            [DefaultValue(2)]
            Precision,

            [Optional]
            [Label("IME Mode")]
            [DefaultOption("Auto", Microsoft.Xrm.Sdk.Metadata.ImeMode.Auto)]
            [Option("Inactive", Microsoft.Xrm.Sdk.Metadata.ImeMode.Inactive)]
            [Option("Active", Microsoft.Xrm.Sdk.Metadata.ImeMode.Active)]
            [Option("Disabled", Microsoft.Xrm.Sdk.Metadata.ImeMode.Disabled)]
            ImeMode,

            [Optional]
            [Label("Existing Option Set Logical Name")]
            GlobalOptionSet,

            [Optional]
            [Label("Options (Format: Label|Value;Label|Value;...)")]
            Options,
            
            [Optional]
            [Label("Default Value")]
            DefaultValue
        }

        public enum EntityFields
        {
            [Optional]
            [Label("Solution Unique Name")]
            SolutionUniqueName,

            [Label("Entity Schema Name (eg. new_Entity)")]
            SchemaName,

            [Label("Display Name")]
            DisplayName,

            [Label("Plural Name")]
            DisplayCollectionName,

            [Optional]
            [Label("Description")]
            Description,

            [Optional]
            [Label("Ownership (User/Organization)")]
            [DefaultOption("User or Team", OwnershipTypes.UserOwned)]
            [Option("Organization", OwnershipTypes.OrganizationOwned)]
            OwnershipType,

            [Optional]
            [Label("Color")]
            EntityColor,

            [Optional]
            [Label("Primary Attribute Schema Name")]
            [DefaultValue("new_Name")]
            PrimaryAttributeSchemaName,

            [Optional]
            [Label("Primary Attribute Display Name")]
            [DefaultValue("Name")]
            PrimaryAttributeDisplayName,

            [Optional]
            [Label("Primary Attribute Max Length")]
            [DefaultValue(100)]
            PrimaryAttributeMaxLength,

            [Optional]
            [Label("Primary Attribute Description")]
            PrimaryAttributeDescription,

            [Optional]
            [Label("Business Process Flows?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsBusinessProcessEnabled,

            [Optional]
            [Label("Notes?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            HasNotes,

            [Optional]
            [Label("Activities?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            HasActivities,

            [Optional]
            [Label("Connections?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsConnectionsEnabled,

            // TODO: Find Sending email setting.

            [Optional]
            [Label("Mail Merge?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsMailMergeEnabled,

            [Optional]
            [Label("Document Management?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsDocumentManagementEnabled,

            // TODO: Find Access Teams setting.

            [Optional]
            [Label("Queues?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsValidForQueue,

            [Optional]
            [Label("Knowledge Management?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsKnowledgeManagementEnabled,

            [Optional]
            [Label("Allow Quick Create?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsQuickCreateEnabled,

            [Optional]
            [Label("Duplicate Detection?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsDuplicateDetectionEnabled,

            [Optional]
            [Label("Auditing?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsAuditEnabled,

            [Optional]
            [Label("Change Tracking?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            ChangeTrackingEnabled,

            [Optional]
            [Label("Enable for phone express?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsVisibleInMobile,

            [Optional]
            [Label("Enable for mobile?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsVisibleInMobileClient,

            [Optional]
            [Label("Read-only in mobile?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsReadOnlyInMobileClient,

            // TODO: Find Reading pane setting

            [Optional]
            [Label("Offline Capability for CRM for Outlook?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            IsAvailableOffline,

            [Optional]
            [Label("Use Custom Help?")]
            [Option("Yes", true)]
            [DefaultOption("No", false)]
            EntityHelpUrlEnabled,

            [Optional]
            [Label("Custom Help URL")]
            EntityHelpUrl
        }

        public enum ManyToManyFields
        {
            [Optional]
            [Label("Solution Unique Name")]
            SolutionUniqueName,

            [Label("Primary Entity Logical Name (eg. new_entity)")]
            Entity1LogicalName,

            [Optional]
            [Label("Primary Entity Display Option")]
            [DefaultOption("Do not Display", AssociatedMenuBehavior.DoNotDisplay)]
            [Option("Use Custom Label", AssociatedMenuBehavior.UseLabel)]
            [Option("Use Plural Name", AssociatedMenuBehavior.UseCollectionName)]
            Entity1MenuBehavior,

            [Optional]
            [Label("Primary Entity Display Area")]
            [DefaultOption("Details", AssociatedMenuGroup.Details)]
            [Option("Marketing", AssociatedMenuGroup.Marketing)]
            [Option("Sales", AssociatedMenuGroup.Sales)]
            [Option("Service", AssociatedMenuGroup.Service)]
            Entity1MenuGroup,

            [Optional]
            [Label("Primary Entity Custom Label")]
            Entity1MenuCustomLabel,

            [Optional]
            [Label("Primary Entity Display Order")]
            [DefaultValue(10000)]
            Entity1MenuOrder,

            [Label("Related Entity Logical Name (eg. new_entity)")]
            Entity2LogicalName,

            [Optional]
            [Label("Related Entity Display Option")]
            [DefaultOption("Do not Display", AssociatedMenuBehavior.DoNotDisplay)]
            [Option("Use Custom Label", AssociatedMenuBehavior.UseLabel)]
            [Option("Use Plural Name", AssociatedMenuBehavior.UseCollectionName)]
            Entity2MenuBehavior,

            [Optional]
            [Label("Related Entity Display Area")]
            [DefaultOption("Details", AssociatedMenuGroup.Details)]
            [Option("Marketing", AssociatedMenuGroup.Marketing)]
            [Option("Sales", AssociatedMenuGroup.Sales)]
            [Option("Service", AssociatedMenuGroup.Service)]
            Entity2MenuGroup,

            [Optional]
            [Label("Related Entity Custom Label")]
            Entity2MenuCustomLabel,

            [Optional]
            [Label("Related Entity Display Order")]
            [DefaultValue(10000)]
            Entity2MenuOrder,

            [Label("Relationship Schema Name (eg. new_Entity_Contact)")]
            SchemaName,

            [Label("Relationship Entity Schema Name (eg. new_Entity_Contact)")]
            IntersectEntitySchemaName,

            [Optional]
            [Label("Searchable?")]
            [DefaultOption("Yes", true)]
            [Option("No", false)]
            Searchable

        }

        public enum OneToManyFields
        {
            [Optional]
            [Label("Solution Unique Name")]
            SolutionUniqueName,

            [Label("Primary Entity Logical Name (eg. new_entity)")]
            EntityLogicalName,

            [Label("Related Entity Logical Name (eg. contact)")]
            RelatedEntityLogicalName,

            [Label("Relationship Schema Name (eg. new_Entity_Contact)")]
            SchemaName,

            [Label("Lookup Attribute Display Name")]
            LookupAttributeDisplayName,

            [Label("Lookup Attribute Schema Name (eg. new_Parent_ContactId)")]
            LookupAttributeSchemaName,

            [Optional]
            [Label("Field Requirement")]
            [DefaultOption("Optional", AttributeRequiredLevel.None)]
            [Option("Business Recommended", AttributeRequiredLevel.Recommended)]
            [Option("Business Required", AttributeRequiredLevel.ApplicationRequired)]
            RequiredLevel,

            [Optional]
            [Label("Lookup Attribute Description")]
            LookupAttributeDescription,

            [Optional]
            [Label("Display Option")]
            [Option("Do not Display", AssociatedMenuBehavior.DoNotDisplay)]
            [Option("Use Custom Label", AssociatedMenuBehavior.UseLabel)]
            [DefaultOption("Use Plural Name", AssociatedMenuBehavior.UseCollectionName)]
            MenuBehavior,

            [Optional]
            [Label("Display Area")]
            [DefaultOption("Details", AssociatedMenuGroup.Details)]
            [Option("Marketing", AssociatedMenuGroup.Marketing)]
            [Option("Sales", AssociatedMenuGroup.Sales)]
            [Option("Service", AssociatedMenuGroup.Service)]
            MenuGroup,

            [Optional]
            [Label("Custom Label")]
            MenuCustomLabel,

            [Optional]
            [Label("Display Order")]
            [DefaultValue(10000)]
            MenuOrder,

            [Optional]
            [Label("Cascade Assign")]
            [Option("Cascade All", CascadeType.Cascade)]
            [Option("Cascade Active", CascadeType.Active)]
            [Option("Cascade User-Owned", CascadeType.UserOwned)]
            [DefaultOption("Cascade None", CascadeType.NoCascade)]
            CascadeAssign,

            [Optional]
            [Label("Cascade Share")]
            [Option("Cascade All", CascadeType.Cascade)]
            [Option("Cascade Active", CascadeType.Active)]
            [Option("Cascade User-Owned", CascadeType.UserOwned)]
            [DefaultOption("Cascade None", CascadeType.NoCascade)]
            CascadeShare,

            [Optional]
            [Label("Cascade Unshare")]
            [Option("Cascade All", CascadeType.Cascade)]
            [Option("Cascade Active", CascadeType.Active)]
            [Option("Cascade User-Owned", CascadeType.UserOwned)]
            [DefaultOption("Cascade None", CascadeType.NoCascade)]
            CascadeUnshare,

            [Optional]
            [Label("Cascade Reparent")]
            [Option("Cascade All", CascadeType.Cascade)]
            [Option("Cascade Active", CascadeType.Active)]
            [Option("Cascade User-Owned", CascadeType.UserOwned)]
            [DefaultOption("Cascade None", CascadeType.NoCascade)]
            CascadeReparent,

            [Optional]
            [Label("Cascade Delete")]
            [Option("Cascade All", CascadeType.Cascade)]
            [DefaultOption("Remove Link", CascadeType.RemoveLink)]
            [Option("Restrict", CascadeType.Restrict)]
            CascadeDelete,

            [Optional]
            [Label("Cascade Merge")]
            [DefaultOption("Cascade All", CascadeType.Cascade)]
            [Option("Cascade Active", CascadeType.Active)]
            [Option("Cascade User-Owned", CascadeType.UserOwned)]
            [Option("Cascade None", CascadeType.NoCascade)]
            CascadeMerge
        }
        public enum OptionSetFields
        {
            [Optional]
            [Label("Solution Unique Name")]
            SolutionUniqueName,

            [Label("Schema Name (eg. new_OptionSet)")]
            SchemaName,

            [Label("Display Name")]
            DisplayName,

            [Optional]
            [Label("Description")]
            Description,

            [Label("Options (Format: Label|Value;Label|Value;...)")]
            Options
        }

        public enum WorkSheets
        {
            [Label("Entities")]
            Entities,
            [Label("Attributes")]
            Attributes,
            [Label("1-N Relationships")]
            OneToManyRelationships,
            [Label("N-M Relationships")]
            ManyToManyRelationships,
            [Label("Option Sets")]
            OptionSets
        }
        public Dictionary<WorkSheets, Worksheet> Worksheets { get; private set; }


        private static List<Column> GetColumns<T>() where T : struct, IConvertible
        {
            var columns = new List<Column>();
            var c = 1;
            foreach (var field in Enum.GetValues(typeof(T)))
            {
                var fieldEnum = (Enum)field;
                columns.Add(new Column(fieldEnum, EnumUtils.Label(fieldEnum), c++, EnumUtils.IsOptional(fieldEnum)));
            }
            return columns;
        }

        public class Column
        {
            public Enum TargetField { get; private set; }
            public string Heading { get; private set; }
            public int Position { get; set; }
            public bool Optional { get; private set; }

            public Column(Enum targetField, string heading, int position, bool optional = false)
            {
                TargetField = targetField;
                Heading = heading;
                Position = position;
                Optional = optional;
            }
        }

        public class Worksheet
        {
            public List<Column> Columns { get; set; }
            public List<object[]> Data { get; set; }
            public Worksheet()
            {
                Columns = new List<Column>();
                Data = new List<object[]>();
            }

            public Worksheet(Column[] initialColumns)
            {
                Columns = new List<Column>();
                foreach (var column in initialColumns)
                {
                    Columns.Add(column);
                }
            }
        }
    }
}
