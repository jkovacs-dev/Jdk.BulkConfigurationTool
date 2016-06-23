using System;
using System.Collections.Generic;
using Jdk.BulkConfigurationTool.Helpers;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Linq;
using System.Reflection;

namespace Jdk.BulkConfigurationTool.AppCode
{
    internal class CreateAttributeRequestDataMapper : RequestDataMapper
    {
        internal CreateAttributeRequestDataMapper(List<ConfigurationFile.Column> orderedColumns, int lcId) : base(orderedColumns, lcId)
        {
        }

        internal override OrganizationRequest Map(object[] dataRowValues)
        {
            var typeColumn = GetColumn(ConfigurationFile.AttributeFields.DataType);
            var typeName = dataRowValues[typeColumn.Position - 1] as string;
            var dataType = (ConfigurationFile.AttributeDataTypes)EnumUtils.GetSelectedOption(ConfigurationFile.AttributeFields.DataType, typeName);

            var request = new CreateAttributeRequest
            {
                Attribute = InvokeDataTypeHandler(dataType, dataRowValues)
            };

            request.Attribute.RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None);

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.AttributeFields)column.TargetField;

                if (value != null)
                {
                    switch (field)
                    {
                        case ConfigurationFile.AttributeFields.SolutionUniqueName:
                            request.SolutionUniqueName = value as string;
                            break;
                        case ConfigurationFile.AttributeFields.EntityLogicalName:
                            request.EntityName = (value as string).ToLower();
                            break;
                        case ConfigurationFile.AttributeFields.SchemaName:
                            request.Attribute.SchemaName = value as string;
                            break;
                        case ConfigurationFile.AttributeFields.DisplayName:
                            request.Attribute.DisplayName = new Label(value as string, LcId);
                            break;
                        case ConfigurationFile.AttributeFields.Description:
                            request.Attribute.Description = new Label(value as string, LcId);
                            break;
                        case ConfigurationFile.AttributeFields.IsValidForAdvancedFind:
                            request.Attribute.IsValidForAdvancedFind = new BooleanManagedProperty((bool)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.AttributeFields.IsSecured:
                            request.Attribute.IsSecured = (bool)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.AttributeFields.IsAuditEnabled:
                            request.Attribute.IsAuditEnabled = new BooleanManagedProperty((bool)EnumUtils.GetSelectedOption(field, value));
                            break;
                        case ConfigurationFile.AttributeFields.RequiredLevel:
                            request.Attribute.RequiredLevel = new AttributeRequiredLevelManagedProperty((AttributeRequiredLevel)EnumUtils.GetSelectedOption(field, value));
                            break;

                    }

                }
                else if (!EnumUtils.IsOptional(field))
                {
                    throw new ArgumentException($"Mandatory data field {EnumUtils.Label(field)} does not contain a value.");
                }
            }

            return request;
        }

        [DataTypeHandler(ConfigurationFile.AttributeDataTypes.Boolean)]
        private AttributeMetadata CreateBooleanAttribute(object[] dataRowValues)
        {
            var metadata = new BooleanAttributeMetadata
            {
                OptionSet = new BooleanOptionSetMetadata
                {
                    TrueOption = new OptionMetadata(new Label("Yes", LcId), 1),
                    FalseOption = new OptionMetadata(new Label("No", LcId), 0)
                },
                DefaultValue = false
            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.AttributeFields)column.TargetField;

                if (value != null && !string.IsNullOrEmpty(value as string))
                {
                    switch (field)
                    {
                        case ConfigurationFile.AttributeFields.DefaultValue:
                            metadata.DefaultValue = Convert.ToBoolean(value);
                            break;
                        case ConfigurationFile.AttributeFields.Options:
                            var options = ParseOptions(value as string);
                            foreach(var option in options)
                            {
                                if(option.Value == 0)
                                {
                                    metadata.OptionSet.FalseOption.Label = option.Label;
                                }
                                else if(option.Value == 1)
                                {
                                    metadata.OptionSet.TrueOption.Label = option.Label;
                                }
                            }
                            break;
                    }

                }
                else if (!EnumUtils.IsOptional(field))
                {
                    throw new ArgumentException($"Mandatory data field {EnumUtils.Label(field)} does not contain a value.");
                }
            }

            return metadata;
        }

        [DataTypeHandler(ConfigurationFile.AttributeDataTypes.DateTime)]
        private AttributeMetadata CreateDateTimeAttribute(object[] dataRowValues)
        {
            var metadata = new DateTimeAttributeMetadata
            {
                DateTimeBehavior = DateTimeBehavior.UserLocal,
                Format = DateTimeFormat.DateOnly,
                ImeMode = ImeMode.Auto
            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.AttributeFields)column.TargetField;

                if (value != null && !string.IsNullOrEmpty(value as string))
                {
                    switch (field)
                    {
                        case ConfigurationFile.AttributeFields.DateTimeBehavior:
                            // Hate doing it this way, but for some reason DateTimeBehaviour is inconsistent with the
                            // rest of the SDK.
                            var selectedBehavior = (int)EnumUtils.GetSelectedOption(field, value);
                            if (selectedBehavior == 2)
                            {
                                metadata.DateTimeBehavior = DateTimeBehavior.DateOnly;
                            }
                            else if (selectedBehavior == 3)
                            {
                                metadata.DateTimeBehavior = DateTimeBehavior.TimeZoneIndependent;
                            }
                            break;
                        case ConfigurationFile.AttributeFields.DateTimeFormat:
                            metadata.Format = (DateTimeFormat)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.AttributeFields.ImeMode:
                            metadata.ImeMode = (ImeMode)EnumUtils.GetSelectedOption(field, value);
                            break;
                    }

                }
                else if (!EnumUtils.IsOptional(field))
                {
                    throw new ArgumentException($"Mandatory data field {EnumUtils.Label(field)} does not contain a value.");
                }
            }

            return metadata;
        }

        [DataTypeHandler(ConfigurationFile.AttributeDataTypes.Decimal)]
        private AttributeMetadata CreateDecimalAttribute(object[] dataRowValues)
        {
            var metadata = new DecimalAttributeMetadata
            {
                Precision = 2,
                MinValue = (decimal)DecimalAttributeMetadata.MinSupportedValue,
                MaxValue = (decimal)DecimalAttributeMetadata.MaxSupportedValue,
                ImeMode = ImeMode.Auto
            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.AttributeFields)column.TargetField;

                if (value != null && !string.IsNullOrEmpty(value as string))
                {
                    switch (field)
                    {
                        case ConfigurationFile.AttributeFields.Precision:
                            var precision = (int)value;
                            if (precision >= DecimalAttributeMetadata.MinSupportedPrecision && precision <= DecimalAttributeMetadata.MaxSupportedPrecision)
                            {
                                metadata.Precision = precision;
                            }
                            break;
                        case ConfigurationFile.AttributeFields.MinimumValue:
                            metadata.MinValue = (decimal)value;
                            break;
                        case ConfigurationFile.AttributeFields.MaximumValue:
                            metadata.MaxValue = (decimal)value;
                            break;
                        case ConfigurationFile.AttributeFields.ImeMode:
                            metadata.ImeMode = (ImeMode)EnumUtils.GetSelectedOption(field, value);
                            break;
                    }

                }
                else if (!EnumUtils.IsOptional(field))
                {
                    throw new ArgumentException($"Mandatory data field {EnumUtils.Label(field)} does not contain a value.");
                }
            }
            return metadata;
        }

        [DataTypeHandler(ConfigurationFile.AttributeDataTypes.Image)]
        private AttributeMetadata CreateImageAttribute(object[] dataRowValues) => new ImageAttributeMetadata();

        [DataTypeHandler(ConfigurationFile.AttributeDataTypes.Integer)]
        private AttributeMetadata CreateIntegerAttribute(object[] dataRowValues)
        {
            var metadata = new IntegerAttributeMetadata
            {
                Format = IntegerFormat.None,
                MinValue = IntegerAttributeMetadata.MinSupportedValue,
                MaxValue = IntegerAttributeMetadata.MaxSupportedValue
            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.AttributeFields)column.TargetField;

                if (value != null && !string.IsNullOrEmpty(value as string))
                {
                    switch (field)
                    {
                        case ConfigurationFile.AttributeFields.IntegerFormat:
                            metadata.Format = (IntegerFormat)EnumUtils.GetSelectedOption(field, value);
                            break;
                        case ConfigurationFile.AttributeFields.MinimumValue:
                            var minValue = Convert.ToInt32(value);
                            if (minValue >= IntegerAttributeMetadata.MinSupportedValue && minValue <= IntegerAttributeMetadata.MaxSupportedValue)
                            {
                                metadata.MinValue = minValue;
                            }
                            break;
                        case ConfigurationFile.AttributeFields.MaximumValue:
                            var maxValue = Convert.ToInt32(value);
                            if (maxValue >= IntegerAttributeMetadata.MinSupportedValue && maxValue <= IntegerAttributeMetadata.MaxSupportedValue)
                            {
                                metadata.MinValue = maxValue;
                            }
                            break;
                    }

                }
                else if (!EnumUtils.IsOptional(field))
                {
                    throw new ArgumentException($"Mandatory data field {EnumUtils.Label(field)} does not contain a value.");
                }
            }
            return metadata;
        }

        [DataTypeHandler(ConfigurationFile.AttributeDataTypes.Memo)]
        private AttributeMetadata CreateMemoAttribute(object[] dataRowValues)
        {
            var metadata = new MemoAttributeMetadata
            {
                MaxLength = 2000,
                ImeMode = ImeMode.Auto

            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.AttributeFields)column.TargetField;

                if (value != null && !string.IsNullOrEmpty(value as string))
                {
                    switch (field)
                    {
                        case ConfigurationFile.AttributeFields.MaxLength:
                            metadata.MaxLength = Convert.ToInt32(value);
                            break;
                        case ConfigurationFile.AttributeFields.ImeMode:
                            metadata.ImeMode = (ImeMode)EnumUtils.GetSelectedOption(field, value);
                            break;
                    }

                }
                else if (!EnumUtils.IsOptional(field))
                {
                    throw new ArgumentException($"Mandatory data field {EnumUtils.Label(field)} does not contain a value.");
                }
            }

            return metadata;
        }

        [DataTypeHandler(ConfigurationFile.AttributeDataTypes.Money)]
        private AttributeMetadata CreateMoneyAttribute(object[] dataRowValues)
        {
            var metadata = new MoneyAttributeMetadata
            {
                MinValue = MoneyAttributeMetadata.MinSupportedValue,
                MaxValue = MoneyAttributeMetadata.MaxSupportedValue,
                ImeMode = ImeMode.Auto
            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.AttributeFields)column.TargetField;

                if (value != null && !string.IsNullOrEmpty(value as string))
                {
                    switch (field)
                    {
                        case ConfigurationFile.AttributeFields.Precision:
                            var precision = (int)value;
                            if (precision >= MoneyAttributeMetadata.MinSupportedPrecision && precision <= MoneyAttributeMetadata.MaxSupportedPrecision)
                            {
                                metadata.Precision = precision;
                            }
                            break;
                        case ConfigurationFile.AttributeFields.MinimumValue:
                            metadata.MinValue = (double)value;
                            break;
                        case ConfigurationFile.AttributeFields.MaximumValue:
                            metadata.MaxValue = (double)value;
                            break;
                        case ConfigurationFile.AttributeFields.ImeMode:
                            metadata.ImeMode = (ImeMode)EnumUtils.GetSelectedOption(field, value);
                            break;
                    }

                }
                else if (!EnumUtils.IsOptional(field))
                {
                    throw new ArgumentException($"Mandatory data field {EnumUtils.Label(field)} does not contain a value.");
                }
            }

            return metadata;
        }

        [DataTypeHandler(ConfigurationFile.AttributeDataTypes.Picklist)]
        private AttributeMetadata CreatePicklistAttribute(object[] dataRowValues)
        {
            var metadata = new PicklistAttributeMetadata()
            {
                OptionSet = new OptionSetMetadata
                {
                    IsGlobal = false,
                    OptionSetType = OptionSetType.Picklist
                }
            };

            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.AttributeFields)column.TargetField;

                if (value != null && !string.IsNullOrEmpty(value as string))
                {
                    switch (field)
                    {
                        case ConfigurationFile.AttributeFields.DefaultValue:
                            metadata.DefaultFormValue = Convert.ToInt32(value);
                            break;
                        case ConfigurationFile.AttributeFields.Options:
                            var options = ParseOptions(value as string);
                            metadata.OptionSet.Options.AddRange(options);
                            break;
                        case ConfigurationFile.AttributeFields.GlobalOptionSet:
                            metadata.OptionSet.IsGlobal = true;
                            metadata.OptionSet.Name = (value as string).ToLower();
                            break;
                    }

                }
                else if (!EnumUtils.IsOptional(field))
                {
                    throw new ArgumentException($"Mandatory data field {EnumUtils.Label(field)} does not contain a value.");
                }
            }

            return metadata;
        }

        [DataTypeHandler(ConfigurationFile.AttributeDataTypes.String)]
        private AttributeMetadata CreateStringAttribute(object[] dataRowValues)
        {
            var metadata = new StringAttributeMetadata
            {
                FormatName = StringFormatName.Text,
                MaxLength = 100,
                ImeMode = ImeMode.Auto
            };
            foreach (var column in Columns)
            {
                var value = dataRowValues[column.Position - 1];
                var field = (ConfigurationFile.AttributeFields)column.TargetField;

                if (value != null && !string.IsNullOrEmpty(value as string))
                {
                    switch (field)
                    {
                        case ConfigurationFile.AttributeFields.StringFormat:
                            var selectedBehavior = (int)EnumUtils.GetSelectedOption(field, value);
                            if (selectedBehavior == 0)
                            {
                                metadata.FormatName = StringFormatName.Email;
                            }
                            else if (selectedBehavior == 2)
                            {
                                metadata.FormatName = StringFormatName.TextArea;
                            }
                            else if (selectedBehavior == 3)
                            {
                                metadata.FormatName = StringFormatName.Url;
                            }
                            else if (selectedBehavior == 4)
                            {
                                metadata.FormatName = StringFormatName.TickerSymbol;
                            }
                            else if (selectedBehavior == 7)
                            {
                                metadata.FormatName = StringFormatName.Phone;
                            }
                            break;
                        case ConfigurationFile.AttributeFields.MaxLength:
                            metadata.MaxLength = Convert.ToInt32(value);
                            break;
                        case ConfigurationFile.AttributeFields.ImeMode:
                            metadata.ImeMode = (ImeMode)EnumUtils.GetSelectedOption(field, value);
                            break;
                    }

                }
                else if (!EnumUtils.IsOptional(field))
                {
                    throw new ArgumentException($"Mandatory data field {EnumUtils.Label(field)} does not contain a value.");
                }
            }

            return metadata;
        }

        private AttributeMetadata InvokeDataTypeHandler(ConfigurationFile.AttributeDataTypes dataType, object[] dataRowValues)
        {
            var handler =
                (from method in GetType().GetMethods(BindingFlags.NonPublic | BindingFlags.Instance)
                 let attributes = method.GetCustomAttributes(typeof(DataTypeHandlerAttribute), true).Where(x => ((DataTypeHandlerAttribute)x).DataType == dataType)
                 where attributes != null && attributes.Count() > 0
                 select new
                 {
                     Method = method,
                     Attributes = attributes.Cast<DataTypeHandlerAttribute>()
                 })
                .First();

            return (AttributeMetadata)handler.Method.Invoke(this, new object[] { dataRowValues });
        }
    }
}
