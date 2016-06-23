using System;

namespace Jdk.BulkConfigurationTool.Helpers
{
    [AttributeUsage(AttributeTargets.Field)]
    internal sealed class LabelAttribute : Attribute
    {

        public LabelAttribute(string value)
        {
            Value = value;
        }
        public string Value { get; private set; }
    }

    [AttributeUsage(AttributeTargets.Field)]
    internal sealed class OptionalAttribute : Attribute
    {
        public OptionalAttribute() { }
    }

    [AttributeUsage(AttributeTargets.Method)]
    internal sealed class DataTypeHandlerAttribute : Attribute
    {

        public DataTypeHandlerAttribute(ConfigurationFile.AttributeDataTypes dataType)
        {
            DataType = dataType;
        }
        public ConfigurationFile.AttributeDataTypes DataType { get; private set; }
    }

    [AttributeUsage(AttributeTargets.Field)]
    internal sealed class DefaultValueAttribute : Attribute
    {

        public DefaultValueAttribute(object value)
        {
            Value = value;
        }
        public object Value { get; private set; }
    }

    [AttributeUsage(AttributeTargets.Field, AllowMultiple = true)]
    internal class OptionAttribute : Attribute
    {

        public OptionAttribute(string label, object value)
        {
            Label = label;
            Value = value;
            IsDefault = false;
        }
        public bool IsDefault { get; protected set; }
        public string Label { get; private set; }
        public object Value { get; private set; }
    }

    [AttributeUsage(AttributeTargets.Field)]
    internal sealed class DefaultOptionAttribute : OptionAttribute
    {
        public DefaultOptionAttribute(string label, object value) : base(label, value)
        {
            IsDefault = true;
        }
    }

}
