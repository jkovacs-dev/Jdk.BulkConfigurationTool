using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Jdk.BulkConfigurationTool.Helpers
{
    internal class EnumUtils
    {
        public static string Label(Enum value)
        {
            var fi = value.GetType().GetField(value.ToString());
            var attributes = fi.GetCustomAttributes<LabelAttribute>(false);
            return (attributes.Count() > 0) ? attributes.First().Value : value.ToString();
        }

        public static bool IsOptional(Enum value)
        {
            var fi = value.GetType().GetField(value.ToString());
            return fi.IsDefined(typeof(OptionalAttribute), false);
        }

        public static T EnumValue<T>(string value) where T : struct, IConvertible
        {
            return EnumValue<T>(value, StringComparison.Ordinal);
        }

        public static T EnumValue<T>(string value, StringComparison comparisonType) where T : struct, IConvertible
        {
            var enumType = typeof(T);
            if (!enumType.IsEnum)
            {
                throw new ArgumentException($"{enumType} must be an enumerated type");
            }
            var names = Enum.GetNames(enumType);
            foreach (string name in names)
            {
                if (Label((Enum)Enum.Parse(enumType, name)).Equals(value, comparisonType))
                {
                    return (T)Enum.Parse(enumType, name);
                }
            }
            throw new ArgumentException("The string is not a StringValue or value of the specified enum.");
        }

        public static object GetDefault(Enum value)
        {
            var fi = value.GetType().GetField(value.ToString());
            var defaultOptionAttribute = fi.GetCustomAttribute<DefaultOptionAttribute>(false);
            if(defaultOptionAttribute == null)
            {
                var defaultValueAttribute = fi.GetCustomAttribute<DefaultValueAttribute>(false);
                return defaultValueAttribute?.Value;
            }
            return defaultOptionAttribute.Label;
        }

        public static IEnumerable<string> GetOptions(Enum value)
        {
            var fi = value.GetType().GetField(value.ToString());
            return fi.GetCustomAttributes<OptionAttribute>(true).Select(x => x.Label);
        }

        public static object GetSelectedOption(Enum field, object value)
        {
            var fi = field.GetType().GetField(field.ToString());
            return fi.GetCustomAttributes<OptionAttribute>(true).First(x => x.Label.Equals(value as string, StringComparison.OrdinalIgnoreCase)).Value;
        }
    }
}
