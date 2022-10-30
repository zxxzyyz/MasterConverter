using System;

namespace MasterConverter
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false, Inherited = true)]
    public class DoubleQuotes : Attribute
    {
        public DoubleQuotes() {}
    }

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false, Inherited = true)]
    public class EndComma : Attribute
    {
        public EndComma() { }
    }

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false, Inherited = true)]
    public class Trim : Attribute
    {
        public Trim() { }
    }

    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false, Inherited = true)]
    public class Ignore : Attribute
    {
        public Ignore() { }
    }
}