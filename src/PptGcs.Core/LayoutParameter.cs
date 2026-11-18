using System;
using System.Linq;

namespace PptGcs.Core
{
    public sealed class LayoutParameter
    {
        public const int MaximumNameLength = 64;
        public const double MaximumMagnitude = 1000000;

        public string Id { get; }
        public string Name { get; }
        public double Value { get; }

        public LayoutParameter(string id, string name, double value)
        {
            Id = ShapeIdentity.Normalize(id);
            name = name?.Trim();
            if (String.IsNullOrEmpty(name) || name.Length > MaximumNameLength || name.Any(Char.IsControl))
                throw new ArgumentException("参数名称必须包含 1 到 64 个可显示字符。", nameof(name));
            if (Double.IsNaN(value) || Math.Abs(value) > MaximumMagnitude)
                throw new ArgumentOutOfRangeException(nameof(value), "命名参数必须在 -1000000 到 1000000 磅之间。");
            Name = name;
            Value = value;
        }
    }
}
