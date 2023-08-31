using System.Collections.Generic;
using System.Linq;

namespace Smartmill.Reporting.OpenXml
{
    internal static class Constants
    {
        internal const char TwoPoints = ':';

        internal static bool IsNullOrEmpty<T>(this IEnumerable<T> list) => list is null || !list.Any();
    }
}