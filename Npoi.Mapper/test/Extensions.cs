using System;

namespace test
{
    public static class Extensions
    {
        public static DateTime Truncate(this DateTime d)
        {
            return new DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second, d.Kind);
        }

        public static DateTimeOffset Truncate(this DateTimeOffset d)
        {
            return new DateTimeOffset(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second, d.Offset);
        }
    }
}
