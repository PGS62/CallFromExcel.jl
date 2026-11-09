using System;

namespace CallFromExcel.AddIn
{
    // Converts between DateTime and the date serials of the wire format ('D' and 'G'): days, with
    // a fraction for the time of day, since 1899-12-30, counted straight through - as Julia's
    // encode_for_xl/decode_from_xl compute them.
    //
    // Not DateTime.ToOADate/FromOADate, which agree from 1899-12-30 on but count earlier
    // times differently: there the whole part is the day and the fraction is a time added to
    // it, so 1899-12-29 06:00 is -1.25 rather than -0.75. That put VBA dates before 1899-12-30
    // with a time of day hours or days out, in both directions.
    public static class DateSerial
    {
        private static readonly DateTime Epoch = new DateTime(1899, 12, 30);
        private const double MillisecondsPerDay = 86400000.0;

        public static double FromDateTime(DateTime d)
        {
            long milliseconds = (d.Ticks - Epoch.Ticks) / TimeSpan.TicksPerMillisecond;
            return milliseconds / MillisecondsPerDay;
        }

        public static DateTime ToDateTime(double serial)
        {
            return Epoch.AddMilliseconds(Math.Round(serial * MillisecondsPerDay));
        }
    }
}
