using System;

namespace ClearPlan.Core.Integration
{
    /// <summary>Explicit site compatibility date; never changes report creation or PDF bytes.</summary>
    public static class AriaDocumentDatePolicy
    {
        public static DateTimeOffset Calculate(DateTimeOffset createdUtc, DateTimeOffset nowUtc, int safetyMinutes)
        {
            if (createdUtc == default(DateTimeOffset) || nowUtc == default(DateTimeOffset) || safetyMinutes < 0 || safetyMinutes > 60)
                throw new ArgumentException("Document-date calculation requires valid instants and a safety margin between 0 and 60 minutes.");
            DateTimeOffset cutoff;
            try { cutoff = nowUtc.ToUniversalTime().AddMinutes(-safetyMinutes); }
            catch (ArgumentOutOfRangeException)
            { throw new ArgumentException("The configured document-date safety margin is outside the supported time range."); }
            var creation = createdUtc.ToUniversalTime();
            return creation <= cutoff ? creation : cutoff;
        }
    }
}
