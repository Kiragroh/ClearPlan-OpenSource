using System;
using System.Globalization;
using System.Text;
using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    public static class ReviewSnapshotJson
    {
        private static readonly UTF8Encoding Utf8WithoutBom =
            new UTF8Encoding(false, true);

        public static string Serialize(ReviewSnapshot snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException("snapshot");
            }

            return JsonConvert.SerializeObject(snapshot, CreateSettings());
        }

        public static byte[] SerializeUtf8(ReviewSnapshot snapshot)
        {
            return Utf8WithoutBom.GetBytes(Serialize(snapshot));
        }

        public static ReviewSnapshot Deserialize(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                throw new ArgumentException("Snapshot JSON must not be empty.", "json");
            }

            ReviewSnapshot snapshot =
                JsonConvert.DeserializeObject<ReviewSnapshot>(json, CreateSettings());
            if (snapshot == null)
            {
                throw new JsonSerializationException(
                    "Snapshot JSON did not contain an object.");
            }

            return snapshot;
        }

        public static ReviewSnapshot DeserializeUtf8(byte[] json)
        {
            if (json == null)
            {
                throw new ArgumentNullException("json");
            }

            return Deserialize(Utf8WithoutBom.GetString(json));
        }

        private static JsonSerializerSettings CreateSettings()
        {
            return new JsonSerializerSettings
            {
                Culture = CultureInfo.InvariantCulture,
                Formatting = Formatting.Indented,
                DateFormatHandling = DateFormatHandling.IsoDateFormat,
                DateFormatString = "yyyy-MM-dd'T'HH:mm:ss.fffK",
                DateParseHandling = DateParseHandling.DateTimeOffset,
                DateTimeZoneHandling = DateTimeZoneHandling.RoundtripKind,
                FloatParseHandling = FloatParseHandling.Double,
                NullValueHandling = NullValueHandling.Ignore
            };
        }
    }
}
