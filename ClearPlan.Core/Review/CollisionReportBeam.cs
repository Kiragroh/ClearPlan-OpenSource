using System.Collections.Generic;
using System;
using System.Linq;
using System.Globalization;
using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    // Private, detached report projection. Not part of distributable review JSON.
    public sealed class CollisionReportBeam
    {
        public string BeamId { get; set; }
        public string Status { get; set; }
        public bool RadialOnly { get; set; }
        public bool BodyOnly { get; set; }
        public string Summary { get; set; }
        public byte[] OverviewPng { get; set; }
        // Captured with OverviewPng at MinimumRow; absent in older private sidecars.
        public byte[] OrientationPng { get; set; }
        public List<CollisionReportRow> Rows { get; set; } = new List<CollisionReportRow>();
        // Nullable for older sidecars: do not infer a minimum from formatted strings.
        [JsonIgnore] public CollisionReportRow MinimumRow { get { return Rows.Where(r => r.MinimumDistanceMm.HasValue &&
            !double.IsNaN(r.MinimumDistanceMm.Value) && !double.IsInfinity(r.MinimumDistanceMm.Value))
            .OrderBy(r => r.MinimumDistanceMm.Value).FirstOrDefault(); } }
        public static CollisionReportBeam MinimumBeam(IEnumerable<CollisionReportBeam> beams)
        { return beams.Where(b => b != null && b.MinimumRow != null).OrderBy(b => b.MinimumRow.MinimumDistanceMm.Value).FirstOrDefault(); }
        [JsonIgnore] public string MinimumCaption { get { var r = MinimumRow; return r == null ? "Minimum model distance unavailable."
            : string.Format(CultureInfo.InvariantCulture, "Minimum model distance: {0:0.0} mm · {1} · Gantry {2:0.#}° · Couch {3:0.#}° · {4}",
                r.MinimumDistanceMm, BeamId, r.GantryDegrees, r.CouchDegrees, RadialOnly ? "radial gap" : "conservative lower bound"); } }
    }
    public sealed class CollisionReportRow
    {
        public double GantryDegrees { get; set; }
        public double CouchDegrees { get; set; }
        public double? MinimumDistanceMm { get; set; }
        public int? CapturedControlPointIndex { get; set; }
        public bool Interpolated { get; set; }
        public string Status { get; set; }
        public string BodyDistance { get; set; }
        public string TableDistance { get; set; }
        public string BodyStatus { get; set; }
        public string TableStatus { get; set; }
        public string Reason { get; set; }
    }
}
