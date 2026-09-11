using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace ClearPlan.Core.PlanAnalysis
{
    public static class PlanAnalysisSnapshot
    {
        public static ReviewPlanAnalysis Copy(ReviewPlanAnalysis source)
        {
            if(source==null) throw new ArgumentNullException("source");
            // In-memory only. Public scalar DTO members stay future-compatible; sensitive
            // rendering geometry excluded from JSON is explicitly copied below.
            var copy=JsonConvert.DeserializeObject<ReviewPlanAnalysis>(JsonConvert.SerializeObject(source));
            copy.NativePlanFingerprint=source.NativePlanFingerprint;
            for(int b=0;b<source.Beams.Count;b++) copy.Beams[b].NativeGeometryFingerprint=source.Beams[b].NativeGeometryFingerprint;
            for(int b=0;b<source.Beams.Count;b++) for(int c=0;c<source.Beams[b].ControlPoints.Count;c++)
            {
                var from=source.Beams[b].ControlPoints[c]; var to=copy.Beams[b].ControlPoints[c];
                to.IsocenterMm=from.IsocenterMm==null ? null : (double[])from.IsocenterMm.Clone();
                if (from.BevImage != null)
                {
                    to.BevImage=JsonConvert.DeserializeObject<BeamEyeViewImage>(JsonConvert.SerializeObject(from.BevImage));
                    to.BevImage.GrayscalePixels=from.BevImage.GrayscalePixels == null ? null : (byte[])from.BevImage.GrayscalePixels.Clone();
                }
                to.TargetOutlines=from.TargetOutlines==null ? new List<List<BeamPoint>>() : from.TargetOutlines
                    .Select(l=>l.Select(p=>new BeamPoint(p.X,p.Y)).ToList()).ToList();
                to.TargetProjectionStrips=from.TargetProjectionStrips==null ? new List<ApertureRectangle>() :
                    from.TargetProjectionStrips.Select(Rectangle).ToList();
                if(from.Aperture!=null) to.Aperture=new ApertureGeometry
                {
                    Jaws=from.Aperture.Jaws==null ? null : Rectangle(from.Aperture.Jaws),
                    FixedBoundingBox=from.Aperture.FixedBoundingBox==null ? null : Rectangle(from.Aperture.FixedBoundingBox),
                    Layers=(from.Aperture.Layers ?? new List<ApertureLayer>()).Select(l=>new ApertureLayer {
                        Label=l.Label,LeafTravelAxis=l.LeafTravelAxis,
                        LeafBoundariesMm=l.LeafBoundariesMm==null ? null : (double[])l.LeafBoundariesMm.Clone(),
                        Bank1PositionsMm=l.Bank1PositionsMm==null ? null : (double[])l.Bank1PositionsMm.Clone(),
                        Bank2PositionsMm=l.Bank2PositionsMm==null ? null : (double[])l.Bank2PositionsMm.Clone() }).ToList(),
                    EffectiveOpenings=(from.Aperture.EffectiveOpenings ?? new List<ApertureRectangle>()).Select(Rectangle).ToList()
                };
            }
            return copy;
        }
        private static ApertureRectangle Rectangle(ApertureRectangle r) { return new ApertureRectangle(r.X1,r.Y1,r.X2,r.Y2); }
    }
}
