using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ClearPlan.Core.PlanAnalysis;
using VMS.TPS.Common.Model.API;

namespace ClearPlan.Review
{
    /// <summary>Read native positions on the ESAPI owner STA, then commit a complete detached beam.</summary>
    public static class EsapiNativeMlcAdapter
    {
        public static void ApplyToBeam(Beam beam, ReviewBeamAnalysis row, NativeMlcProfileCatalog catalog)
        {
            if (beam == null) throw new ArgumentNullException("beam");
            if (row == null) throw new ArgumentNullException("row");
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Native ESAPI geometry must be read on its owner STA thread.");
            try
            {
                if (beam.MLC == null) return; // Existing explicit jaw-only extraction remains authoritative.
                string model = beam.MLC.Model;
                var points = beam.ControlPoints.ToList();
                if (row.ControlPoints == null || points.Count == 0 || points.Count > 20000 ||
                    points.Count != row.ControlPoints.Count || !row.BeamNumber.HasValue || row.BeamNumber.Value != beam.BeamNumber ||
                    points.Where((point, index) => point.Index != row.ControlPoints[index].Index).Any())
                {
                    Invalidate(row, "snapshot_mismatch", "Native beam/control-point identity differs from the detached snapshot. Refresh the plan first.");
                    return;
                }
                if (beam.Applicator != null || beam.Blocks.Any())
                {
                    Invalidate(row, "unsupported_accessory", "An applicator or blocking accessory is not represented by the native MLC profile.");
                    return;
                }
                var pending = new List<ApertureGeometry>();
                foreach (var cp in points)
                {
                    // No native Beam, ControlPoint or position array is retained or sent to a worker.
                    float[,] nativePositions = cp.LeafPositions;
                    var nativeJaws = cp.JawPositions;
                    var mapped = NativeMlcGeometryMapper.Map(catalog, model, nativePositions,
                        new ApertureRectangle(nativeJaws.X1, nativeJaws.Y1, nativeJaws.X2, nativeJaws.Y2));
                    if (mapped.Geometry == null)
                    {
                        Invalidate(row, mapped.Code, SafeMachineDescription(model) + "; " + mapped.Reason);
                        return;
                    }
                    // A profile cannot turn a known jawless dual-layer machine into a C-arm beam.
                    if (row.HasJaws == false && mapped.Geometry.Jaws != null ||
                        row.MlcLayerCount == 2 && mapped.Geometry.Layers.Count != 2)
                    {
                        Invalidate(row, "machine_profile_mismatch", "The explicit native profile conflicts with the beam's known jawless or dual-layer hardware.");
                        return;
                    }
                    if (pending.Count > 0 && !SameFixedLimits(pending[0].FixedBoundingBox, mapped.Geometry.FixedBoundingBox))
                    {
                        Invalidate(row, "fixed_limits_changed", "A profile declares fixed field limits, but native limits change between control points.");
                        return;
                    }
                    pending.Add(mapped.Geometry);
                }
                CommitCompleteBeam(row, pending);
            }
            catch (Exception)
            {
                // Vendor exception text can disclose patient data. Never propagate its values.
                Invalidate(row, "native_read", "Native MLC geometry could not be read completely on the ESAPI owner thread.");
            }
        }

        private static void CommitCompleteBeam(ReviewBeamAnalysis row, List<ApertureGeometry> pending)
        {
            for (int i = 0; i < pending.Count; i++)
            {
                row.ControlPoints[i].Aperture = pending[i];
                row.ControlPoints[i].GeometryReason = null;
            }
            row.MlcLayerCount = pending[0].Layers.Count;
            row.HasJaws = pending[0].Jaws != null;
            row.GeometryStatus = "available";
            row.GeometryReason = null;
            // The caller recalculates plan metrics from the newly complete detached apertures.
        }
        private static void Invalidate(ReviewBeamAnalysis row, string code, string reason)
        {
            row.GeometryStatus = "unavailable"; row.GeometryReason = code + ": " + reason;
            if (row.ControlPoints == null) return;
            foreach (var point in row.ControlPoints)
            {
                point.Aperture = null;
                point.GeometryReason = row.GeometryReason;
            }
        }
        private static bool SameFixedLimits(ApertureRectangle a, ApertureRectangle b)
        {
            if (a == null || b == null) return a == null && b == null;
            return Math.Abs(a.X1 - b.X1) <= 1e-6 && Math.Abs(a.X2 - b.X2) <= 1e-6 &&
                Math.Abs(a.Y1 - b.Y1) <= 1e-6 && Math.Abs(a.Y2 - b.Y2) <= 1e-6;
        }
        private static string SafeMachineDescription(string model)
        {
            // This is machine configuration, not a patient/beam identifier. Bound and sanitize it.
            if (string.IsNullOrWhiteSpace(model)) return "Native MLC model unavailable";
            return "Native MLC model " + new string(model.Where(c => !char.IsControl(c)).Take(128).ToArray());
        }
    }
}
