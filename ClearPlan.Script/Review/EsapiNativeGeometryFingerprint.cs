using System;
using System.Linq;
using System.Threading;
using ClearPlan.Core.PlanAnalysis;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Review
{
    /// <summary>
    /// Native read-only state stamps, never model-derived geometry. These detect captured beam
    /// values and plan/image metadata changes; they do NOT prove CT HU or contours unchanged.
    /// Identity strings enter the in-memory hash only and are neither returned nor persisted.
    /// </summary>
    internal static class EsapiNativeGeometryFingerprint
    {
        internal static string CaptureBeam(Beam beam,CancellationToken token=default(CancellationToken))
        {
            VerifyOwner();token.ThrowIfCancellationRequested();
            if(beam==null) throw new ArgumentNullException("beam");
            using(var stamp=new GeometryFingerprintWriter())
            {
                stamp.Add("beam-v2");stamp.Add(beam.BeamNumber);
                stamp.Add(beam.WeightFactor);stamp.Add(beam.IsImagingTreatmentField ? 1 : 0);
                var meterset=beam.Meterset;
                stamp.Add((int)meterset.Unit);stamp.Add(meterset.Value);stamp.Add(beam.DoseRate);
                stamp.Add(beam.GantryDirection.ToString());
                AddPoint(stamp,beam.IsocenterPosition);
                var mlc=beam.MLC;
                stamp.Add(mlc==null ? null : mlc.Model);
                var unit=beam.TreatmentUnit;
                stamp.Add(unit==null ? null : unit.Id);
                stamp.Add(unit==null ? null : unit.MachineModelName);
                stamp.Add(unit==null ? null : unit.MachineModel);
                stamp.Add(beam.EnergyModeDisplayName);
                stamp.Add(beam.Technique==null ? null : beam.Technique.Id);
                // Presence changes can invalidate an otherwise unchanged supported aperture.
                stamp.Add(beam.Applicator==null ? 0 : 1);
                stamp.Add(beam.Blocks.Any() ? 1 : 0);
                var cps=beam.ControlPoints.Take(4097).ToList();
                if(cps.Count<1 || cps.Count>4096) throw new ArgumentException("Native control-point count exceeds the fingerprint bound.");
                stamp.Add(cps.Count);
                foreach(var cp in cps)
                {
                    token.ThrowIfCancellationRequested();
                    stamp.Add(cp.Index);stamp.Add(cp.GantryAngle);stamp.Add(cp.CollimatorAngle);
                    stamp.Add(cp.PatientSupportAngle);stamp.Add(cp.MetersetWeight);
                    stamp.AddOptional(cp.TableTopLateralPosition);stamp.AddOptional(cp.TableTopLongitudinalPosition);
                    stamp.AddOptional(cp.TableTopVerticalPosition);
                    var jaws=cp.JawPositions;
                    stamp.Add(jaws.X1);stamp.Add(jaws.Y1);stamp.Add(jaws.X2);stamp.Add(jaws.Y2);
                    var leaves=cp.LeafPositions;
                    if(mlc!=null && leaves==null) throw new ArgumentException("Native MLC positions are unavailable for a fingerprint.");
                    stamp.Add(leaves);
                }
                return stamp.Complete();
            }
        }

        internal static string CapturePlan(PlanSetup plan,CancellationToken token=default(CancellationToken))
        {
            VerifyOwner();token.ThrowIfCancellationRequested();
            if(plan==null) throw new ArgumentNullException("plan");
            using(var stamp=new GeometryFingerprintWriter())
            {
                stamp.Add("plan-v1");stamp.Add(plan.PlanNormalizationValue);stamp.Add((int)plan.TreatmentOrientation);
                var fraction=plan.DosePerFraction;
                stamp.Add(fraction.Dose);stamp.Add((int)fraction.Unit);stamp.Add(plan.TargetVolumeID);
                var structures=plan.StructureSet;
                stamp.Add(structures==null ? null : structures.UID);
                var image=structures==null ? null : structures.Image;
                stamp.Add(image==null ? "no-image" : "image");
                if(image!=null)
                {
                    stamp.Add(image.UID);stamp.Add(image.FOR);stamp.Add(image.DisplayUnit);
                    stamp.Add(image.XSize);stamp.Add(image.YSize);stamp.Add(image.ZSize);
                    stamp.Add(image.XRes);stamp.Add(image.YRes);stamp.Add(image.ZRes);
                    AddPoint(stamp,image.Origin);AddPoint(stamp,image.XDirection);
                    AddPoint(stamp,image.YDirection);AddPoint(stamp,image.ZDirection);
                }
                return stamp.Complete();
            }
        }

        private static void AddPoint(GeometryFingerprintWriter stamp,VVector point)
        { stamp.Add(point.x);stamp.Add(point.y);stamp.Add(point.z); }
        private static void VerifyOwner()
        {
            if(Thread.CurrentThread.GetApartmentState()!=ApartmentState.STA)
                throw new InvalidOperationException("Native fingerprint capture requires the ESAPI STA owner.");
            if(System.Windows.Application.Current!=null) System.Windows.Application.Current.Dispatcher.VerifyAccess();
        }
    }
}
