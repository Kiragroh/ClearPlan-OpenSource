using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Collections.Generic;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Reporting;
using Newtonsoft.Json.Linq;

// Explicit protected operator artifact, never the public ReviewSnapshot JSON contract.
internal static class ProtectedReportSidecar
{
    internal static string Hash(byte[] bytes)
    { using(var sha=SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-","").ToLowerInvariant(); }

    internal static void Export(string sourceJsonPath, byte[] sourceJson, ReviewReportDocument document, string patientId)
    {
        var images=new JArray();
        foreach(var beam in document.PlanAnalysis.Beams)
        foreach(var cp in beam.ControlPoints.Where(c=>c.BevImage!=null))
        {
            var image=cp.BevImage;
            bool unavailable=IsExplicitUnavailable(image);
            if(!unavailable && (image.SourceStatus!="available" || image.GrayscalePixels==null || image.WidthPixels<=0 || image.HeightPixels<=0 || image.GrayscalePixels.LongLength!=(long)image.WidthPixels*image.HeightPixels))
                throw new InvalidOperationException("Invalid protected image state.");
            images.Add(new JObject { ["BeamNumber"]=beam.BeamNumber,["ControlPointIndex"]=cp.Index,
                ["Width"]=image.WidthPixels,["Height"]=image.HeightPixels,
                ["ImageMetadata"]=JObject.FromObject(image),["Aperture"]=cp.Aperture==null ? null : JObject.FromObject(cp.Aperture),
                ["IsocenterMm"]=cp.IsocenterMm==null ? null : JArray.FromObject(cp.IsocenterMm),
                ["TargetOutlines"]=JArray.FromObject(cp.TargetOutlines),["TargetProjectionStrips"]=JArray.FromObject(cp.TargetProjectionStrips),
                ["PixelSha256"]=unavailable ? null : Hash(image.GrayscalePixels),["Pixels"]=unavailable ? null : Convert.ToBase64String(image.GrayscalePixels) });
        }
        var data=new JObject { ["SchemaVersion"]=1,["SourceReportSha256"]=Hash(sourceJson),
            ["SourcePatientIdSha256"]=Hash(Encoding.UTF8.GetBytes(patientId)),["ActivePlanKey"]=document.ActivePlanKey,["Images"]=images };
        using(var writer=new StreamWriter(new FileStream(sourceJsonPath+".pixels.json",FileMode.CreateNew,FileAccess.Write))) writer.Write(data.ToString());
    }

    internal static string Apply(string sourceJsonPath,byte[] sourceJson,ReviewReportDocument document,string patientId)
    {
        string path=sourceJsonPath+".pixels.json";
        if(new FileInfo(path).Length>32*1024*1024) throw new InvalidOperationException("Oversized protected sidecar.");
        byte[] bytes=File.ReadAllBytes(path);
        var data=JObject.Parse(Encoding.UTF8.GetString(bytes));
        if((int?)data["SchemaVersion"]!=1 || (string)data["SourceReportSha256"]!=Hash(sourceJson) ||
            (string)data["SourcePatientIdSha256"]!=Hash(Encoding.UTF8.GetBytes(patientId)) ||
            (string)data["ActivePlanKey"]!=document.ActivePlanKey) throw new InvalidOperationException("Protected source binding mismatch.");
        var images=data["Images"] as JArray;
        if(images==null || images.Count==0) throw new InvalidOperationException("Protected DRR pixels missing.");
        foreach(var token in images)
        {
            var beam=document.PlanAnalysis.Beams.Single(b=>b.BeamNumber==(int)token["BeamNumber"]);
            var cp=beam.ControlPoints.Single(c=>c.Index==(int)token["ControlPointIndex"]);
            if(cp.BevImage!=null) throw new InvalidOperationException("Duplicate or prepopulated protected image.");
            var image=token["ImageMetadata"].ToObject<BeamEyeViewImage>();
            if(image==null || image.ControlPointIndex!=cp.Index || image.WidthPixels!=(int)token["Width"] || image.HeightPixels!=(int)token["Height"])
                throw new InvalidOperationException("Protected image binding mismatch.");
            if(IsExplicitUnavailable(image))
            {
                if(token["Pixels"].Type!=JTokenType.Null || token["PixelSha256"].Type!=JTokenType.Null)
                    throw new InvalidOperationException("Unavailable protected image must not carry pixels.");
            }
            else
            {
                if(image.SourceStatus!="available" || image.WidthPixels<=0 || image.HeightPixels<=0 || string.IsNullOrWhiteSpace((string)token["Pixels"]))
                    throw new InvalidOperationException("Protected DRR pixels missing.");
                byte[] pixels=Convert.FromBase64String((string)token["Pixels"]);
                if(pixels.LongLength!=(long)image.WidthPixels*image.HeightPixels || Hash(pixels)!=(string)token["PixelSha256"])
                    throw new InvalidOperationException("Protected image checksum mismatch.");
                image.GrayscalePixels=pixels;
            }
            cp.BevImage=image;
            cp.Aperture=token["Aperture"].ToObject<ApertureGeometry>();
            cp.IsocenterMm=token["IsocenterMm"].ToObject<double[]>();
            cp.TargetOutlines=token["TargetOutlines"].ToObject<List<List<BeamPoint>>>();
            cp.TargetProjectionStrips=token["TargetProjectionStrips"].ToObject<List<ApertureRectangle>>();
        }
        if(document.PlanAnalysis.Beams.Any(b=>b.ControlPoints.Count==0 || b.ControlPoints[0].BevImage==null || (b.ControlPoints[0].BevImage.GrayscalePixels==null && !IsExplicitUnavailable(b.ControlPoints[0].BevImage))))
            throw new InvalidOperationException("A field-start DRR is missing.");
        return Hash(bytes);
    }

    private static bool IsExplicitUnavailable(BeamEyeViewImage image)
    { return image.SourceStatus=="unavailable" && !string.IsNullOrWhiteSpace(image.UnavailableReason) && image.GrayscalePixels==null && image.WidthPixels==0 && image.HeightPixels==0; }
}
