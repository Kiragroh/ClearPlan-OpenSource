namespace ClearPlan.Core.Fields
{
    public sealed class BeamNamingInput
    {
        public string CurrentId { get; set; }
        public string CurrentName { get; set; }
        public int BeamNumber { get; set; }
        public int? TreatmentOrderIndex { get; set; }
        public double GantryStartAngle { get; set; }
        public double GantryStopAngle { get; set; }
        public double PatientSupportAngle { get; set; }
        public string GantryDirection { get; set; }

        public static BeamNamingInput Static(
            string currentId,
            int beamNumber,
            double gantryAngle)
        {
            return new BeamNamingInput
            {
                CurrentId = currentId,
                BeamNumber = beamNumber,
                GantryStartAngle = gantryAngle,
                GantryStopAngle = gantryAngle,
                GantryDirection = string.Empty
            };
        }

        public static BeamNamingInput Arc(
            string currentId,
            int beamNumber,
            double gantryStartAngle,
            double gantryStopAngle,
            string gantryDirection)
        {
            return new BeamNamingInput
            {
                CurrentId = currentId,
                BeamNumber = beamNumber,
                GantryStartAngle = gantryStartAngle,
                GantryStopAngle = gantryStopAngle,
                GantryDirection = gantryDirection
            };
        }
    }
}
