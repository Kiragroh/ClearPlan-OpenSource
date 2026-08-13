namespace ClearPlan
{
    public sealed class FieldNamePreviewViewModel : ViewModelBase
    {
        public int Order { get; set; }
        public int BeamNumber { get; set; }
        public string CurrentId { get; set; }
        public string ExpectedId { get; set; }
        public string CurrentName { get; set; }
        public string SuggestedName { get; set; }
        public bool IdWouldChange { get; set; }
        public bool NameWouldChange { get; set; }
        public bool WouldChange { get; set; }
        public string IdStatus
        {
            get { return IdWouldChange ? "ID ändern" : "ID okay"; }
        }

        public string NameStatus
        {
            get { return NameWouldChange ? "Name ändern" : "Name okay"; }
        }

        public string Status
        {
            get
            {
                if (!WouldChange)
                {
                    return "ID und Name passend";
                }

                if (IdWouldChange && NameWouldChange)
                {
                    return "ID- und Namensvorschlag";
                }

                return IdWouldChange ? "ID-Vorschlag" : "Namensvorschlag";
            }
        }
    }
}
