using VMS.TPS.Common.Model.API;

namespace ClearPlan
{
    public sealed class DvhStructureViewModel : ViewModelBase
    {
        private bool _isSelected;

        public DvhStructureViewModel(Structure structure)
        {
            Structure = structure;
            Id = structure == null ? string.Empty : structure.Id;
        }

        public Structure Structure { get; private set; }

        public string Id { get; private set; }

        public bool RequiredForTargetReview { get; set; }

        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                value = value || RequiredForTargetReview;
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                NotifyPropertyChanged("IsSelected");
            }
        }
    }
}
