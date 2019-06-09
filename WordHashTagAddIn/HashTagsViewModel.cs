using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using WordHashTagAddIn.Annotations;

namespace WordHashTagAddIn
{
    public class HashTagsViewModel : IHashTagsViewModel
    {
        private ObservableCollection<HashTag> _hashTags;
        private string _search;
        public event PropertyChangedEventHandler PropertyChanged;

        public string Search
        {
            get => _search;
            set => SetProperty(ref _search,value);
        }

        public ObservableCollection<HashTag> HashTags
        {
            get => _hashTags;
            set => SetProperty(ref _hashTags,value);
        }

        protected void SetProperty<T>(ref T field, T value)
        {
            if (field.Equals(value) != true)
            {
                field = value;
            }
        }
        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}