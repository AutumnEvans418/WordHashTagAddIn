using System.Collections.ObjectModel;
using System.ComponentModel;

namespace WordHashTagAddIn
{
    public interface IHashTagsViewModel : INotifyPropertyChanged
    {
        string Search { get; set; }
        ObservableCollection<HashTag> HashTags { get; set; }
    }
}