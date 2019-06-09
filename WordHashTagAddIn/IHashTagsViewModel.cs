using System.Collections.ObjectModel;
using System.ComponentModel;

namespace WordHashTagAddIn
{
    public interface IAddIn
    {
        void UpdateTags();
    }
    public interface IHashTagsViewModel : INotifyPropertyChanged
    {
        string Search { get; set; }
        ObservableCollection<HashTagParagraphs> HashTags { get; }
        void AddTag(HashTagItem hashTag);
    }
}