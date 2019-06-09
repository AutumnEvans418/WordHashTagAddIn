using System.Collections.ObjectModel;
using System.ComponentModel;

namespace WordHashTagAddIn
{
    public interface IAddIn
    {
        void UpdateTags();
        void NavigateToParagraph(HashTagItem selectedParagraph);
    }
    public interface IHashTagsViewModel : INotifyPropertyChanged
    {
        string Search { get; set; }
        bool IsHighlightingTags { get; set; }
        void AddTag(HashTagItem hashTag);
        void ClearTags();
    }
}