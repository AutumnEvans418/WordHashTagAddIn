using System.Collections.ObjectModel;
using Prism.Mvvm;

namespace WordHashTagAddIn
{
    public class HashTagParagraphs : BindableBase
    {
        private readonly IAddIn _addIn;
        private ObservableCollection<HashTagItem> _paragraphs;
        private HashTagItem _selectedParagraph;

        public HashTagParagraphs(IAddIn addIn)
        {
            _addIn = addIn;
            Paragraphs = new ObservableCollection<HashTagItem>();
        }
        public string Name { get; set; }

        public int Count { get; set; }

        public ObservableCollection<HashTagItem> Paragraphs
        {
            get => _paragraphs;
            set => SetProperty(ref _paragraphs,value);
        }

        public HashTagItem SelectedParagraph
        {
            get => _selectedParagraph;
            set => SetProperty(ref _selectedParagraph,value, SelectedParagraphChanged);
        }

        private void SelectedParagraphChanged()
        {
            if(SelectedParagraph != null)
                _addIn.NavigateToParagraph(SelectedParagraph);
        }
    }
}