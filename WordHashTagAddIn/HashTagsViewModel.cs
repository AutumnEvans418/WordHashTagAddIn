using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Prism.Commands;
using Prism.Mvvm;
using WordHashTagAddIn.Annotations;

namespace WordHashTagAddIn
{
    public class HashTagsViewModel : BindableBase, IHashTagsViewModel
    {
        private readonly IAddIn _addin;
        private ObservableCollection<HashTagParagraphs> _hashTags;
        private string _search;
        private bool _isHighlightingTags;

        public HashTagsViewModel(IAddIn addin)
        {
            _addin = addin;
            HashTags = new ObservableCollection<HashTagParagraphs>();
            _allHashTags = new ObservableCollection<HashTagParagraphs>();
            UpdateTagsCommand = new DelegateCommand(UpdateTags);
        }

        private void UpdateTags()
        {
            _addin.UpdateTags();
        }

        public string Search
        {
            get => _search;
            set => SetProperty(ref _search, value, SearchChanged);
        }

        private void SearchChanged()
        {
            if (Search != null)
            {
                HashTags = new ObservableCollection<HashTagParagraphs>(_allHashTags.Where(p=>p.Name.Contains(Search)));
            }
            else
            {
                HashTags = new ObservableCollection<HashTagParagraphs>(_allHashTags);
            }
        }

        private readonly ObservableCollection<HashTagParagraphs> _allHashTags;
        public ObservableCollection<HashTagParagraphs> HashTags
        {
            get => _hashTags;
            set => SetProperty(ref _hashTags, value);
        }

        public bool IsHighlightingTags
        {
            get => _isHighlightingTags;
            set => SetProperty(ref _isHighlightingTags, value);
        }

        public bool IsShowingHashTagsView
        {
            get => Panes.HashTags.Visible;
            set => Panes.HashTags.Visible = value;
        }

        public DelegateCommand UpdateTagsCommand { get; set; }

        public void AddTag(HashTagItem hashTag)
        {
            var tag = HashTags.FirstOrDefault(p => p.Name.ToLowerInvariant() == hashTag.Name.ToLowerInvariant());
            if (tag == null)
            {
                tag = new HashTagParagraphs(_addin)
                {
                    Name = hashTag.Name,
                };
                HashTags.Add(tag);
                _allHashTags.Add(tag);
            }

            if (tag.Paragraphs.Any(p => p.Paragraph == hashTag.Paragraph) != true)
            {
                tag.Paragraphs.Add(hashTag);
                tag.Count = tag.Paragraphs.Count;
            }

        }

        public void ClearTags()
        {
            _allHashTags.Clear();
            HashTags.Clear();
        }
    }
}