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

        public HashTagsViewModel(IAddIn addin)
        {
            _addin = addin;
            HashTags = new ObservableCollection<HashTagParagraphs>();
            UpdateTagsCommand = new DelegateCommand(UpdateTags);
        }

        private void UpdateTags()
        {
            _addin.UpdateTags();
        }

        public string Search
        {
            get => _search;
            set => SetProperty(ref _search,value);
        }

        public ObservableCollection<HashTagParagraphs> HashTags
        {
            get => _hashTags;
            set => SetProperty(ref _hashTags,value);
        }

        public DelegateCommand UpdateTagsCommand { get; set; }

        public void AddTag(HashTagItem hashTag)
        {
            var tag = HashTags.FirstOrDefault(p => p.Name.ToLowerInvariant() == hashTag.Name.ToLowerInvariant());
            if (tag == null)
            {
                tag = new HashTagParagraphs()
                {
                    Name = hashTag.Name,
                };
                HashTags.Add(tag);
            }
            tag.Paragraphs.Add(hashTag.Paragraph);
            tag.Count = tag.Paragraphs.Count;

        }

        
    }
}