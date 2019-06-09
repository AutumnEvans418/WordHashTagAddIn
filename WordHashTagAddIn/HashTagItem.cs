using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace WordHashTagAddIn
{
    public class HashTagItem
    {
        public string Name { get; set; }
        public string Paragraph { get; set; }
    }
    public class HashTagParagraphs
    {
        public HashTagParagraphs()
        {
            Paragraphs = new ObservableCollection<string>();
        }
        public string Name { get; set; }

        public int Count { get; set; }

        public ObservableCollection<string> Paragraphs { get; set; }
    }
}