using System.Collections.Generic;

namespace WordHashTagAddIn
{
    public class HashTag
    {
        public string Name { get; set; }
        public int Count { get; set; }
        public IEnumerable<string> Paragraphs { get; set; }
    }
}