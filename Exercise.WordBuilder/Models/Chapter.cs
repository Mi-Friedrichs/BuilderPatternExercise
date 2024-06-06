using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exercise.WordBuilder.Models
{
    internal class Chapter
    {
        public Guid Id { get; set; } = Guid.NewGuid();

        public string Header { get; set; } = "";

        public string? Teaser { get; set; }


        private List<string> paragraphs;
        public List<string> Paragraphs
        {
            get
            {
                paragraphs ??= new List<string>();
                return paragraphs;
            }
            set { paragraphs = value; }
        }


        public int Pages { get; set; } = 1;

    }
}
