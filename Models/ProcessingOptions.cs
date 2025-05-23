using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Text_to_Image.Models
{
    public class ProcessingOptions
    {
        public string SelectedFile { get; set; }
        public string FileName { get; set; }
        public string CustomDate { get; set; }
        public string SelectedDay { get; set; }
        public string SelectedMonth { get; set; }
        public string SelectedYear { get; set; }
        public string ColumnInput { get; set; }
        public string SoundColumns { get; set; }
        public string KanjiColumn { get; set; }
        public string KanjiOutputColumn { get; set; }
        public bool RenameAudioFiles { get; set; }
        public string AudioFolderPath { get; set; }

        public ProcessingOptions()
        {
            SelectedYear = DateTime.Now.Year.ToString();
        }
    }

}
