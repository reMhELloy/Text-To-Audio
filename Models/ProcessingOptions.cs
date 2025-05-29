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

        public string AudioOutputFolder { get; set; }
        public bool CreateAudioFiles { get; set; } = false;

        public bool SaveToDatabase { get; set; } = false;

        // Audio source columns
        public string VietnameseColumn { get; set; }
        public string EnglishColumn { get; set; }
        public string JapaneseColumn { get; set; }
        public string ChineseColumn { get; set; }

        // Audio file type để xác định logic xử lý
        public string AudioFileType { get; set; }


        public ProcessingOptions()
        {
            SelectedYear = DateTime.Now.Year.ToString();
        }
    }

}
