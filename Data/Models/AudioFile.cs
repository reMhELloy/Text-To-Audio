using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Text_to_Image.Data.Models
{
    public class AudioFile
    {
        [Key]
        public int AudioId { get; set; }

        public int VocabId { get; set; }

        [ForeignKey("VocabId")]
        public virtual Vocabulary Vocabulary { get; set; }

        [MaxLength(10)]
        public string? Language { get; set; } // EN, JP, ZH, VI

        [MaxLength(255)]
        public string? FileName { get; set; }

        [MaxLength(500)]
        public string? FilePath { get; set; }

        public long FileSize { get; set; }

        public int Duration { get; set; } // in seconds

        [MaxLength(100)]
        public string? VoiceName { get; set; }

        [Column(TypeName = "decimal(3,2)")]
        public decimal SpeechRate { get; set; }

        public bool IsOddFile { get; set; }

        public DateTime CreatedDate { get; set; }

        public bool IsGenerated { get; set; } = false;
    }
}