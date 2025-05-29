using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Text_to_Image.Data.Models
{
    public class Vocabulary
    {
        [Key]
        public int VocabId { get; set; }

        [MaxLength(500)]
        public string? EnglishText { get; set; }

        [MaxLength(500)]
        public string? VietnameseText { get; set; }

        [MaxLength(500)]
        public string? JapaneseText { get; set; }

        [MaxLength(500)]
        public string? ChineseText { get; set; }

        [MaxLength(500)]
        public string? KanjiText { get; set; }

        [MaxLength(500)]
        public string? ReadingText { get; set; }

        [MaxLength(1000)]
        public string? ImageTags { get; set; }

        [MaxLength(100)]
        public string? Category { get; set; }

        [MaxLength(50)]
        public string? SourceFile { get; set; }

        public DateTime CreatedDate { get; set; }

        public DateTime? LastModified { get; set; }

        public bool IsActive { get; set; } = true;
    }
}