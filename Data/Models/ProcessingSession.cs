using System;
using System.ComponentModel.DataAnnotations;

namespace Text_to_Image.Data.Models
{
    public class ProcessingSession
    {
        [Key]
        public int SessionId { get; set; }

        [MaxLength(255)]
        public string? FileName { get; set; }

        [MaxLength(50)]
        public string? FileType { get; set; } // English, Japanese, Chinese, TuVung

        public DateTime ProcessedDate { get; set; }

        public int TotalRows { get; set; }

        public int ProcessedRows { get; set; }

        public int AudioFilesCreated { get; set; }

        [MaxLength(20)]
        public string? DateUsed { get; set; }

        [MaxLength(1000)]
        public string? Notes { get; set; }

        public bool IsCompleted { get; set; } = false;
    }
}