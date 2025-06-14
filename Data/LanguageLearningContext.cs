using Microsoft.EntityFrameworkCore;
using Text_to_Image.Data.Models;

namespace Text_to_Image.Data
{
    public class LanguageLearningContext : DbContext
    {
        public DbSet<Vocabulary> Vocabularies { get; set; }
        public DbSet<AudioFile> AudioFiles { get; set; }
        public DbSet<ProcessingSession> ProcessingSessions { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // Connection string cho SQL Server của bạn
            optionsBuilder.UseSqlServer(
                //"Server=REM\\MSSQLSERVER_MEGA;Database=LanguageLearningDB;Integrated Security=true;TrustServerCertificate=true;"
                "Server=REM\\MSSQLSERVER_MEGA;Database=LanguageLearningDB;User Id=sa;Password=1;TrustServerCertificate=true;"
            );
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // Vocabulary indexes để tối ưu hóa performance
            modelBuilder.Entity<Vocabulary>()
                .HasIndex(v => v.EnglishText);

            modelBuilder.Entity<Vocabulary>()
                .HasIndex(v => v.SourceFile);

            modelBuilder.Entity<Vocabulary>()
                .HasIndex(v => v.CreatedDate);

            // AudioFile indexes
            modelBuilder.Entity<AudioFile>()
                .HasIndex(a => new { a.VocabId, a.Language });

            modelBuilder.Entity<AudioFile>()
                .HasIndex(a => a.FileName);

            // ProcessingSession indexes
            modelBuilder.Entity<ProcessingSession>()
                .HasIndex(p => p.ProcessedDate);

            modelBuilder.Entity<ProcessingSession>()
                .HasIndex(p => p.FileName);
        }
    }
}