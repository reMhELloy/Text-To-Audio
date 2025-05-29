using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Text_to_Image.Migrations
{
    public partial class AllowNullImageTags : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "ProcessingSessions",
                columns: table => new
                {
                    SessionId = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    FileName = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: false),
                    FileType = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: false),
                    ProcessedDate = table.Column<DateTime>(type: "datetime2", nullable: false),
                    TotalRows = table.Column<int>(type: "int", nullable: false),
                    ProcessedRows = table.Column<int>(type: "int", nullable: false),
                    AudioFilesCreated = table.Column<int>(type: "int", nullable: false),
                    DateUsed = table.Column<string>(type: "nvarchar(20)", maxLength: 20, nullable: false),
                    Notes = table.Column<string>(type: "nvarchar(1000)", maxLength: 1000, nullable: false),
                    IsCompleted = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ProcessingSessions", x => x.SessionId);
                });

            migrationBuilder.CreateTable(
                name: "Vocabularies",
                columns: table => new
                {
                    VocabId = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    EnglishText = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: false),
                    VietnameseText = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: false),
                    JapaneseText = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: false),
                    ChineseText = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: false),
                    KanjiText = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: false),
                    ReadingText = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: false),
                    ImageTags = table.Column<string>(type: "nvarchar(1000)", maxLength: 1000, nullable: true),
                    Category = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: false),
                    SourceFile = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: false),
                    CreatedDate = table.Column<DateTime>(type: "datetime2", nullable: false),
                    LastModified = table.Column<DateTime>(type: "datetime2", nullable: true),
                    IsActive = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Vocabularies", x => x.VocabId);
                });

            migrationBuilder.CreateTable(
                name: "AudioFiles",
                columns: table => new
                {
                    AudioId = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    VocabId = table.Column<int>(type: "int", nullable: false),
                    Language = table.Column<string>(type: "nvarchar(10)", maxLength: 10, nullable: false),
                    FileName = table.Column<string>(type: "nvarchar(255)", maxLength: 255, nullable: false),
                    FilePath = table.Column<string>(type: "nvarchar(500)", maxLength: 500, nullable: false),
                    FileSize = table.Column<long>(type: "bigint", nullable: false),
                    Duration = table.Column<int>(type: "int", nullable: false),
                    VoiceName = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: false),
                    SpeechRate = table.Column<decimal>(type: "decimal(3,2)", nullable: false),
                    IsOddFile = table.Column<bool>(type: "bit", nullable: false),
                    CreatedDate = table.Column<DateTime>(type: "datetime2", nullable: false),
                    IsGenerated = table.Column<bool>(type: "bit", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_AudioFiles", x => x.AudioId);
                    table.ForeignKey(
                        name: "FK_AudioFiles_Vocabularies_VocabId",
                        column: x => x.VocabId,
                        principalTable: "Vocabularies",
                        principalColumn: "VocabId",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_AudioFiles_FileName",
                table: "AudioFiles",
                column: "FileName");

            migrationBuilder.CreateIndex(
                name: "IX_AudioFiles_VocabId_Language",
                table: "AudioFiles",
                columns: new[] { "VocabId", "Language" });

            migrationBuilder.CreateIndex(
                name: "IX_ProcessingSessions_FileName",
                table: "ProcessingSessions",
                column: "FileName");

            migrationBuilder.CreateIndex(
                name: "IX_ProcessingSessions_ProcessedDate",
                table: "ProcessingSessions",
                column: "ProcessedDate");

            migrationBuilder.CreateIndex(
                name: "IX_Vocabularies_CreatedDate",
                table: "Vocabularies",
                column: "CreatedDate");

            migrationBuilder.CreateIndex(
                name: "IX_Vocabularies_EnglishText",
                table: "Vocabularies",
                column: "EnglishText");

            migrationBuilder.CreateIndex(
                name: "IX_Vocabularies_SourceFile",
                table: "Vocabularies",
                column: "SourceFile");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "AudioFiles");

            migrationBuilder.DropTable(
                name: "ProcessingSessions");

            migrationBuilder.DropTable(
                name: "Vocabularies");
        }
    }
}
