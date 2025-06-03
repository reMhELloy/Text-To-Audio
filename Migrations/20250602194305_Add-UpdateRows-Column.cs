using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Text_to_Image.Migrations
{
    public partial class AddUpdateRowsColumn : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "UpdatedRows",
                table: "ProcessingSessions",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "UpdatedRows",
                table: "ProcessingSessions");
        }
    }
}
