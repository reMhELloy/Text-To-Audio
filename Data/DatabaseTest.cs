using Microsoft.EntityFrameworkCore;
using Text_to_Image.Data;

public static class DatabaseTest
{
    public static async Task TestConnection()
    {
        try
        {
            using var context = new LanguageLearningContext();
            await context.Database.EnsureCreatedAsync();
            Console.WriteLine("✅ Database connection successful!");

            var sessionCount = await context.ProcessingSessions.CountAsync();
            Console.WriteLine($"📊 Current sessions in database: {sessionCount}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ Database connection failed: {ex.Message}");
        }
    }
}