using Microsoft.Extensions.Configuration;
using Text_to_Image.Models;

namespace Text_to_Image.Services
{
    public class SpeechServiceFactory
    {
        private readonly IConfiguration _configuration;

        public SpeechServiceFactory(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public ISpeechService CreateSpeechService()
        {
            // Đọc từ appsettings.json xem dùng service nào
            string serviceType = _configuration["SpeechService:Type"]?.ToLower() ?? "azure";

            return serviceType switch
            {
                "azure" => CreateAzureService(),
                "google" => CreateGoogleService(),
                _ => throw new ArgumentException($"Unsupported speech service type: {serviceType}")
            };
        }

        public ISpeechService CreateSpeechService(SpeechServiceType serviceType)
        {
            return serviceType switch
            {
                SpeechServiceType.Azure => CreateAzureService(),
                SpeechServiceType.Google => CreateGoogleService(),
                _ => throw new ArgumentException($"Unsupported speech service type: {serviceType}")
            };
        }

        public ISpeechService CreateSpeechService(ProcessingOptions options)
        {
            // Cho phép override từ ProcessingOptions
            if (!string.IsNullOrEmpty(options.PreferredSpeechService))
            {
                return options.PreferredSpeechService.ToLower() switch
                {
                    "google" => CreateGoogleService(),
                    "azure" => CreateAzureService(),
                    _ => CreateSpeechService(options.SpeechServiceType)
                };
            }

            return CreateSpeechService(options.SpeechServiceType);
        }

        private ISpeechService CreateAzureService()
        {
            string speechKey = _configuration["AzureSpeech:Key"];
            string speechRegion = _configuration["AzureSpeech:Region"];

            if (string.IsNullOrEmpty(speechKey) || string.IsNullOrEmpty(speechRegion))
            {
                throw new InvalidOperationException("Azure Speech configuration is missing. Please check AzureSpeech:Key and AzureSpeech:Region in appsettings.json");
            }

            Console.WriteLine("🔵 Using AZURE Text-to-Speech Service");
            return new AzureSpeechService(speechKey, speechRegion);
        }

        private ISpeechService CreateGoogleService()
        {
            string credentialsPath = _configuration["GoogleCloud:CredentialsPath"];

            if (string.IsNullOrEmpty(credentialsPath))
            {
                throw new InvalidOperationException("Google Cloud configuration is missing. Please check GoogleCloud:CredentialsPath in appsettings.json");
            }

            if (!File.Exists(credentialsPath))
            {
                throw new FileNotFoundException($"Google Cloud credentials file not found at: {credentialsPath}");
            }

            Console.WriteLine("🟢 Using GOOGLE Text-to-Speech Service");
            return new GoogleSpeechService(credentialsPath);
        }
    }
}