using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Text_to_Image.Models;

namespace Text_to_Image.Services
{
    public interface ISpeechService
    {
        Task<bool> CreateAudioFile(string text, string outputPath, string voiceName);
        Task ProcessExcelForAudio(ProcessingOptions options);
        string GetSSMLPreview(string text, string voiceName);
    }

    public enum SpeechServiceType
    {
        Azure,
        Google,
        GoogleTranslate
    }
}
