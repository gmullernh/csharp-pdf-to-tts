using System;
using System.IO;
using System.Speech.Synthesis;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace PDFTTS
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                SpeechSynthesizer speech = new SpeechSynthesizer();

                // Configurações Voz
                var voices = speech.GetInstalledVoices();

                Console.WriteLine("Vozes instaladas no sistema.");
                Console.WriteLine("ID\tNome\tIdioma");
                for (int i = 0; i < voices.Count; i++)
                {
                    Console.WriteLine($"{i}\t{voices[i].VoiceInfo.Name}\t({voices[i].VoiceInfo.Culture})");
                }

                Console.Write("Digite o ID para escolher a voz: ");
                int.TryParse(Console.ReadLine(), out int voiceId);
                Console.WriteLine($"Voz selecionada: {voices[voiceId].VoiceInfo.Name}");
                speech.SelectVoice(voices[voiceId].VoiceInfo.Name);

                Console.WriteLine(new string('-', 20));

                Console.Write("Velocidade (-10 até 10): ");
                int.TryParse(Console.ReadLine(), out int speed);
                Console.WriteLine($"Velocidade selecionada: {speed}");

                Console.WriteLine(new string('-', 20));

                Console.Write("Volume (0 até 200): ");
                int.TryParse(Console.ReadLine(), out int volume);
                Console.WriteLine($"Volume selecionado: {volume}");

                speech.Rate = speed;
                speech.Volume = volume;

                Console.WriteLine(new string('-', 20));

                // Configurações PDF

                string path = string.Empty;

                do
                {
                    Console.Write("Digite o caminho do arquivo: ");
                    path = Console.ReadLine();
                } while (string.IsNullOrEmpty(path) || !File.Exists(path));

                var pdfDocument = new PdfDocument(new PdfReader(path));
                int totalPages = pdfDocument.GetNumberOfPages();
                Console.WriteLine($"O PDF tem {totalPages} páginas.");

                Console.WriteLine(new string('-', 20));

                int pageStart;
                do
                {
                    Console.Write("Digite a página inicial: ");
                    int.TryParse(Console.ReadLine(), out pageStart);
                } while (pageStart < 1);

                Console.WriteLine(new string('-', 20));

                int pageEnd;
                do
                {
                    Console.Write("Digite a página final: ");
                    int.TryParse(Console.ReadLine(), out pageEnd);
                } while (pageEnd > totalPages || pageEnd < pageStart);

                Console.WriteLine(new string('-', 20));
                Console.WriteLine("Iniciando a leitura.");

                // Itera e lê as páginas.
                int curPage = pageStart;

                do
                {
                    var page = pdfDocument.GetPage(curPage);
                    if (page == null)
                    {
                        Console.WriteLine($"Pulando: {curPage}");
                        curPage++;
                        continue;
                    }

                    string text = ExtractTextFromPdf(page);
                    Console.WriteLine(new string('-', 20));
                    Console.WriteLine($"Página: {curPage}");
                    Console.WriteLine(text);
                    Console.WriteLine(new string('-', 20));

                    ReadText(text, speech);
                    curPage++;
                }
                while (curPage <= pageEnd);

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.ReadLine();
            }

        }

        public static void ReadText(string text, SpeechSynthesizer speech)
        {
            speech.Speak(text);
        }

        public static string ExtractTextFromPdf(PdfPage page)
        {
            var strategy = new LocationTextExtractionStrategy();
            return PdfTextExtractor.GetTextFromPage(page, strategy).Replace("\n", " ").Replace("  ", " ");
        }

    }
}
