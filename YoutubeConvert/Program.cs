using System;
using System.IO;
using System.Threading.Tasks;
using YoutubeExplode;
using YoutubeExplode.Videos.Streams;

class Program
{
    static async Task Main(string[] args)
    {
        var youtube = new YoutubeClient();

        // Insira o link do vídeo do YouTube que você deseja baixar
        string videoUrl = "https://www.youtube.com/watch?v=BF0uf7apZDQ";

        // Obtém o manifest das streams
        var streamManifest = await youtube.Videos.Streams.GetManifestAsync(videoUrl);

        // Seleciona a melhor stream de áudio (a de maior taxa de bits)
        var audioStreamInfo = streamManifest
            .GetAudioStreams()
            .Where(s => s.Container == Container.Mp4) // Filtro para streams MP4
            .GetWithHighestBitrate();

        if (audioStreamInfo != null)
        {
            // Caminho para o diretório do projeto
            string diretorioDoProjeto = Directory.GetCurrentDirectory();

            // Pasta onde você deseja salvar o arquivo final (no caso, o MP3)
            string pastaDeDestino = diretorioDoProjeto;

            // Nome do arquivo final (MP3)
            string nomeDoArquivoFinal = "video.mp4";

            // Caminho completo para o arquivo MP3
            string caminhoDoMp3 = Path.Combine(pastaDeDestino, nomeDoArquivoFinal);

            // Download da stream de áudio (MP4)
            await youtube.Videos.Streams.DownloadAsync(audioStreamInfo, caminhoDoMp3);

            Console.WriteLine("URL convertida para video");
        }
        else
        {
            Console.WriteLine("Não foi possível encontrar uma versão de áudio para o vídeo.");
        }
    }
}
