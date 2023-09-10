using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using YoutubeExplode;
using YoutubeExplode.Converter;
using YoutubeExplode.Videos.Streams;
using CliWrap;
using CliWrap.EventStream;
using CliWrap.Buffered;
using System.Text;

class Program
{
    static async Task Main(string[] args)
    {
        var youtube = new YoutubeClient();

        // Get stream manifest
        var videoUrl = "https://www.youtube.com/watch?v=BF0uf7apZDQ";
        var streamManifest = await youtube.Videos.Streams.GetManifestAsync(videoUrl);

        // Select best audio stream (highest bitrate)
        var audioStreamInfo = streamManifest
            .GetAudioStreams()
            .Where(s => s.Container == Container.Mp4)
            .GetWithHighestBitrate();

        // Select best video stream (1080p60 in this example)
        var videoStreamInfo = streamManifest
            .GetVideoStreams()
            .Where(s => s.Container == Container.Mp4)
            .OrderByDescending(s => s.VideoQuality).FirstOrDefault();

        // Caminho para o diretório do projeto
        string diretorioDoProjeto = Directory.GetCurrentDirectory();

        // Caminho completo para o executável ffmpeg
        string caminhoDoFFmpeg = Path.Combine(diretorioDoProjeto, "ffmpeg.exe");

        // Pasta onde você deseja salvar o arquivo final
        string pastaDeDestino = diretorioDoProjeto;

        // Nome do arquivo final
        string nomeDoArquivoFinal = "video.mp4";

        // Caminhos completos para os arquivos de áudio e vídeo baixados
        string caminhoDoAudio = Path.Combine(pastaDeDestino, $"audio.{audioStreamInfo.Container}");
        string caminhoDoVideo = Path.Combine(pastaDeDestino, $"video.{videoStreamInfo.Container}");

        // Download dos streams de áudio e vídeo
        await youtube.Videos.Streams.DownloadAsync(audioStreamInfo, caminhoDoAudio);
        await youtube.Videos.Streams.DownloadAsync(videoStreamInfo, caminhoDoVideo);

        // Combinar áudio e vídeo usando o ffmpeg
        var ffmpegProcess = new Command(caminhoDoFFmpeg)
            .WithArguments($"-i \"{caminhoDoAudio}\" -i \"{caminhoDoVideo}\" -c:v copy -c:a aac -strict experimental \"{nomeDoArquivoFinal}\"")
            .WithWorkingDirectory(pastaDeDestino);

        var stdErr = new StringBuilder();

        await foreach (var cmdEvent in ffmpegProcess.ListenAsync())
        {
            if (cmdEvent is StandardErrorCommandEvent stdErrEvent)
            {
                stdErr.Append(stdErrEvent.Text);
            }
        }

        
    }
}
