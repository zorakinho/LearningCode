using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using YoutubeExplode;
using YoutubeExplode.Videos.Streams;

class Program
{
    static async Task Main(string[] args)
    {

        // URL do vídeo do Youtube
        string videoUrl = "https://www.youtube.com/watch?v=BF0uf7apZDQ";

        var youtube = new YoutubeClient();


        var audioInfo = await youtube.Videos.GetAsync(videoUrl);

        var streamInfoSet = await youtube.Videos.Streams.GetManifestAsync(audioInfo.Id);

        var audioStreamInfo = streamInfoSet
            .GetAudioOnlyStreams()
            .GetWithHighestBitrate();


        // Insira o link do vídeo do YouTube que você deseja baixar
        

        // Obtém o manifest das streams
        var streamManifest = await youtube.Videos.Streams.GetManifestAsync(videoUrl);

        // Seleciona a melhor stream de áudio (a de maior taxa de bits)
        var videoStreamInfo = streamManifest
            .GetAudioStreams()
            .Where(s => s.Container == Container.Mp4) // Filtro para streams MP4
            .GetWithHighestBitrate();

        if (videoStreamInfo != null)
        {
            // Caminho para o diretório do projeto
            string diretorioDoProjeto = Directory.GetCurrentDirectory();

            // Pasta onde você deseja salvar o arquivo final (MP4 e MP3)
            string pastaDeDestino = diretorioDoProjeto;

            // Obtém informações do vídeo para obter o título
            var videoInfo = await youtube.Videos.GetAsync(videoUrl);

            // Título do vídeo (limpo para remover caracteres inválidos em nomes de arquivo)
            string tituloDoVideo = GetSafeFileName(videoInfo.Title);

            // Nome do arquivo final para o vídeo MP4
            string nomeDoArquivoFinalVideo = $"{tituloDoVideo}.mp4";

            // Nome do arquivo final para o áudio MP3
            string nomeDoArquivoFinalAudio = $"{tituloDoVideo}.mp3";

            // Caminho completo para o arquivo MP4
            string caminhoDoMp4 = Path.Combine(pastaDeDestino, nomeDoArquivoFinalVideo);

            // Caminho completo para o arquivo MP3
            string caminhoDoMp3 = Path.Combine(pastaDeDestino, nomeDoArquivoFinalAudio);

            // Download da stream de áudio (MP4)
            await youtube.Videos.Streams.DownloadAsync(videoStreamInfo, caminhoDoMp4);
            await youtube.Videos.Streams.DownloadAsync(audioStreamInfo, caminhoDoMp3);

            // Converta o arquivo MP4 em MP3 usando FFmpeg
            //ConvertMp4ToMp3(caminhoDoMp4, caminhoDoMp3);

            Console.WriteLine("URL convertida para Vídeo MP4 com sucesso!");
            Console.WriteLine("URL convertida para Áudio MP3 com sucesso!");
        }
        else
        {
            Console.WriteLine("Não foi possível encontrar uma versão de áudio para o vídeo.");
        }
    }

    // Função para obter um nome de arquivo seguro removendo caracteres inválidos
    static string GetSafeFileName(string fileName)
    {
        return string.Join("_", fileName.Split(Path.GetInvalidFileNameChars()));
    }

    // Função para converter um arquivo MP4 em MP3 usando FFmpeg
    // Função para converter um arquivo MP4 em MP3 usando FFmpeg

}


