using consultandoIp.Service;

static class MyProgram
{
    static void Main()
    {
        while (true)
        {

        string ip = new IpService().ConsultarIp().Ip;
        new ArquivoService().addLog(ip);
        
            System.Threading.Thread.Sleep(1000 * 2);

        }
    }
}