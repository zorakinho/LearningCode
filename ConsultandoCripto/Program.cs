using ConsultandoCripto.Service;


static class MyProgram
{
    static void Main()
    {

        // MÉTODO LINQ com declarãção implicita 
        new CriptoService().GetCriptos().Data.ForEach(cripto => PrintCriptoInfo(cripto));



        //Método com linq com declaração  explicita

        /*
        var criptos = new CriptoService().GetCriptos().Data;

        criptos.ForEach(cripto =>
            Console.WriteLine($@"
            ID: {cripto.Id}
            Name: {cripto.Name}
            Minimo: {cripto.min_size:F6}
            "));
        
       */

        //CriptoListResponse criptos = new CriptoService().GetCriptos();


        /*
        foreach (var cripto in new CriptoService().GetCriptos().Data)
        {
            PrintCriptoInfo(cripto);
        }
        */


        // Método para tratar o retorno no loop
        static void PrintCriptoInfo(CriptoResponse cripto)
        {
            Console.WriteLine($@"
                ID: {cripto.Id}
                Name: {cripto.Name}
                Minimo: {Math.Round((decimal)cripto.min_size, 6)}
                ");
        }
    }
}