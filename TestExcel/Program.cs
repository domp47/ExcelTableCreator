namespace TestExcel
{
    class Program
    {

        static void Main(string[] args)
        {
            var file = "C:\\temp\\Excel\\test.xlsx";

            new GeneratedClass().CreatePackage(file);
        }
    }
}