using System;
using System.Data;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using ReportSaleSystem.Models;
using System.Linq;
using System.Globalization;

namespace ReportSaleSystem
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                DirectoryInfo diretorio = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\data\in\");
                if (!diretorio.Exists)
                {
                    Console.WriteLine($"Diretorio gerado: {diretorio}");
                    diretorio.Create();
                }
                FileInfo[] arquivos = diretorio.GetFiles("*.dat", SearchOption.AllDirectories);
                List<Clientes> listaClientes = new List<Clientes>();
                List<Vendedores> listaVendedores = new List<Vendedores>();
                List<VendaItens> listaVendaItens = new List<VendaItens>();

                if (arquivos.Length > 0)
                {
                    foreach (FileInfo arquivo in arquivos)
                    {
                        string[] linhas = File.ReadAllText(arquivo.FullName).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);

                        if (linhas.Length > 0)
                        {
                            AdicionarVendedores(linhas.Where(x => x.Substring(0, 3) == "001"), listaVendedores);
                            AdicionarClientes(linhas.Where(x => x.Substring(0, 3) == "002"), listaClientes);
                            AdicionarVendaItens(linhas.Where(x => x.Substring(0, 3) == "003"), listaVendaItens);
                        }
                    }
                    CriarArquivo(listaClientes, listaVendedores, listaVendaItens);
                }
                else
                    Console.WriteLine("Nenhum arquivo com a extensão .dat foi encontrado");

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private static void CriarArquivo(List<Clientes> listaclientes, List<Vendedores> listaVendedores, List<VendaItens> listaVendaItens)
        {
            DirectoryInfo diretorio = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\data\out\");
            if (!diretorio.Exists)
            {
                Console.WriteLine($"Diretorio gerado: {diretorio}");
                diretorio.Create();
            }
            StreamWriter fileOutput = File.CreateText($"{ diretorio }relatorio_{ DateTime.Now.ToString("dd.MM.yyyy_HH.mm.ss") }.done.dat");
            fileOutput.WriteLine("Quantidade de clientes: {0} ", listaclientes.Count());
            fileOutput.WriteLine("Quantidade de vendedores: {0} ", listaVendedores.Count());
            fileOutput.WriteLine("IdVenda mais cara: {0}", listaVendaItens.GroupBy(c => c.SaleID).Select(s => new
            {
                IdVenda = s.Key,
                Venda = s.Sum(v => v.Total)
            }).OrderByDescending(a => a.Venda).ToList().FirstOrDefault().IdVenda);
            fileOutput.WriteLine("O pior vendedor: {0}", listaVendaItens.GroupBy(v => v.SalesmanName).Select(s => new
            {
                IdVenda = s.Key,
                Venda = s.Sum(v => v.Total)
            }).OrderBy(a => a.Venda).ToList().FirstOrDefault().IdVenda);
            Console.WriteLine($"Arquivo gerado: {((FileStream)fileOutput.BaseStream).Name}");
            fileOutput.Close();
        }

        private static void AdicionarVendaItens(IEnumerable<string> vendasRealizadas, List<VendaItens> listaVendaItens)
        {
            string[] vendasArrayString;
            string[] itemArrayString;

            foreach (var vendaRealizada in vendasRealizadas)
            {
                vendasArrayString = vendaRealizada.Replace("[", "").Replace("]", "").Split(new char[] { 'ç' });
                foreach (var itensvenda in vendasArrayString[2].Split(new char[] { ',' }))
                {
                    itemArrayString = itensvenda.Split(new char[] { '-' });
                    listaVendaItens.Add(new VendaItens
                    {
                        ItemID = double.Parse(itemArrayString[0], CultureInfo.InvariantCulture),
                        ItemQuantity = double.Parse(itemArrayString[1], CultureInfo.InvariantCulture),
                        ItemPrice = double.Parse(itemArrayString[2], CultureInfo.InvariantCulture),
                        SaleID = vendasArrayString[1],
                        SalesmanName = vendasArrayString[3]
                    });
                }
            }
        }

        private static void AdicionarClientes(IEnumerable<string> clientes, List<Clientes> listaClientes)
        {
            foreach (string cliente in clientes)
            {
                string[] split = cliente.Split(new char[] { 'ç' });

                listaClientes.Add(new Clientes
                {
                    CNPJ = split[1],
                    Name = split[2],
                    BusinessArea = split[3]
                });
            }
        }

        private static void AdicionarVendedores(IEnumerable<string> vendedores, List<Vendedores> listaVendedores)
        {
            foreach (string vendedor in vendedores)
            {
                string[] split = vendedor.Split(new char[] { 'ç' });

                listaVendedores.Add(new Vendedores
                {
                    CPF = split[1],
                    Name = split[2],
                    Salary = split[3]
                });
            }
        }
    }
}
