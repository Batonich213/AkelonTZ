using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Excel_Table
{
    class Product
    {
        public string Name { get; set; }
        private uint ProductID { get; set; }
        private string Unit { get; set; }
        private float Price { get; set; }
    }
    class Customer
    {
        public int Id { get; set; }
        public string NameOfTheOrganization { get; set; }
        public string Address { get; set; }
        public string Name { get; set; }
        public Product OrderedProduct { get; set; }
        public List<Order> Orders { get; } = new List<Order>();
    }
    class Order
    {
        public int Id { get; set; }
        public int ProductID { get; set; }
        public int CustomerID { get; set; }
        public int orderId { get; set; }
        public int Quantity { get; set; }
        public DateTime Date { get; set; }
    }
    class Program
    {

        static void Main()
        {
            Console.WriteLine("Введите путь до файла с данными:");
            string filePath = Console.ReadLine();
            using (IXLWorkbook workbook = new XLWorkbook(filePath))
            {
                Console.WriteLine("\n1. Вывести по наименованию товара информацию о клиентах, заказавших этот товар \n2. Изменить контактное лицо \n3. Показать золотого клиента с наибольшим количеством заказов за указанный год или месяц ");
                string userSelection = Console.ReadLine();
                switch (Convert.ToInt16(userSelection))
                {
                    case 1:
                        Console.WriteLine("Введите название товара");
                        string productName = Console.ReadLine();
                        var worksheet = workbook.Worksheet(1);
                        var rows = worksheet.RowsUsed();
                        foreach (var row in rows)
                        {
                            if (row.Cell(2).Value.ToString() == productName)
                            {
                                Console.WriteLine(row);
                            }
                        }

                        break;


                }

            }
        }

        static void ChangeNameOfTheOrganization(string name, string NameOfTheOrganization)
        {

        }
    }

}
