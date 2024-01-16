using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Excel_Table
{
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
                Customer customer = new Customer();
               
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
