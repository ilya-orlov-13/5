using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using ClosedXML.Excel;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            HotelManager manager = new HotelManager();
            manager.LoadData();

            while (true)
            {
                Console.Clear();
                Console.WriteLine("1. Просмотреть базу данных");
                Console.WriteLine("2. Добавить клиента");
                Console.WriteLine("3. Добавить номер");
                Console.WriteLine("4. Добавить бронирование");
                Console.WriteLine("5. Удалить клиента");
                Console.WriteLine("6. Удалить номер");
                Console.WriteLine("7. Удалить бронирование");
                Console.WriteLine("8. Запрос 1 (одна таблица)");
                Console.WriteLine("9. Запрос 2 (две таблицы)");
                Console.WriteLine("10. Запрос 3 (три таблицы)");
                Console.WriteLine("11. Запрос 4 (три таблицы)");
                Console.WriteLine("12. Сохранить изменения в Excel");
                Console.WriteLine("0. Выход");


                Console.WriteLine("\nВыберите действие: ");
                int choice = ConsoleValidator.ReadValidatedInt("", 0, 12);

                Console.Clear();

                switch (choice)
                {
                    case 1:
                        manager.DisplayData();
                        break;
                    case 2:
                        manager.AddClient();
                        break;
                    case 3:
                        manager.AddRoom();
                        break;
                    case 4:
                        manager.AddBooking();
                        break;
                    case 5:
                        manager.DeleteClient();
                        break;
                    case 6:
                        manager.DeleteRoom();
                        break;
                    case 7:
                        manager.DeleteBooking();
                        break;
                    case 8:
                        manager.Query1_SingleTable();
                        break;
                    case 9:
                        manager.Query2_TwoTables();
                        break;
                    case 10:
                        manager.Query3_ThreeTables_SingleValue();
                        break;
                    case 11:
                        manager.Query4_ThreeTables_SingleValue();
                        break;
                    case 12:
                        manager.SaveData();
                        break;
                    case 0:
                        Console.WriteLine("Выход из программы...");
                        return;
                }

                if (choice != '0')
                {
                    Console.WriteLine("\nНажмите любую клавишу для возврата в меню...");
                    Console.ReadKey();
                }
            }
        }
    }
}