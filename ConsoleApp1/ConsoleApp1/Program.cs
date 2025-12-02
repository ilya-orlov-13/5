using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using ClosedXML.Excel;

public class Client
{
    public int ClientId { get; set; }
    public string LastName { get; set; }
    public string FirstName { get; set; }
    public string Patronymic { get; set; }
    public string Residence { get; set; }

    public Client(int clientId, string lastName, string firstName, string patronymic, string residence)
    {
        ClientId = clientId;
        LastName = lastName;
        FirstName = firstName;
        Patronymic = patronymic;
        Residence = residence;
    }

    public override string ToString()
    {
        return $"ID: {ClientId}, {LastName} {FirstName} {Patronymic}, {Residence}";
    }
}

public class Booking
{
    public int BookingId { get; set; }
    public int ClientId { get; set; }
    public int RoomId { get; set; }
    public DateTime BookingDate { get; set; }
    public DateTime CheckInDate { get; set; }
    public DateTime CheckOutDate { get; set; }

    public Booking(int bookingId, int clientId, int roomId, DateTime bookingDate, DateTime checkInDate, DateTime checkOutDate)
    {
        BookingId = bookingId;
        ClientId = clientId;
        RoomId = roomId;
        BookingDate = bookingDate;
        CheckInDate = checkInDate;
        CheckOutDate = checkOutDate;
    }

    public override string ToString()
    {
        return $"Бронь #{BookingId}, Клиент: {ClientId}, Номер: {RoomId}, {BookingDate:dd.MM.yyyy} - {CheckInDate:dd.MM.yyyy} до {CheckOutDate:dd.MM.yyyy}";
    }
}

public class Room
{
    public int RoomId { get; set; }
    public int Floor { get; set; }
    public int Capacity { get; set; }
    public decimal Price { get; set; }
    public int Category { get; set; }

    public Room(int roomId, int floor, int capacity, decimal price, int category)
    {
        RoomId = roomId;
        Floor = floor;
        Capacity = capacity;
        Price = price;
        Category = category;
    }

    public override string ToString()
    {
        return $"Номер {RoomId}, этаж {Floor}, мест: {Capacity}, цена: {Price:C}, категория: {Category}";
    }
}

public class HotelManager
{
    private List<Client> clients = new List<Client>();
    private List<Booking> bookings = new List<Booking>();
    private List<Room> rooms = new List<Room>();
    private string filePath = "LR5-var9.xlsx";

    public void LoadData()
    {
        try
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                LoadClients(workbook.Worksheet("Клиенты"));
                LoadBookings(workbook.Worksheet("Бронирование"));
                LoadRooms(workbook.Worksheet("Номера"));
            }
            Console.WriteLine("Данные успешно загружены из Excel файла.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при загрузке данных: {ex.Message}");
            Console.WriteLine("Созданы пустые списки для демонстрации функционала.");
        }
    }

    private void LoadClients(IXLWorksheet sheet)
    {
        var rows = sheet.RowsUsed().Skip(1);
        foreach (var row in rows)
        {
            try
            {
                int id = int.Parse(row.Cell(1).Value.ToString());
                string lastName = row.Cell(2).Value.ToString();
                string firstName = row.Cell(3).Value.ToString();
                string patronymic = row.Cell(4).Value.ToString();
                string residence = row.Cell(5).Value.ToString();

                clients.Add(new Client(id, lastName, firstName, patronymic, residence));
            }
            catch { }
        }
    }

    private void LoadBookings(IXLWorksheet sheet)
    {
        var rows = sheet.RowsUsed().Skip(1);
        foreach (var row in rows)
        {
            try
            {
                int bookingId = int.Parse(row.Cell(1).Value.ToString());
                int clientId = int.Parse(row.Cell(2).Value.ToString());
                int roomId = int.Parse(row.Cell(3).Value.ToString());
                DateTime bookingDate = DateTime.Parse(row.Cell(4).Value.ToString());
                DateTime checkInDate = DateTime.Parse(row.Cell(5).Value.ToString());
                DateTime checkOutDate = DateTime.Parse(row.Cell(6).Value.ToString());

                bookings.Add(new Booking(bookingId, clientId, roomId, bookingDate, checkInDate, checkOutDate));
            }
            catch { }
        }
    }

    private void LoadRooms(IXLWorksheet sheet)
    {
        var rows = sheet.RowsUsed().Skip(1);
        foreach (var row in rows)
        {
            try
            {
                int roomId = int.Parse(row.Cell(1).Value.ToString());
                int floor = int.Parse(row.Cell(2).Value.ToString());
                int capacity = int.Parse(row.Cell(3).Value.ToString());

                string priceStr = row.Cell(4).Value.ToString().Replace("р.", "").Replace(" ", "").Replace(",", ".");
                decimal price = decimal.Parse(priceStr, CultureInfo.InvariantCulture);

                int category = int.Parse(row.Cell(5).Value.ToString());

                rooms.Add(new Room(roomId, floor, capacity, price, category));
            }
            catch { }
        }
    }


    public void DisplayData()
    {
        Console.WriteLine("\nКлиенты:");
        foreach (var client in clients)
        {
            Console.WriteLine(client);
        }

        Console.WriteLine("\nНомера:");
        foreach (var room in rooms)
        {
            Console.WriteLine(room);
        }

        Console.WriteLine("\nБронирования:");
        foreach (var booking in bookings)
        {
            Console.WriteLine(booking);
        }
    }

    public void AddClient()
    {
        int id = ConsoleValidator.ReadValidatedInt("Введите ID клиента: ", 1, 10000);
        string lastName = ConsoleValidator.ReadValidatedString("Введите фамилию: ", 1, 50);
        string firstName = ConsoleValidator.ReadValidatedString("Введите имя: ", 1, 50);
        string patronymic = ConsoleValidator.ReadValidatedString("Введите отчество: ", 1, 50);
        string residence = ConsoleValidator.ReadValidatedString("Введите место жительства: ", 1, 100);

        if (clients.Any(c => c.ClientId == id))
        {
            Console.WriteLine("Ошибка: Клиент с таким ID уже существует.");
            return;
        }

        clients.Add(new Client(id, lastName, firstName, patronymic, residence));
        Console.WriteLine("Клиент успешно добавлен.");
    }

    public void AddRoom()
    {
        int id = ConsoleValidator.ReadValidatedInt("Введите номер комнаты: ", 1, 10000);
        int floor = ConsoleValidator.ReadValidatedInt("Введите этаж: ", 1, 20);
        int capacity = ConsoleValidator.ReadValidatedInt("Введите количество мест: ", 1, 10);
        decimal price = (decimal)ConsoleValidator.ReadValidatedDouble("Введите стоимость проживания за сутки: ", 0, 100000);
        int category = ConsoleValidator.ReadValidatedInt("Введите категорию гостиницы: ", 1, 5);

        if (rooms.Any(r => r.RoomId == id))
        {
            Console.WriteLine("Ошибка: Номер с таким ID уже существует.");
            return;
        }

        rooms.Add(new Room(id, floor, capacity, price, category));
        Console.WriteLine("Номер успешно добавлен.");
    }

    public void AddBooking()
    {
        int bookingId = ConsoleValidator.ReadValidatedInt("Введите ID бронирования: ", 1, 100000);
        int clientId = ConsoleValidator.ReadValidatedInt("Введите ID клиента: ", 1, 10000);
        int roomId = ConsoleValidator.ReadValidatedInt("Введите ID номера: ", 1, 10000);

        if (!clients.Any(c => c.ClientId == clientId))
        {
            Console.WriteLine("Ошибка: Клиент с таким ID не существует.");
            return;
        }

        if (!rooms.Any(r => r.RoomId == roomId))
        {
            Console.WriteLine("Ошибка: Номер с таким ID не существует.");
            return;
        }

        if (bookings.Any(b => b.BookingId == bookingId))
        {
            Console.WriteLine("Ошибка: Бронирование с таким ID уже существует.");
            return;
        }

        DateTime bookingDate = ReadDate("Введите дату бронирования (dd.MM.yyyy): ");
        DateTime checkInDate = ReadDate("Введите дату заезда (dd.MM.yyyy): ");
        DateTime checkOutDate = ReadDate("Введите дату выезда (dd.MM.yyyy): ");

        if (checkOutDate <= checkInDate)
        {
            Console.WriteLine("Ошибка: Дата выезда должна быть позже даты заезда.");
            return;
        }

        bookings.Add(new Booking(bookingId, clientId, roomId, bookingDate, checkInDate, checkOutDate));
        Console.WriteLine("Бронирование успешно добавлено.");
    }

    private DateTime ReadDate(string prompt)
    {
        while (true)
        {
            try
            {
                Console.Write(prompt);
                string input = Console.ReadLine();
                if (DateTime.TryParseExact(input, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                {
                    return date;
                }
                Console.WriteLine("Ошибка: Неверный формат даты. Используйте формат dd.MM.yyyy");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка: {ex.Message}");
            }
        }
    }

    public void DeleteClient()
    {
        if (clients.Count == 0)
        {
            Console.WriteLine("Список клиентов пуст.");
            return;
        }

        int id = ConsoleValidator.ReadValidatedInt("Введите ID клиента для удаления: ", 1, 10000);
        var client = clients.FirstOrDefault(c => c.ClientId == id);

        if (client == null)
        {
            Console.WriteLine("Клиент с таким ID не найден.");
            return;
        }

        var relatedBookings = bookings.Where(b => b.ClientId == id).ToList();
        if (relatedBookings.Any())
        {
            Console.WriteLine($"У клиента есть {relatedBookings.Count} бронирований. Удаление невозможно.");
            return;
        }

        clients.Remove(client);
        Console.WriteLine("Клиент успешно удален.");
    }

    public void DeleteRoom()
    {
        if (rooms.Count == 0)
        {
            Console.WriteLine("Список номеров пуст.");
            return;
        }

        int id = ConsoleValidator.ReadValidatedInt("Введите ID номера для удаления: ", 1, 10000);
        var room = rooms.FirstOrDefault(r => r.RoomId == id);

        if (room == null)
        {
            Console.WriteLine("Номер с таким ID не найден.");
            return;
        }

        var relatedBookings = bookings.Where(b => b.RoomId == id).ToList();
        if (relatedBookings.Any())
        {
            Console.WriteLine($"Номер задействован в {relatedBookings.Count} бронированиях. Удаление невозможно.");
            return;
        }

        rooms.Remove(room);
        Console.WriteLine("Номер успешно удален.");
    }

    public void DeleteBooking()
    {
        if (bookings.Count == 0)
        {
            Console.WriteLine("Список бронирований пуст.");
            return;
        }

        int id = ConsoleValidator.ReadValidatedInt("Введите ID бронирования для удаления: ", 1, 100000);
        var booking = bookings.FirstOrDefault(b => b.BookingId == id);

        if (booking == null)
        {
            Console.WriteLine("Бронирование с таким ID не найдено.");
            return;
        }

        bookings.Remove(booking);
        Console.WriteLine("Бронирование успешно удалено.");
    }

    public void Query1_SingleTable()
    {
        Console.WriteLine("\nЗапрос 1 (одна таблица): Найти все номера указанной категории с ценой больше заданной суммы");

        int category = ConsoleValidator.ReadValidatedInt("Введите категорию номера (1-5): ", 1, 5);
        decimal minPrice = (decimal)ConsoleValidator.ReadValidatedDouble("Введите минимальную стоимость (руб.): ", 0, 1000000);

        var result = rooms
            .Where(r => r.Category == category && r.Price > minPrice)
            .OrderBy(r => r.Price)
            .ToList();

        if (!result.Any())
        {
            Console.WriteLine("Нет результатов, удовлетворяющих условию.");
            return;
        }

        Console.WriteLine($"Найдено {result.Count} номеров:");
        foreach (var room in result)
        {
            Console.WriteLine(room);
        }
    }

    public void Query2_TwoTables()
    {
        Console.WriteLine("\nЗапрос 2 (две таблицы): Найти всех клиентов из указанного города и их бронирования");

        string city = ConsoleValidator.ReadValidatedString("Введите название города: ", 2, 50);

        var result = clients
            .Where(c => c.Residence.Contains(city))
            .Join(bookings,
                client => client.ClientId,
                booking => booking.ClientId,
                (client, booking) => new { client, booking })
            .OrderBy(x => x.client.LastName)
            .ToList();

        if (!result.Any())
        {
            Console.WriteLine("Нет результатов, удовлетворяющих условию.");
            return;
        }

        Console.WriteLine($"Найдено {result.Count} записей:");
        foreach (var item in result)
        {
            Console.WriteLine($"{item.client.LastName} {item.client.FirstName} - {item.booking}");
        }
    }

    public void Query3_ThreeTables_SingleValue()
    {
        Console.WriteLine("\nЗапрос 3 (три таблицы, одно значение): Количество бронирований номеров указанной категории для клиентов из заданного города за указанный период");

        // Получаем ввод данных от пользователя
        string city = ConsoleValidator.ReadValidatedString("Введите название города: ", 2, 50);
        int category = ConsoleValidator.ReadValidatedInt("Введите категорию номера (1-5): ", 1, 5);

        DateTime startDate = ReadDate("Введите начальную дату периода (dd.MM.yyyy): ");
        DateTime endDate = ReadDate("Введите конечную дату периода (dd.MM.yyyy): ");

        // Выполняем LINQ-запрос к трем таблицам
        var count = clients
            .Where(c => c.Residence.Contains(city))
            .Join(bookings,
                client => client.ClientId,
                booking => booking.ClientId,
                (client, booking) => new { client, booking })
            .Join(rooms,
                cb => cb.booking.RoomId,
                room => room.RoomId,
                (cb, room) => new { cb.client, cb.booking, room })
            .Count(x => x.room.Category == category &&
                       x.booking.CheckInDate >= startDate &&
                       x.booking.CheckInDate <= endDate);

        Console.WriteLine($"Количество бронирований: {count}");
        Console.WriteLine($"Ответ (целое число): {count}");
    }

    public void Query4_ThreeTables_SingleValue()
    {
        Console.WriteLine("\nЗапрос 4 (три таблицы, одно значение): Средняя стоимость проживания в номерах указанной категории для клиентов из заданного города за указанный период");

        // Получаем ввод данных от пользователя
        string city = ConsoleValidator.ReadValidatedString("Введите название города: ", 2, 50);
        int category = ConsoleValidator.ReadValidatedInt("Введите категорию номера (1-5): ", 1, 5);

        DateTime startDate = ReadDate("Введите начальную дату периода (dd.MM.yyyy): ");
        DateTime endDate = ReadDate("Введите конечную дату периода (dd.MM.yyyy): ");

        // Выполняем LINQ-запрос к трем таблицам
        var bookingsInPeriod = clients
            .Where(c => c.Residence.Contains(city))
            .Join(bookings,
                client => client.ClientId,
                booking => booking.ClientId,
                (client, booking) => new { client, booking })
            .Join(rooms,
                cb => cb.booking.RoomId,
                room => room.RoomId,
                (cb, room) => new { cb.client, cb.booking, room })
            .Where(x => x.room.Category == category &&
                       x.booking.CheckInDate >= startDate &&
                       x.booking.CheckInDate <= endDate)
            .ToList();

        if (!bookingsInPeriod.Any())
        {
            Console.WriteLine("Нет данных для расчета средней стоимости.");
            return;
        }

        // Рассчитываем среднюю стоимость
        double averageCost = bookingsInPeriod.Average(x => (double)x.room.Price);

        Console.WriteLine($"Средняя стоимость проживания: {averageCost:F2} руб.");
        Console.WriteLine($"Ответ: {(int)averageCost}");
    }



    public void SaveData()
    {
        try
        {
            using (var workbook = new XLWorkbook())
            {
                CreateClientsSheet(workbook);
                CreateBookingsSheet(workbook);
                CreateRoomsSheet(workbook);

                workbook.SaveAs(filePath);
                Console.WriteLine("Данные успешно сохранены в Excel файл.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при сохранении данных: {ex.Message}");
        }
    }

    private void CreateClientsSheet(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.Add("Клиенты");
        sheet.Cell(1, 1).Value = "Код клиента";
        sheet.Cell(1, 2).Value = "Фамилия";
        sheet.Cell(1, 3).Value = "Имя";
        sheet.Cell(1, 4).Value = "Отчество";
        sheet.Cell(1, 5).Value = "Место жительства";

        int row = 2;
        foreach (var client in clients.OrderBy(c => c.ClientId))
        {
            sheet.Cell(row, 1).Value = client.ClientId;
            sheet.Cell(row, 2).Value = client.LastName;
            sheet.Cell(row, 3).Value = client.FirstName;
            sheet.Cell(row, 4).Value = client.Patronymic;
            sheet.Cell(row, 5).Value = client.Residence;
            row++;
        }
    }

    private void CreateBookingsSheet(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.Add("Бронирование");
        sheet.Cell(1, 1).Value = "Код бронирования";
        sheet.Cell(1, 2).Value = "Код клиента";
        sheet.Cell(1, 3).Value = "Код номера";
        sheet.Cell(1, 4).Value = "Дата бронирования";
        sheet.Cell(1, 5).Value = "Дата заезда";
        sheet.Cell(1, 6).Value = "Дата выезда";

        int row = 2;
        foreach (var booking in bookings.OrderBy(b => b.BookingId))
        {
            sheet.Cell(row, 1).Value = booking.BookingId;
            sheet.Cell(row, 2).Value = booking.ClientId;
            sheet.Cell(row, 3).Value = booking.RoomId;
            sheet.Cell(row, 4).Value = booking.BookingDate.ToString("dd.MM.yyyy");
            sheet.Cell(row, 5).Value = booking.CheckInDate.ToString("dd.MM.yyyy");
            sheet.Cell(row, 6).Value = booking.CheckOutDate.ToString("dd.MM.yyyy");
            row++;
        }
    }

    private void CreateRoomsSheet(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.Add("Номера");
        sheet.Cell(1, 1).Value = "Код номера";
        sheet.Cell(1, 2).Value = "Этаж";
        sheet.Cell(1, 3).Value = "Число мест";
        sheet.Cell(1, 4).Value = "Стоимость проживания";
        sheet.Cell(1, 5).Value = "Категория";

        int row = 2;
        foreach (var room in rooms.OrderBy(r => r.RoomId))
        {
            sheet.Cell(row, 1).Value = room.RoomId;
            sheet.Cell(row, 2).Value = room.Floor;
            sheet.Cell(row, 3).Value = room.Capacity;
            sheet.Cell(row, 4).Value = $"{room.Price:F0} р.";
            sheet.Cell(row, 5).Value = room.Category;
            row++;
        }
    }
}

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
