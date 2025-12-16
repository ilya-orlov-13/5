using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
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
}
