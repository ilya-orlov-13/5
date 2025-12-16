using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
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
}
