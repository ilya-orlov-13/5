using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
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
}
