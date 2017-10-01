using System;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.WindowsAzure.Storage.Table;

namespace ContactsApp.Models
{
    public class Contact : TableEntity
    {
        public string FullName { get; set; }
        public string PhoneNumber { get; set; }
        public string PostalCode { get; set; }
        public string Region { get; set; }
        public string City { get; set; }
        public string Address { get; set; }
        public string Email { get; set; }

        public Contact()
        {
        }

        public Contact(string phone, string name, string zip = "", string region = "", string city = "", string address = "", string email = "")
        {
            PartitionKey = phone;
            RowKey = string.Empty;
            PhoneNumber = phone;
            FullName = name;
            Email = IsValidEmail(email) ? email : string.Empty;
            SetAddress(zip, region, city, address);
        }

        private void SetAddress(string zip, string region, string city, string address)
        {
            if (zip.Any(x => !Char.IsDigit(x)) || region.Any(x => !Char.IsLetter(x) || !Char.IsWhiteSpace(x) || x.Equals('-')) || city.Any(x => !Char.IsLetter(x) || !Char.IsWhiteSpace(x) || x.Equals('-')) || address.Any(x => Char.IsSymbol(x)))
            {
                PostalCode = string.Empty;
                Region = string.Empty;
                City = string.Empty;
                Address = string.Empty;
            }
            else
            {
                PostalCode = zip;
                Region = region;
                City = city;
                Address = address;
            }
        }

        private bool IsValidEmail(string email)
        {
            if(string.IsNullOrEmpty(email))
            {
                return false;
            }

            return Regex.IsMatch(email, @"\A(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)\Z", RegexOptions.IgnoreCase); 
        }
    }
}
