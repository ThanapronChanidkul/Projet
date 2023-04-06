using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectCode1
{
    internal class Bike : vehicle
    {
        private string bike;
        public Bike(string name, int price, string amount) : base(name, price)
        {
            this.bike = amount;
        }
        public string getAmount() { return bike; }
    }
}
