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
        public Bike(string name, int price, string color) : base(name, price)
        {
            this.bike = color;
        }
        public string getSize() { return bike; }
    }
}
