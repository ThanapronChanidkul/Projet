using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectCode1
{
    internal class Car : vehicle
    {
        private string car;
        public Car(string name, int price, string color) : base(name,price)
        {
            this.car = color;
        }
        public string getMeat() { return car; }
    }
}
