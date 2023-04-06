using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProjectCode1
{
    internal class vehicle 
    {
        private string name;
        private int amount;

        public vehicle(string name, int amount)
        {
            this.name = name;
            this.amount = amount;
        }
        public string getName() { return name; }
        public int getAmount() { return amount; }

    }
}
