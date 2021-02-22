using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Newtonsoft.Json;
using System.Threading.Tasks;

namespace GmailQuickstart
{
    class user
    {
        public string userName { get; set; }
        public user(string userName)
        {
            this.userName = userName;
        }
        public int check()
        {
            using (StreamReader r = new StreamReader(@"D:\gmail-store.net\gmail store\users.json"))
            {
                string a = r.ReadToEnd();
                dynamic array = JsonConvert.DeserializeObject(a);
                foreach (var item in array)
                {
                    if (item.user.Value == this.userName)
                    {
                        if (item.orderType.Value != 1)
                        {
                            Console.WriteLine("order type of {0} is not by email", this.userName);
                            return 1;
                        }
                        return 0;
                    }
                }
                Console.WriteLine("{0} is not a registred user", this.userName);
                return 1;
            }
        }
    }
}
