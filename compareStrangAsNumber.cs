using System;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;

namespace printSheets
{
    class CompareStrangAsNumber : IComparer
    {
        public CompareStrangAsNumber() { }

        public int Compare(object object1, object object2)
        {

            var Temp1 = Path.GetFileName((string)object1);
            var Temp2 = Path.GetFileName((string)object2);

            var temp1Number = Regex.Match(Temp1, @"\d+").Value;
            var temp2Number = Regex.Match(Temp2, @"\d+").Value;
            var number1 = Int32.Parse(temp1Number);
            var number2 = Int32.Parse(temp2Number);

            if (number1 > number2)
            {
                return 1;
            }
            if (number1 < number2)
            {
                return -1;
            }
            else
                return 0;
        }
    }
}
