using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SES.Utility
{
    public static partial class Ext
    {
        
        /// <summary>
        /// 获得指定小数位的小数(值只舍不入)
        /// </summary>
        /// <param name="number">数值</param>
        public static decimal ToShortDecimal(this decimal number,int decimalLength)
        {
            int num = 1;
            for (int i = 0; i < decimalLength; i++)
            {
                number = number * 10;
                num = num * 10;
            }

            return ((Int64)number) * 1.0m / num;

        }
    }
}
