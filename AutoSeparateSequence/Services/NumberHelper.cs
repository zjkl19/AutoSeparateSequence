using System;
using System.Collections.Generic;
using System.Text;

namespace AutoSeparateSequence.Services
{
    public static class NumberHelper
    {
        /// <summary>
        /// 数字转中文（仅支持1至99）
        /// </summary>
        /// <param name="number">eg: 22</param>
        /// <returns></returns>
        public static string NumberToChinese(int number)
        {
            switch (number)
            {
                case 10: return "十";
                case 11: return "十一";
                case 12: return "十二";
                case 13: return "十三";
                case 14: return "十四";
                case 15: return "十五";
                case 16: return "十六";
                case 17: return "十七";
                case 18: return "十八";
                case 19: return "十九";
            }

            string res = string.Empty;
            string str = number.ToString();
            string schar = str.Substring(0, 1);
            res = schar switch
            {
                "1" => "一",
                "2" => "二",
                "3" => "三",
                "4" => "四",
                "5" => "五",
                "6" => "六",
                "7" => "七",
                "8" => "八",
                "9" => "九",
                _ => "",
            };
            if (str.Length > 1)
            {
                switch (str.Length)
                {
                    case 2:
                    case 6:
                        res += "十";
                        break;
                    case 3:
                    case 7:
                        res += "百";
                        break;
                    case 4:
                        res += "千";
                        break;
                    case 5:
                        res += "万";
                        break;
                    default:
                        res += "";
                        break;
                }
                res += NumberToChinese(int.Parse(str.Substring(1, str.Length - 1)));
            }
            return res;
        }
    }
}
