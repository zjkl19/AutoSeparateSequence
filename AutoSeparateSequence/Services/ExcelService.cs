using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace AutoSeparateSequence.Services
{
    public static class ExcelService
    {
        /// <summary>
        /// 通过名称查找列索引
        /// </summary>
        /// <param name="workSheet">工作簿项目URL:https://epplussoftware.com/<
        /// <param name="keyWord">查找关键字
        /// <param name="maxSearchColumnCounts">最大查找列数，默认为100</param>
        /// <returns>找到则返回正确的列数（索引从1开始），否则返回0</returns>
        public static int FindColumnIndexByName(ExcelWorksheet workSheet, string keyWord, int maxSearchColumnCounts = 100)
        {
            for (int i = 1; i < maxSearchColumnCounts; i++)
            {
                if ((workSheet.Cells[1, i].Value?.ToString() ?? string.Empty) == keyWord)
                {
                    return i;
                }
            }

            return 0;
        }
    }
}
