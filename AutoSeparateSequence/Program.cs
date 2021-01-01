using OfficeOpenXml;
using System;
using System.IO;
using AutoSeparateSequence.Services;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace AutoSeparateSequence
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourceFileName = $"图纸.xlsx";
            string saveFileName = $"输出-{sourceFileName}";

            if (!File.Exists(sourceFileName))
            {
                Console.WriteLine("图纸.xlsx文件缺失。");
                return;
            }

            string sheetName = "Sheet1";
            int maxNumber = 0; string maxNumberString = string.Empty;
            var sourceStringList1 = new List<string>();    //第1种原始字符串，eg:RC10B-2-04-01~03
            var sourceStringList2 = new List<string>();    //第2种原始字符串，主线高架桥 30m跨简支小箱梁预应力钢束图（一）~（三）
            string primaryString = string.Empty;    //第1种原始字符串对应的主字符串，RC10B-2-04-01~03 => RC10B-2-04
            string primaryString2 = string.Empty;    //第2种原始字符串对应的主字符串，主线高架桥 30m跨简支小箱梁预应力钢束图（一）~（三） => 主线高架桥 30m跨简支小箱梁预应力钢束图

            Regex regex;
            regex = new Regex(".*(?=（[一二三四五六七八九十]{1,3}）~（[一二三四五六七八九十]{1,3}）)");    //匹配结尾类似"（一）~（三）"之前的所有字符
            MatchCollection matchCollection;

            var file = new FileInfo(sourceFileName);

            int rowIndex = 2;

            //读入2种原始字符串
            using (var excelPackage = new ExcelPackage(file))
            {

                var worksheet = excelPackage.Workbook.Worksheets[sheetName];

                int rowCount = 2;// worksheet.Dimension.Rows;   //worksheet.Dimension.Rows指的是所有列中最大行
                //首行：表头不导入
                bool rowCur = true;    //行游标指示器
                                       //rowCur=false表示到达行尾
                                       //计算行数
                while (rowCur)
                {
                    try
                    {
                        //跳过表头
                        if (string.IsNullOrWhiteSpace(worksheet.Cells[rowCount + 1, 1].Value.ToString()))
                        {
                            rowCur = false;
                        }
                    }
                    catch (Exception)   //读取异常则终止
                    {
                        rowCur = false;
                    }

                    if (rowCur)
                    {
                        rowCount++;
                    }
                }

                int row = 2;    //excel中行指针
                for (row = 2; row <= rowCount; row++)
                {
                    try
                    {
                        sourceStringList1.Add(worksheet.Cells[row, ExcelService.FindColumnIndexByName(worksheet, "图号")].Value?.ToString() ?? string.Empty);
                        sourceStringList2.Add(worksheet.Cells[row, ExcelService.FindColumnIndexByName(worksheet, "名称")].Value?.ToString() ?? string.Empty);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine($"第{row}行数据读取出错。");
                        continue;
                    }

                }

            }

            var saveFile = new FileInfo(saveFileName);
            if (File.Exists(saveFileName))
            {
                File.Delete(saveFileName);
            }

            try
            {
                List<int> failRows = new List<int>();    //excel中写入失败的行数
                int rowCurr = 2;
                using (var excelPackage = new ExcelPackage(saveFile))
                {
                    // 添加worksheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    //第1种字符串
                    for (int i = 0; i < sourceStringList1.Count; i++)
                    {
                        //primaryString = sourceStringList1[i].Substring(0, sourceStringList1[i].Length - 6);
                        primaryString = sourceStringList1[i][0..^6];
                        maxNumberString = sourceStringList1[i].Substring(sourceStringList1[i].Length - 2);
                        if (maxNumberString.Substring(0) != "0")    //如果第1个字符不是0，则直接转换
                        {
                            maxNumber = Convert.ToInt32(maxNumberString);
                        }
                        else    //如果第1个字符是"0"，截取第2个字符
                        {
                            maxNumber = Convert.ToInt32(maxNumberString.Substring(1));
                        }

                        matchCollection = regex.Matches(sourceStringList2[i]);
                        //第2种字符串
                        for (int j = 0; j < maxNumber; j++)
                        {
                            if (j <= 8)    //如果是0-9,前面要+0
                                worksheet.Cells[rowCurr, 2].Value = $"{primaryString}-0{j + 1}";
                            else
                                worksheet.Cells[rowCurr, 2].Value = $"{primaryString}-{j + 1}";
                            primaryString2 = matchCollection[0].Value.ToString();    //仅取第1个匹配值
                            worksheet.Cells[rowCurr, 3].Value = $"{primaryString2}（{NumberHelper.NumberToChinese(j + 1)}）";

                            rowCurr++;
                        }

                    }

                    excelPackage.Save();
                }
                
            }
            catch (Exception ex)
            {
                Debug.Print($"保存excel出错，错误信息：{ex.Message}");
                return;
            }


            Console.WriteLine($"运行完成！已在当前目录生成{saveFileName}");
            Console.ReadKey();
        }


    }


}

