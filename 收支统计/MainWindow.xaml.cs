using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace 收支统计
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private Dispatcher dispatcher = Application.Current.Dispatcher;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Btn1_Click(object sender, RoutedEventArgs e)
        {
            rtbInfobox.Clear();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            string path = Path.Combine(Environment.CurrentDirectory, "基础数据");
            string[] files = Directory.GetFiles(path);

            #region 校验Excel格式，仅使用xlsx

            var xlsExcel = files.Where(a => Path.GetExtension(a) == ".xls");
            if (xlsExcel.Count() > 0)
            {
                MessageBox.Show("仅限使用.xlsx数据文件，而当前文件中存在.xls公式!\r\n请先转换格式再使用！", "提示"
                    , MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            #endregion

            new Thread((ThreadStart)delegate
            {
                var enumerable = from a in GetVoucherDetail(files)
                                 orderby a.发票管理区
                                 group a by a.发票管理区 into g
                                 select new
                                 {
                                     name = g.Key,
                                     totalPrice = g.Sum((Info a) => a.价税合计),
                                     collects = g.ToList()
                                 };
                string text = Path.Combine(Environment.CurrentDirectory, "结果数据");
                if (!Directory.Exists(text))
                {
                    Directory.CreateDirectory(text);
                }
                string name = default;
                foreach (var item in enumerable)
                {
                    name = item.name;
                    if (name == "")
                    {
                        name = "未知发票管理单位";
                    }
                    string text2 = Path.Combine(text, name + ".xlsx");
                    if (File.Exists(text2))
                    {
                        new FileInfo(text2).Delete();
                    }

                    ExcelPackage pack = new ExcelPackage(new FileInfo(text2));
                    try
                    {
                        ExcelWorksheet sheet1 = pack.Workbook.Worksheets.Add("收款统计表");
                        ExcelWorksheet sheet2 = pack.Workbook.Worksheets.Add("未收款");
                        sheet1.Cells[1, 1].Value = name + " 公司";
                        sheet1.Cells[1, 1].Style.Font.Size = 18;

                        sheet1.Cells[1, 1].Style.Font.Name = "宋体";
                        sheet1.Cells[1, 1].Style.Font.Bold = true;

                        sheet1.Cells[1, 1, 1, 15].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                        sheet1.Cells[2, 1].Value = "制表日期：";
                        sheet1.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        sheet1.Cells[2, 2].Value = (object)DateTime.Now.ToShortDateString();
                        sheet1.Cells[2, 13].Value = "管理区名称：";

                        sheet1.Cells[2, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet1.Cells[2, 14].Value = name;
                        sheet1.Cells[2, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        string[] array = "会计凭证号,出纳凭证号,收款单号,收款账户,收款日期,发票号码,实际管理区,客户名称,开票日期,商品名称/收费项目,项目所属月份,备注/摘要,金额,税额,价税合计".Split(',');
                        sheet1.Row(3).Height = 22;
                        sheet1.Row(2).Height = 22;
                        for (int i = 0; i < array.Length; i++)
                        {
                            SetStyleValue(sheet1.Cells[3, i + 1], array[i], IsHorizontalAligCenter: true, isBold: true);
                        }
                        sheet2.Cells[1, 1].Value = name + " 公司";
                        sheet2.Cells[1, 1].Style.Font.Size = 18f;
                        sheet2.Cells[1, 1].Style.Font.Name = "宋体";
                        sheet2.Cells[1, 1].Style.Font.Bold = true;//.get_Font().set_Bold(true);
                        sheet2.Cells[1, 1, 1, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                        sheet2.Cells[2, 1].Value = (object)"制表日期：";

                        sheet2.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet2.Cells[2, 2].Value = ((object)DateTime.Now.ToString("yyyy-MM-dd"));
                        sheet2.Cells[2, 11].Value = ((object)"管理区名称：");
                        sheet2.Cells[2, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;//
                        sheet2.Cells[2, 12].Value = ((object)name);
                        sheet2.Cells[2, 12].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        string[] array2 = "会计凭证号,出纳凭证号,收款账户,收款日期,发票号码,客户名称,开票日期,商品名称/收费项目,项目所属月份,备注/摘要,金额,税额,价税合计".Split(',');
                        for (int j = 0; j < array2.Length; j++)
                        {
                            SetStyleValue(sheet2.Cells[3, j + 1], array2[j], IsHorizontalAligCenter: true, isBold: true);
                        }
                        sheet2.Row(3).Height = 22;
                        sheet2.Row(2).Height = 22;
                        int num = 3;
                        int num2 = 3;
                        var orderedEnumerable = (from x in item.collects
                                                 group x by x.出纳凭证号 into x
                                                 select new
                                                 {
                                                     name = x.Key,
                                                     collects = from c in x.ToList()
                                                                orderby c.收款单号
                                                                select c,
                                                     total = x.Sum((Info y) => y.价税合计)
                                                 } into z
                                                 orderby Regex.Match(z.name, "\\D+").Value
                                                 select z).ThenBy(z =>
                                                 {
                                                     string value = Regex.Match(z.name, "\\d+").Value;
                                                     return (!(value == "")) ? Convert.ToInt32(value) : 0;
                                                 });
                        decimal num3 = orderedEnumerable.Where(a => a.name != "").Sum(a => a.total);
                        decimal num4 = orderedEnumerable.Where(a => a.name == "").Sum(a => a.total);
                        TraceHelper.GetInstance().Info($"{name} ,已收款总合计={num3},未收款总合计={num4}", "main");
                        foreach (var a2 in orderedEnumerable)
                        {
                            if (a2.name.Trim() != "")
                            {
                                foreach (Info collect in a2.collects)
                                {
                                    num++;
                                    sheet1.Row(num).Height = (22.0);
                                    SetStyleValue(sheet1.Cells[num, 1], collect.会计凭证号, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet1.Cells[num, 2], collect.出纳凭证号, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet1.Cells[num, 3], collect.收款单号, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet1.Cells[num, 4], collect.收款账户, IsHorizontalAligCenter: true);
                                    if (collect.收款日期 == DateTime.MinValue)
                                    {
                                        SetStyleValue(sheet1.Cells[num, 5], "", IsHorizontalAligCenter: true);
                                    }
                                    else
                                    {
                                        SetStyleValue(sheet1.Cells[num, 5], collect.收款日期.ToString("yyyy年MM月"), IsHorizontalAligCenter: true);
                                    }
                                    SetStyleValue(sheet1.Cells[num, 6], collect.发票号码, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet1.Cells[num, 7], collect.实际管理区, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet1.Cells[num, 8], collect.客户名称, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet1.Cells[num, 9], collect.开票日期, IsHorizontalAligCenter: true);
                                    sheet1.Cells[num, 9].Style.Numberformat.Format = "yyyy-MM-dd";

                                    SetStyleValue(sheet1.Cells[num, 10], collect.商品名称);
                                    SetStyleValue(sheet1.Cells[num, 11], collect.项目所属月份, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet1.Cells[num, 12], collect.摘要);
                                    if (collect.金额 > 0m)
                                    {
                                        sheet1.Cells[num, 13].Value = collect.金额;
                                    }
                                    sheet1.Cells[num, 13].Style.Border.BorderAround((ExcelBorderStyle)4);
                                    if (collect.税额 > 0m)
                                    {
                                        sheet1.Cells[num, 14].Value = collect.税额;
                                    }
                                    sheet1.Cells[num, 14].Style.Border.BorderAround((ExcelBorderStyle)4);
                                    SetStyleValue(sheet1.Cells[num, 15], collect.价税合计);
                                }
                                foreach (var item2 in from x in a2.collects
                                                      where x.出纳凭证号 == a2.name
                                                      group x by x.商品名称 into x
                                                      select new
                                                      {
                                                          name = x.Key,
                                                          total = x.Sum((Info y) => y.价税合计)
                                                      })
                                {
                                    num++;
                                    sheet1.Row(num).Height = (22.0);
                                    for (int k = 1; k <= 15; k++)
                                    {
                                        SetStyleValue(sheet1.Cells[num, k], "");
                                    }
                                    SetStyleValue(sheet1.Cells[$"J{num}"], item2.name + "  合计", Color.IndianRed, IsHorizontalAligCenter: true, isBold: true);
                                    SetStyleValue(sheet1.Cells[$"O{num}"], item2.total, Color.IndianRed, IsHorizontalAligCenter: false, isBold: true);
                                }
                            }
                            else
                            {
                                foreach (Info collect2 in a2.collects)
                                {
                                    num2++;
                                    sheet2.Row(num2).Height = (22.0);
                                    SetStyleValue(sheet2.Cells[num2, 1], collect2.会计凭证号, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet2.Cells[num2, 2], "", IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet2.Cells[num2, 3], collect2.收款账户, IsHorizontalAligCenter: true);
                                    if (DateTime.MinValue == collect2.收款日期)
                                    {
                                        SetStyleValue(sheet2.Cells[num2, 4], "", IsHorizontalAligCenter: true);
                                    }
                                    else
                                    {
                                        SetStyleValue(sheet2.Cells[num2, 4], collect2.收款日期.ToString("yyyy年MM月"), IsHorizontalAligCenter: true);
                                    }
                                    SetStyleValue(sheet2.Cells[num2, 5], collect2.发票号码, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet2.Cells[num2, 6], collect2.客户名称, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet2.Cells[num2, 7], collect2.开票日期, IsHorizontalAligCenter: true);
                                    sheet2.Cells[num2, 7].Style.Numberformat.Format = "yyyy-MM-dd";
                                    SetStyleValue(sheet2.Cells[num2, 8], collect2.商品名称);
                                    SetStyleValue(sheet2.Cells[num2, 9], collect2.项目所属月份, IsHorizontalAligCenter: true);
                                    SetStyleValue(sheet2.Cells[num2, 10], collect2.摘要);
                                    if (collect2.金额 > 0m)
                                    {
                                        sheet2.Cells[num2, 11].Value = ((object)collect2.金额);
                                    }
                                    sheet2.Cells[num2, 11].Style.Border.BorderAround((ExcelBorderStyle)4);
                                    if (collect2.税额 > 0m)
                                    {
                                        sheet2.Cells[num2, 12].Value = ((object)collect2.税额);
                                    }
                                    sheet2.Cells[$"L{num2}"].Style.Border.BorderAround((ExcelBorderStyle)4);
                                    SetStyleValue(sheet2.Cells[num2, 13], collect2.价税合计);
                                }
                                num2++;
                                sheet2.Row(num2).Height = (22.0);
                                for (int l = 1; l <= 13; l++)
                                {
                                    SetStyleValue(sheet2.Cells[num2, l], "");
                                }
                                SetStyleValue(sheet2.Cells[$"J{num2}"], " 合计", Color.IndianRed, IsHorizontalAligCenter: true, isBold: true);
                                SetStyleValue(sheet2.Cells[$"O{num2}"], "", Color.IndianRed, IsHorizontalAligCenter: false, isBold: true);
                                sheet2.Cells[$"K{num2}"].Formula = $"=SUM(k4:k{num2 - 1})";
                                sheet2.Cells[$"L{num2}"].Formula = $"=SUM(L4:L{num2 - 1})";
                                sheet2.Cells[$"m{num2}"].Formula = $"=SUM(M4:M{num2 - 1})";

                                sheet2.Cells[$"K{num2}"].Style.Font.Color.SetColor(Color.IndianRed);
                                sheet2.Cells[$"L{num2}"].Style.Font.Color.SetColor(Color.IndianRed);
                                sheet2.Cells[$"m{num2}"].Style.Font.Color.SetColor(Color.IndianRed);
                                sheet2.Cells[3, 1, num, 13].Style.Border.BorderAround((ExcelBorderStyle)11, Color.Black);
                            }
                        }
                        num++;
                        sheet1.Row(num).Height = 22;
                        for (int m = 1; m <= 15; m++)
                        {
                            SetStyleValue(sheet1.Cells[num, m], "");
                        }
                        SetStyleValue(sheet1.Cells[$"L{num}"], " 合计", Color.IndianRed, IsHorizontalAligCenter: true, isBold: true);
                        if (num > 4)
                        {
                            sheet1.Cells[$"M{num}"].Formula = $"=SUM(M4:M{num - 1})/2";
                            sheet1.Cells[$"N{num}"].Formula = $"=SUM(N4:N{num - 1})/2";
                            sheet1.Cells[$"O{num}"].Formula = $"=SUM(O4:O{num - 1})/2";
                        }
                        sheet1.Cells[$"M{num}"].Style.Font.Color.SetColor(Color.IndianRed);
                        sheet1.Cells[$"N{num}"].Style.Font.Color.SetColor(Color.IndianRed);
                        sheet1.Cells[$"O{num}"].Style.Font.Color.SetColor(Color.IndianRed);
                        sheet1.Cells[3, 1, num, 15].Style.Border.BorderAround((ExcelBorderStyle)11, Color.Black);

                        sheet1.Cells.AutoFitColumns(0.0);
                        sheet1.Column(1).Width = 18;
                        sheet1.View.ZoomScale = 85;
                        sheet2.Cells.AutoFitColumns(0.0);
                        sheet2.Column(1).Width = 18;
                        sheet2.View.ZoomScale = 85;//
                        pack.Save();
                    }
                    finally
                    {
                        ((IDisposable)pack)?.Dispose();
                    }
                    dispatcher.Invoke(delegate
                    {
                        ShowInfo("生成报表文件：" + name + ".xlsx");
                    });
                }
                dispatcher.Invoke(delegate
                {
                    ShowInfo("-----全部完成！-----");
                });
            }).Start();
        }

        private List<Info> GetVoucherDetail(string[] files)
        {
            //IL_0047: Unknown result type (might be due to invalid IL or missing references)
            //IL_004e: Expected O, but got Unknown
            List<Info> list = new List<Info>();
            foreach (string file in files)
            {
                dispatcher.Invoke(delegate
                {
                    ShowLable(Path.GetFileNameWithoutExtension(file));
                });
                ExcelPackage val = new ExcelPackage(new FileInfo(file));
                try
                {
                    ExcelWorksheet sheet = val.Workbook.Worksheets[0];
                    if (file.Contains("发票信息"))
                    {
                        dispatcher.Invoke(delegate
                        {
                            ShowInfo("读取发票信息(" + Path.GetFileNameWithoutExtension(file) + ")");
                        });
                        list.AddRange(GetInvoiceInfo(sheet));
                    }
                    else if (file.Contains("经营收入"))
                    {
                        dispatcher.Invoke(delegate
                        {
                            ShowInfo("读取经营性收入:" + Path.GetFileNameWithoutExtension(file));
                        });
                        list.AddRange(GetOperatingInfo(sheet));
                    }
                    else if (file.Contains("收款收据"))
                    {
                        dispatcher.Invoke(delegate
                        {
                            ShowInfo("读取收款收据信息:" + Path.GetFileNameWithoutExtension(file));
                        });
                        list.AddRange(GetReceiptInfo(sheet));
                    }
                }
                finally
                {
                    ((IDisposable)val)?.Dispose();
                }
            }
            return list;
        }

        private List<Info> GetReceiptInfo(ExcelWorksheet sheet)
        {
            List<Info> list = new List<Info>();
            int rowend = sheet.Dimension.End.Row;
            int r;
            for (r = 6; r <= rowend; r++)
            {
                if (sheet.Cells[r, 1].Value == null)
                {
                    continue;
                }
                try
                {
                    Info info = new Info();
                    info.会计凭证号 = Convert.ToString(sheet.Cells[$"K{r}"].Value);
                    info.出纳凭证号 = Convert.ToString(sheet.Cells[$"J{r}"].Value);
                    info.收款单号 = "";
                    info.收款账户 = Convert.ToString(sheet.Cells[$"L{r}"].Value);
                    info.收款日期 = GetDateTimeFromOADate(sheet.Cells[$"I{r}"].Value);
                    info.客户名称 = Convert.ToString(sheet.Cells[$"C{r}"].Value);
                    info.开票日期 = GetDateTimeFromOADate(sheet.Cells[$"A{r}"].Value);
                    info.商品名称 = Convert.ToString(sheet.Cells[$"E{r}"].Value);
                    info.项目所属月份 = Convert.ToString(sheet.Cells[$"F{r}"].Value);
                    info.摘要 = Convert.ToString(sheet.Cells[$"D{r}"].Value);
                    info.价税合计 = Convert.ToDecimal(sheet.Cells[$"H{r}"].Value);
                    info.实际管理区 = "";
                    info.发票管理区 = Convert.ToString(sheet.Cells[$"B{r}"].Value).Trim('\r', '\n');
                    list.Add(info);
                    dispatcher.Invoke(delegate
                    {
                        ShowProgress(6L, rowend, r);
                    });
                }
                catch (Exception ex)
                {
                    TraceHelper.GetInstance().Warning($"表格第{r}行数据可能存在错误！详情{ex.Message}", "收款收据信息");
                    dispatcher.Invoke(delegate
                {
                    ShowInfo($"表格第{r}行数据可能存在错误！");
                });
                }
            }
            return list;
        }

        private List<Info> GetOperatingInfo(ExcelWorksheet sheet)
        {
            List<Info> list = new List<Info>();
            int r;
            for (r = 6; r <= sheet.Dimension.End.Row; r++)
            {
                try
                {
                    object value = sheet.Cells[r, 3].Value;
                    if (value != null && value.ToString() != "")
                    {
                        int num = 0;
                        for (int i = 11; i <= 56; i++)
                        {
                            object value2 = sheet.Cells[r, i].Value;
                            if (value2 != null && value2.ToString() != "" && Convert.ToDecimal(value2.ToString()) > 0m)
                            {
                                num++;
                                Info info = new Info();
                                info.会计凭证号 = Convert.ToString(sheet.Cells[$"BG{r}"].Value);
                                info.出纳凭证号 = Convert.ToString(sheet.Cells[$"BF{r}"].Value);
                                info.收款单号 = Convert.ToString(sheet.Cells[$"BI{r}"].Value);
                                info.收款账户 = Convert.ToString(sheet.Cells[$"BH{r}"].Value);
                                info.收款日期 = GetDateTimeFromOADate(sheet.Cells[$"BE{r}"].Value);
                                info.客户名称 = Convert.ToString(sheet.Cells[$"G{r}"].Value);
                                info.开票日期 = GetDateTimeFromOADate(sheet.Cells[$"B{r}"].Value);
                                info.商品名称 = Convert.ToString(sheet.Cells[5, i].Value);
                                info.摘要 = Convert.ToString(sheet.Cells[$"F{r}"].Value);
                                info.价税合计 = Convert.ToDecimal(sheet.Cells[r, i].Value);
                                info.实际管理区 = "";
                                info.发票管理区 = Convert.ToString(sheet.Cells[$"H{r}"].Value).Trim('\r', '\n');
                                list.Add(info);
                            }
                        }
                        if (num == 0)
                        {
                            Info info2 = new Info();
                            info2.会计凭证号 = Convert.ToString(sheet.Cells[$"BG{r}"].Value);
                            info2.出纳凭证号 = Convert.ToString(sheet.Cells[$"BF{r}"].Value);
                            info2.收款单号 = Convert.ToString(sheet.Cells[$"BI{r}"].Value);
                            info2.收款账户 = Convert.ToString(sheet.Cells[$"BH{r}"].Value);
                            info2.收款日期 = GetDateTimeFromOADate(sheet.Cells[$"BE{r}"].Value);
                            info2.客户名称 = Convert.ToString(sheet.Cells[$"G{r}"].Value);
                            info2.开票日期 = GetDateTimeFromOADate(sheet.Cells[$"B{r}"].Value);
                            info2.商品名称 = "其它";
                            info2.摘要 = Convert.ToString(sheet.Cells[$"F{r}"].Value);
                            info2.价税合计 = Convert.ToDecimal(sheet.Cells[$"J{r}"].Value);
                            info2.实际管理区 = "";
                            info2.发票管理区 = Convert.ToString(sheet.Cells[$"H{r}"].Value);
                            list.Add(info2);
                        }
                    }
                }
                catch (Exception ex)
                {
                    dispatcher.Invoke(delegate
                    {
                        ShowInfo($"表格第{r}行数据可能存在错误！");
                    });
                    TraceHelper.GetInstance().Warning($"表格第{r}行数据可能存在错误！详情{ex.Message}", "经营收入");
                }
                dispatcher.Invoke(delegate
                {
                    ShowProgress(6L, sheet.Dimension.End.Row, r);
                });
            }
            return list;
        }

        private List<Info> GetInvoiceInfo(ExcelWorksheet sheet)
        {
            string value = Regex.Match(sheet.Cells["A1"].Value.ToString(), "管理区：(.+?)（").Groups[1].Value;
            List<Info> list = new List<Info>();
            int r;
            for (r = 5; r <= sheet.Dimension.End.Row; r++)
            {
                try
                {
                    object value2 = sheet.Cells[r, 10].Value;
                    if (value2 != null && !value2.ToString().Contains("计"))
                    {
                        Info info = new Info();
                        info.会计凭证号 = Convert.ToString(sheet.Cells[$"W{r}"].Value);
                        info.出纳凭证号 = Convert.ToString(sheet.Cells[$"V{r}"].Value);
                        info.收款单号 = Convert.ToString(sheet.Cells[$"Z{r}"].Value);
                        info.收款账户 = Convert.ToString(sheet.Cells[$"x{r}"].Value);
                        info.收款日期 = GetDateTimeFromOADate(sheet.Cells[$"U{r}"].Value);
                        info.发票号码 = Convert.ToString(sheet.Cells[$"B{r}"].Value);
                        info.实际管理区 = Convert.ToString(sheet.Cells[$"Y{r}"].Value).Trim('\r', '\n');
                        info.客户名称 = Convert.ToString(sheet.Cells[$"D{r}"].Value);
                        info.开票日期 = GetDateTimeFromOADate(sheet.Cells[$"H{r}"].Value);
                        info.商品名称 = Convert.ToString(sheet.Cells[$"J{r}"].Value);
                        info.项目所属月份 = "";
                        info.摘要 = Convert.ToString(sheet.Cells[$"T{r}"].Value);
                        info.金额 = Convert.ToDecimal(sheet.Cells[$"P{r}"].Value);
                        info.税额 = Convert.ToDecimal(sheet.Cells[$"Q{r}"].Value);
                        info.价税合计 = Convert.ToDecimal(sheet.Cells[$"R{r}"].Value);
                        info.发票管理区 = value;
                        list.Add(info);
                    }
                }
                catch (Exception ex)
                {
                    dispatcher.Invoke(delegate
                    {
                        ShowInfo($"表格第{r}行数据可能存在错误！");
                    });
                    TraceHelper.GetInstance().Warning($"表格第{r}行数据可能存在错误！详情{ex.Message}", "发票信息");
                }
                dispatcher.Invoke(delegate
                {
                    ShowProgress(5L, sheet.Dimension.End.Row, r);
                });
            }
            return list;
        }

        public DateTime GetDateTimeFromOADate(object value)
        {
            try
            {
                return Convert.ToDateTime(value);
            }
            catch (Exception)
            {
                return DateTime.FromOADate(Convert.ToDouble(value));
            }
        }

        public void SetStyleValue(ExcelRange rng, object value, bool IsHorizontalAligCenter = false, bool isBold = false)
        {
            rng.Value = value;
            rng.Style.Border.BorderAround((ExcelBorderStyle)4);
            rng.Style.VerticalAlignment = ((ExcelVerticalAlignment)1);
            if (IsHorizontalAligCenter)
            {
                rng.Style.HorizontalAlignment = (ExcelHorizontalAlignment)2;
            }
            rng.Style.Font.Bold = isBold;
        }

        public void SetStyleValue(ExcelRange rng, object value, Color color, bool IsHorizontalAligCenter = false, bool isBold = false)
        {
            rng.Value = (value);
            rng.Style.Border.BorderAround((ExcelBorderStyle)4);
            rng.Style.VerticalAlignment = ((ExcelVerticalAlignment)1);
            if (IsHorizontalAligCenter)
            {
                rng.Style.HorizontalAlignment = ((ExcelHorizontalAlignment)2);
            }
            rng.Style.Font.Bold = isBold;
            rng.Style.Font.Color.SetColor(color);
        }

        public void ShowInfo(string msg)
        {
            TextBox textBox = rtbInfobox;
            textBox.Text = textBox.Text + DateTime.Now.ToString() + "\t" + msg + "\r";
        }

        public void ShowProgress(long min, long max, long curNum)
        {
            prb1.Value = curNum;
            prb1.Maximum = max;
            prb1.Minimum = min;
            lbl2.Content = $"正在读取表格 {curNum}行，共{max}行";
        }

        public void ShowLable(string file)
        {
            lbl1.Content = file;
        }
    }
}