using OfficeOpenXml;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using OfficeOpenXml.Style;
using System.Drawing;


namespace UniPlan
{

    public partial class MainWindow : Window
    {
        private ObservableCollection<ClassRecord> _records = new ObservableCollection<ClassRecord>();
        private System.Windows.Data.CollectionViewSource _viewSource = new System.Windows.Data.CollectionViewSource();

        public MainWindow()
        {
            InitializeComponent();
            _viewSource.Source = _records;
            DataGridClasses.ItemsSource = _viewSource.View;

            UpdateRecordCount();
        }

        private void UpdateRecordCount()
        {
            if (_viewSource?.View == null)
            {
                TxtRecordCount.Text = $"تعداد رکوردها: {_records.Count}";
                return;
            }

            int count = _viewSource.View.Cast<object>().Count();
            TxtRecordCount.Text = $"تعداد رکوردها: {count}";
        }


        private void BtnSaveData_Click(object sender, RoutedEventArgs e)
        {
            var win = new AddEditRecordWindow(_records);
            win.Owner = this;

            if (win.ShowDialog() == true)
            {
                _records.Add(win.ResultRecord);
                UpdateRecordCount();
            }
        }
        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            string q = TxtSearch.Text?.Trim();

            if (string.IsNullOrWhiteSpace(q))
            {
                _viewSource.View.Filter = null;
            }
            else
            {
                _viewSource.View.Filter = obj =>
                {
                    if (obj is not ClassRecord r) return false;

                    return (r.CourseTitle ?? "").Contains(q) ||
                           (r.CourseCode ?? "").Contains(q) ||
                           (r.MainInstructor ?? "").Contains(q) ||
                           (r.Semester ?? "").Contains(q);
                };
            }

            UpdateRecordCount();
        }


        private void BtnEditRecord_Click(object sender, RoutedEventArgs e)
        {
            if (DataGridClasses.SelectedItem is not ClassRecord selected)
            {
                MessageBox.Show("ابتدا یک رکورد را انتخاب کنید.", "خطا", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var win = new AddEditRecordWindow(_records, selected);
            win.Owner = this;

            if (win.ShowDialog() == true)
            {
                selected.Semester = win.ResultRecord.Semester;
                selected.CourseTitle = win.ResultRecord.CourseTitle;
                selected.CourseCode = win.ResultRecord.CourseCode;
                selected.MainInstructor = win.ResultRecord.MainInstructor;
                selected.ClassTime = win.ResultRecord.ClassTime;
                selected.ExamDate = win.ResultRecord.ExamDate;
                selected.Capacity = win.ResultRecord.Capacity;

                DataGridClasses.Items.Refresh();
            }
        }

        private void BtnDeleteRecord_Click(object sender, RoutedEventArgs e)
        {
            if (DataGridClasses.SelectedItem is not ClassRecord selected)
            {
                MessageBox.Show("ابتدا یک رکورد را انتخاب کنید.", "خطا", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show("آیا مطمئن هستید این رکورد حذف شود؟",
                "تأیید حذف", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _records.Remove(selected);
                UpdateRecordCount();
            }
        }
        private void BtnDeleteAll_Click(object sender, RoutedEventArgs e)
        {
            var res = MessageBox.Show(
                "آیا مطمئن هستید همه رکوردها حذف شوند؟",
                "هشدار",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (res == MessageBoxResult.Yes)
            {
                _records.Clear();
                UpdateRecordCount();
            }
        }

        private void ExportToExcel(string path, System.Collections.Generic.List<ClassRecord> list)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using var package = new OfficeOpenXml.ExcelPackage();
                var ws = package.Workbook.Worksheets.Add("برنامه هفتگی");
                ws.View.RightToLeft = true;

                ws.PrinterSettings.Orientation = eOrientation.Landscape;
                ws.PrinterSettings.PaperSize = ePaperSize.A4;
                ws.PrinterSettings.FitToPage = true;
                ws.PrinterSettings.FitToWidth = 1;
                ws.PrinterSettings.FitToHeight = 1;

                string semesterName = list.FirstOrDefault()?.Semester ?? "نامشخص";

                ws.Row(1).Height = 35;

                var cellRight = ws.Cells["A1:C1"];
                cellRight.Merge = true;
                cellRight.Value = "نام و کد رشته: ";
                cellRight.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                var cellCenter = ws.Cells["F1:H1"];
                cellCenter.Merge = true;
                cellCenter.Value = "مقطع تحصیلی: ";
                cellCenter.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                var cellLeft = ws.Cells["K1:M1"];
                cellLeft.Merge = true;
                cellLeft.Value = $"برنامه ترم: {semesterName}";
                cellLeft.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                var headerRow = ws.Cells["A1:M1"];
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Font.Name = "Tahoma";
                headerRow.Style.Font.Size = 10f;
                headerRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                int startHour = 8;
                int endHour = 20;
                string[] days = { "شنبه", "یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه" };

                for (int h = startHour; h < endHour; h++)
                {
                    var cell = ws.Cells[2, h - startHour + 2];
                    cell.Value = $"{h:D2}:00 - {h + 1:D2}:00";
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(230, 230, 230));
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    ws.Column(h - startHour + 2).Width = 18;
                }

                for (int i = 0; i < days.Length; i++)
                {
                    var cell = ws.Cells[i + 3, 1];
                    cell.Value = days[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(70, 70, 70));
                    cell.Style.Font.Color.SetColor(Color.White);
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Row(i + 3).Height = 75;
                }
                ws.Column(1).Width = 12;

                string Normalize(string t) => t?.Replace("ي", "ی").Replace("ك", "ک").Replace("‌", "").Replace(" ", "") ?? "";

                foreach (var item in list)
                {
                    if (string.IsNullOrEmpty(item.ClassTime)) continue;
                    string normTime = Normalize(item.ClassTime);
                    int rowIndex = -1;

                    if (normTime.Contains(Normalize("یکشنبه"))) rowIndex = 4;
                    else if (normTime.Contains(Normalize("دوشنبه"))) rowIndex = 5;
                    else if (normTime.Contains(Normalize("سه‌شنبه"))) rowIndex = 6;
                    else if (normTime.Contains(Normalize("چهارشنبه"))) rowIndex = 7;
                    else if (normTime.Contains(Normalize("شنبه"))) rowIndex = 3;

                    if (rowIndex == -1) continue;

                    var matches = System.Text.RegularExpressions.Regex.Matches(item.ClassTime, @"(\d{1,2})");
                    if (matches.Count >= 2)
                    {
                        int sHour = int.Parse(matches[0].Value);
                        int eHour = matches.Count >= 3 ? int.Parse(matches[2].Value) : int.Parse(matches[1].Value);

                        int startCol = sHour - startHour + 2;
                        int endCol = eHour - startHour + 1;

                        if (startCol >= 2 && endCol >= startCol)
                        {
                            var firstCell = ws.Cells[rowIndex, startCol];
                            string info = $"{item.CourseTitle}\n{item.MainInstructor}";

                            if (firstCell.Value != null && !string.IsNullOrEmpty(firstCell.Value.ToString()))
                            {
                                if (!firstCell.Value.ToString().Contains(item.CourseTitle))
                                    firstCell.Value = firstCell.Value.ToString() + "\n----------\n" + info;
                            }
                            else { firstCell.Value = info; }

                            var range = ws.Cells[rowIndex, startCol, rowIndex, endCol];
                            if (range.Merge) range.Merge = false;
                            range.Merge = true;

                            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(firstCell.Value.ToString().Contains("----------") ? Color.FromArgb(245, 245, 245) : Color.White);
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            range.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
                            range.Style.Font.Size = 8.5f;
                            range.Style.Font.Name = "Tahoma";
                            range.Style.WrapText = true;
                        }
                    }
                }

                var tableArea = ws.Cells[2, 1, 7, 13];
                tableArea.Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Black);

                File.WriteAllBytes(path, package.GetAsByteArray());
                MessageBox.Show("خروجی شماتیکی با موفقیت ساخته شد ✅");
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در ایجاد اکسل: " + ex.Message);
            }
        }
        private void ExportSemesterToExcel(string semester)
        {
            var list = string.IsNullOrEmpty(semester)
                ? _records.ToList()
                : _records.Where(r => r.Semester == semester).ToList();

            if (list.Count == 0)
            {
                MessageBox.Show("رکوردی برای خروجی وجود ندارد");
                return;
            }
            SaveFileDialog dlg = new SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = $"UniPlan_{semester ?? "All"}.xlsx"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage())
                    {
                        var ws = package.Workbook.Worksheets.Add("لیست دروس");
                        ws.View.RightToLeft = true;

                        string[] headers = { "نیمسال", "درس", "کددرس", "مدرس‌اصلي", "ساعت‌کلاس", "تاریخ آزمون", "ظرفیت" };
                        for (int i = 0; i < headers.Length; i++)
                        {
                            ws.Cells[1, i + 1].Value = headers[i];
                            ws.Cells[1, i + 1].Style.Font.Bold = true;
                            ws.Cells[1, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ws.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        int row = 2;
                        foreach (var item in list)
                        {
                            ws.Cells[row, 1].Value = item.Semester;
                            ws.Cells[row, 2].Value = item.CourseTitle;
                            ws.Cells[row, 3].Value = item.CourseCode;
                            ws.Cells[row, 4].Value = item.MainInstructor;
                            ws.Cells[row, 5].Value = item.ClassTime;
                            ws.Cells[row, 6].Value = item.ExamDate;
                            ws.Cells[row, 7].Value = item.Capacity;
                            row++;
                        }

                        ws.Cells[ws.Dimension.Address].AutoFitColumns();

                        System.IO.File.WriteAllBytes(dlg.FileName, package.GetAsByteArray());
                        MessageBox.Show($"خروجی مربوط به {semester ?? "همه موارد"} با موفقیت ذخیره شد.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("خطا در هنگام ذخیره فایل: " + ex.Message);
                }
            }
        }

        private void BtnSaveSem1_Click(object sender, RoutedEventArgs e)
        {
            ExportSemesterToExcel("اول");
        }

        private void BtnSaveSem2_Click(object sender, RoutedEventArgs e)
        {
            ExportSemesterToExcel("دوم");
        }

        private void BtnSaveSem3_Click(object sender, RoutedEventArgs e)
        {
            ExportSemesterToExcel("سوم");
        }

        private void BtnSaveSem4_Click(object sender, RoutedEventArgs e)
        {
            ExportSemesterToExcel("چهارم");
        }

        private void BtnSaveSemAll_Click(object sender, RoutedEventArgs e)
        {
            ExportSemesterToExcel(null); // همه
        }

        private void ExportWeekly(string semester)
        {
            var list = _records.Where(r => r.Semester == semester).ToList();

            if (list.Count == 0)
            {
                MessageBox.Show("رکوردی برای این نیمسال وجود ندارد");
                return;
            }

            SaveFileDialog dlg = new SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = $"Weekly_{semester}.xlsx"
            };

            if (dlg.ShowDialog() == true)
            {
                ExportToExcel(dlg.FileName, list);
            }
        }

        private void BtnShematicSem1_Click(object sender, RoutedEventArgs e)
        {
            ExportWeekly("اول");
        }

        private void BtnShematicSem2_Click(object sender, RoutedEventArgs e)
        {
            ExportWeekly("دوم");
        }

        private void BtnShematicSem3_Click(object sender, RoutedEventArgs e)
        {
            ExportWeekly("سوم");
        }

        private void BtnShematicSem4_Click(object sender, RoutedEventArgs e)
        {
            ExportWeekly("چهارم");
        }

        private void BtnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog()
            {
                Filter = "Excel File (*.xlsx)|*.xlsx"
            };

            if (openDialog.ShowDialog() != true)
                return;

            try
            {
                using var package = new ExcelPackage(new FileInfo(openDialog.FileName));
                var ws = package.Workbook.Worksheets.FirstOrDefault();

                if (ws == null)
                {
                    MessageBox.Show("هیچ شیتی داخل فایل پیدا نشد.", "خطا", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                _records.Clear();



                int rowCount = ws.Dimension.End.Row;

                for (int row = 2; row <= rowCount; row++)
                {
                    var record = new ClassRecord()
                    {
                        Semester = ws.Cells[row, 1].Text,
                        CourseTitle = ws.Cells[row, 2].Text,
                        CourseCode = ws.Cells[row, 3].Text,
                        MainInstructor = ws.Cells[row, 4].Text,
                        ClassTime = ws.Cells[row, 5].Text,
                        ExamDate = ws.Cells[row, 6].Text,
                        Capacity = ws.Cells[row, 7].Text,
                    };

                    if (string.IsNullOrWhiteSpace(record.CourseTitle) &&
                        string.IsNullOrWhiteSpace(record.CourseCode))
                        continue;

                    _records.Add(record);
                }

                UpdateRecordCount();
                MessageBox.Show("فایل اکسل با موفقیت وارد شد ✅", "موفق", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در وارد کردن فایل اکسل:\n" + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            if (_viewSource?.View != null)
            {
                _viewSource.View.Filter = null;

                _viewSource.View.Refresh();

                UpdateRecordCount();
            }
        }
        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}
