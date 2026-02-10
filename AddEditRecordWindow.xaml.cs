using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace UniPlan
{
    public partial class AddEditRecordWindow : Window
    {
        public ClassRecord ResultRecord { get; private set; }
        private readonly ObservableCollection<ClassRecord> _allRecords;
        private readonly bool _isEditMode;
        private readonly ClassRecord _editingRecord;

        public AddEditRecordWindow(ObservableCollection<ClassRecord> allRecords, ClassRecord recordToEdit = null)
        {
            InitializeComponent();
            _allRecords = allRecords;

            FillTimeCombos();

            if (recordToEdit != null)
            {
                _isEditMode = true;
                _editingRecord = recordToEdit;
                TxtTitle.Text = "ویرایش رکورد";
                BtnSave.Content = "ویرایش";
                LoadRecordToUI(recordToEdit);
            }
            else
            {
                _isEditMode = false;
                CmbSemester.SelectedIndex = 0;
                CmbClassDay.SelectedIndex = 0;
                TxtExamDateShamsi.Text = "1405/01/15";
            }
        }

        private void FillTimeCombos()
        {
            var classTimes = new List<string>();
            for (int h = 8; h <= 20; h++) classTimes.Add($"{h:D2}:00");
            CmbStartTime.ItemsSource = classTimes;
            CmbEndTime.ItemsSource = classTimes;

            var examTimes = new List<string>();
            for (int h = 8; h <= 16; h++)
            {
                examTimes.Add($"{h:D2}:00");
                if (h < 16) examTimes.Add($"{h:D2}:30");
            }
            CmbExamTime.ItemsSource = examTimes;

            CmbStartTime.SelectedItem = "08:00";
            CmbEndTime.SelectedItem = "10:00";
            CmbExamTime.SelectedItem = "08:00";
        }

        private void LoadRecordToUI(ClassRecord rec)
        {
            TxtCourseTitle.Text = rec.CourseTitle;
            TxtCourseCode.Text = rec.CourseCode;
            TxtInstructor.Text = rec.MainInstructor;
            TxtCapacity.Text = rec.Capacity;

            foreach (ComboBoxItem item in CmbSemester.Items)
            {
                if (item.Content.ToString() == rec.Semester)
                {
                    CmbSemester.SelectedItem = item;
                    break;
                }
            }

            if (!string.IsNullOrEmpty(rec.ExamDate) && rec.ExamDate.Contains(")"))
            {
                string[] examParts = rec.ExamDate.Split(')');
                CmbExamTime.SelectedItem = examParts[0].Replace("(", "").Trim();
                TxtExamDateShamsi.Text = examParts[1].Trim();
            }
            else
            {
                TxtExamDateShamsi.Text = rec.ExamDate;
            }

            if (!string.IsNullOrEmpty(rec.ClassTime))
            {
                string dayPart = rec.ClassTime.Split(' ')[0];
                foreach (ComboBoxItem d in CmbClassDay.Items)
                {
                    if (d.Content.ToString() == dayPart)
                    {
                        CmbClassDay.SelectedItem = d;
                        break;
                    }
                }

                try
                {
                    int indexFrom = rec.ClassTime.IndexOf("از ") + 3;
                    int indexTo = rec.ClassTime.IndexOf(" تا");

                    if (indexFrom > 2 && indexTo > indexFrom)
                    {
                        string startTime = rec.ClassTime.Substring(indexFrom, 5).Trim();
                        string endTime = rec.ClassTime.Substring(indexTo + 3).Trim();

                        CmbStartTime.SelectedItem = startTime;
                        CmbEndTime.SelectedItem = endTime;
                    }
                }
                catch
                {
                    CmbStartTime.SelectedIndex = 0;
                    CmbEndTime.SelectedIndex = 1;
                }
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            HideError();

            var semester = (CmbSemester.SelectedItem as ComboBoxItem)?.Content?.ToString();
            var day = (CmbClassDay.SelectedItem as ComboBoxItem)?.Content?.ToString();
            var startTimeStr = CmbStartTime.SelectedItem?.ToString();
            var endTimeStr = CmbEndTime.SelectedItem?.ToString();
            var examTimeStr = CmbExamTime.SelectedItem?.ToString();

            if (string.IsNullOrWhiteSpace(TxtInstructor.Text))
            { ShowError("❌ نام مدرس اصلی را وارد کنید."); return; }

            if (string.IsNullOrWhiteSpace(TxtCourseTitle.Text))
            { ShowError("❌ نام درس نمی‌تواند خالی باشد."); return; }

            if (!long.TryParse(TxtCourseCode.Text.Trim(), out long code) || code <= 0)
            {
                ShowError("❌ کد درس باید یک عدد مثبت و معتبر باشد (بدون حروف).");
                return;
            }

            if (!int.TryParse(TxtCapacity.Text.Trim(), out int cap) || cap <= 0)
            {
                ShowError("❌ ظرفیت باید یک عدد مثبت و معتبر باشد.");
                return;
            }

            if (!IsValidShamsiDate(TxtExamDateShamsi.Text.Trim()))
            { ShowError("❌ تاریخ آزمون نامعتبر است. نمونه: 1405/01/15"); return; }

            TimeSpan start = TimeSpan.Parse(startTimeStr);
            TimeSpan end = TimeSpan.Parse(endTimeStr);

            if (start >= end)
            { ShowError("❌ ساعت شروع کلاس نمی‌تواند بعد از ساعت پایان باشد!"); return; }

            if (HasTimeOverlap(semester, day, start, end))
            {
                ShowError($"⚠️ تداخل! در روز {day} برای نیمسال {semester} در این ساعت کلاس دیگری رزرو شده.");
                return;
            }
            bool isDuplicateCode = _allRecords.Any(r =>
                r.Semester == semester &&
                r.CourseCode == TxtCourseCode.Text.Trim() &&
                (!_isEditMode || r != _editingRecord));

            if (isDuplicateCode)
            {
                ShowError("❌ این کد درس قبلاً در این نیمسال ثبت شده است.");
                return;
            }

            string finalClassTime = $"{day} از {startTimeStr} تا {endTimeStr}";
            string finalExamDate = $"({examTimeStr}){TxtExamDateShamsi.Text.Trim()}";

            ResultRecord = new ClassRecord()
            {
                Semester = semester,
                CourseTitle = TxtCourseTitle.Text.Trim(),
                CourseCode = TxtCourseCode.Text.Trim(),
                MainInstructor = TxtInstructor.Text.Trim(),
                ClassTime = finalClassTime,
                ExamDate = finalExamDate,
                Capacity = TxtCapacity.Text.Trim()
            };

            DialogResult = true;
            Close();
        }

        private void OnlyNumericInPreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {

            e.Handled = !char.IsDigit(e.Text, e.Text.Length - 1);
        }
        private bool HasTimeOverlap(string semester, string day, TimeSpan start, TimeSpan end)
        {
            foreach (var item in _allRecords)
            {
                if (_isEditMode && item == _editingRecord) continue;
                if (item.Semester != semester || string.IsNullOrEmpty(item.ClassTime)) continue;

                if (!item.ClassTime.StartsWith(day)) continue;

                try
                {
                    var parts = item.ClassTime.Split(new[] { " از ", " تا " }, StringSplitOptions.None);
                    if (parts.Length >= 3)
                    {
                        TimeSpan s = TimeSpan.Parse(parts[1]);
                        TimeSpan e = TimeSpan.Parse(parts[2]);

                        if (start < e && end > s) return true;
                    }
                }
                catch { continue; }
            }
            return false;
        }

        private void ShowError(string message)
        {
            TxtError.Text = message;
            ErrorBorder.Visibility = Visibility.Visible;
        }

        private void HideError()
        {
            ErrorBorder.Visibility = Visibility.Collapsed;
        }
        private void BtnCancel_Click(object sender, RoutedEventArgs e) { DialogResult = false; Close(); }

        private bool IsValidShamsiDate(string input)
        {
            try
            {
                var p = input.Split('/');
                if (p.Length != 3) return false;
                new PersianCalendar().ToDateTime(int.Parse(p[0]), int.Parse(p[1]), int.Parse(p[2]), 0, 0, 0, 0);
                return true;
            }
            catch { return false; }
        }
    }
}