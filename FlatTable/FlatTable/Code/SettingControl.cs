using System;
using System.Diagnostics;
using System.IO;
using Eto.Drawing;
using Eto.Forms;

namespace FlatTable
{
    public class SettingControl
    {
        private TableLayout settingAreaLayout;
        private Label excelFolderLabel;
        private Button excelFolderSelectButton;
        private Button jumpToExcelFolderButton;

        private Label csharpFolderLabel;
        private Button csharpFolderSelectButton;
        private Button jumpToCSharpFolderButton;

        private Label binaryFileFolderLabel;
        private Button binaryFileFolderSelectButton;
        private Button jumpToBinaryFileFolderButton;

        private Action onExcelPathChanged;

        public SettingControl(Action onExcelPathChanged)
        {
            settingAreaLayout = new TableLayout
            {
                Padding = new Padding(10, 30, 10, 10),
                Spacing = new Size(10, 20)
            };
            this.onExcelPathChanged = onExcelPathChanged;
        }

        public TableLayout CreateSettingAreaLayout()
        {
            TableRow excelFolderRow = CreateExcelFolderControlRow();
            TableRow csharpFolderRow = CreateCSharpFileFolderControlRow();
            TableRow binaryFolderRow = CreateBinaryFileFolderControlRow();

            settingAreaLayout.Rows.Add(excelFolderRow);
            settingAreaLayout.Rows.Add(csharpFolderRow);
            settingAreaLayout.Rows.Add(binaryFolderRow);
            settingAreaLayout.Rows.Add(new TableRow {ScaleHeight = true});

            return settingAreaLayout;
        }

        private TableRow CreateExcelFolderControlRow()
        {
            excelFolderLabel = new Label
                {Text = string.IsNullOrEmpty(AppData.ExcelFolderPath) ? "请选择目录" : AppData.ExcelFolderPath};

            excelFolderSelectButton = new Button {Text = "选择Excel文件目录"};
            excelFolderSelectButton.Click += OnExcelFolderSelectButtonClick;

            jumpToExcelFolderButton = new Button {Text = "跳转到文件目录"};
            jumpToExcelFolderButton.Click += OnJumpToExcelFolderButtonClick;

            return new TableRow(new TableCell(excelFolderLabel, true), excelFolderSelectButton, jumpToExcelFolderButton);
        }

        private TableRow CreateCSharpFileFolderControlRow()
        {
            csharpFolderLabel = new Label
                {Text = string.IsNullOrEmpty(AppData.CSharpFolderPath) ? "请选择目录" : AppData.CSharpFolderPath};

            csharpFolderSelectButton = new Button {Text = "选择C#输出文件目录"};
            csharpFolderSelectButton.Click += OnCSharpFolderSelectButtonClick;

            jumpToCSharpFolderButton = new Button {Text = "跳转到文件目录"};
            jumpToCSharpFolderButton.Click += OnJumpToCSharpFolderButtonClick;

            return new TableRow(new TableCell(csharpFolderLabel, true), csharpFolderSelectButton, jumpToCSharpFolderButton);
        }

        private TableRow CreateBinaryFileFolderControlRow()
        {
            binaryFileFolderLabel = new Label
                {Text = string.IsNullOrEmpty(AppData.BinaryFileFolderPath) ? "请选择目录" : AppData.BinaryFileFolderPath};

            binaryFileFolderSelectButton = new Button {Text = "选择二进制文件输出目录"};
            binaryFileFolderSelectButton.Click += OnBinaryFolderSelectButtonClick;

            jumpToBinaryFileFolderButton = new Button {Text = "跳转到文件目录"};
            jumpToBinaryFileFolderButton.Click += OnJumpToBinaryFolderButtonClick;

            return new TableRow(new TableCell(binaryFileFolderLabel, true), binaryFileFolderSelectButton,
                jumpToBinaryFileFolderButton);
        }

        private void OnJumpToExcelFolderButtonClick(object sender, EventArgs e)
        {
            if (Directory.Exists(AppData.ExcelFolderPath))
            {
                Application.Instance.Open(AppData.ExcelFolderPath);
            }
        }

        private void OnJumpToCSharpFolderButtonClick(object sender, EventArgs e)
        {
            if (Directory.Exists(AppData.CSharpFolderPath))
            {
                Application.Instance.Open(AppData.CSharpFolderPath);
            }
        }

        private void OnJumpToBinaryFolderButtonClick(object sender, EventArgs e)
        {
            if (Directory.Exists(AppData.BinaryFileFolderPath))
            {
                Application.Instance.Open(AppData.BinaryFileFolderPath);
            }
        }

        private void OnExcelFolderSelectButtonClick(object sender, EventArgs e)
        {
            SelectFolderDialog selectFolderDialog = new SelectFolderDialog {Directory = AppData.ExcelFolderPath};
            if (selectFolderDialog.ShowDialog(settingAreaLayout) == DialogResult.Ok)
            {
                AppData.ExcelFolderPath = selectFolderDialog.Directory;
                excelFolderLabel.Text = AppData.ExcelFolderPath;

                onExcelPathChanged?.Invoke();
            }
        }

        private void OnCSharpFolderSelectButtonClick(object sender, EventArgs e)
        {
            SelectFolderDialog selectFolderDialog = new SelectFolderDialog {Directory = AppData.CSharpFolderPath};
            if (selectFolderDialog.ShowDialog(settingAreaLayout) == DialogResult.Ok)
            {
                AppData.CSharpFolderPath = selectFolderDialog.Directory;
                csharpFolderLabel.Text = AppData.CSharpFolderPath;
            }
        }

        private void OnBinaryFolderSelectButtonClick(object sender, EventArgs e)
        {
            SelectFolderDialog selectFolderDialog = new SelectFolderDialog {Directory = AppData.BinaryFileFolderPath};
            if (selectFolderDialog.ShowDialog(settingAreaLayout) == DialogResult.Ok)
            {
                AppData.BinaryFileFolderPath = selectFolderDialog.Directory;
                binaryFileFolderLabel.Text = AppData.BinaryFileFolderPath;
            }
        }
    }
}