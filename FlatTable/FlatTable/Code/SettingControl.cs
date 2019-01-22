using System;
using System.Diagnostics;
using Eto.Drawing;
using Eto.Forms;

namespace FlatTable
{
    public class SettingControl
    {
        private TableLayout settingAreaLayout;
        private Label excelFolderLabel;
        private Button excelFolderSelectButton;
        

        public SettingControl()
        {
            settingAreaLayout = new TableLayout
            {
                Padding = new Padding(10,30,10,10),
            };
        }

        public TableLayout CreateSettingAreaLayout()
        {
            //file folder select
            excelFolderLabel = new Label();
            excelFolderLabel.Text = AppData.ExcelFolderPath;

            excelFolderSelectButton = new Button();
            excelFolderSelectButton.Text = "选择Excel文件目录";
            excelFolderSelectButton.Click += OnExcelFolderSelectButtonClick;
            

            settingAreaLayout.Rows.Add(new TableRow(new TableCell(excelFolderLabel, true), excelFolderSelectButton));
            settingAreaLayout.Rows.Add(new TableRow{ScaleHeight = true});

            return settingAreaLayout;
        }

        private void OnExcelFolderSelectButtonClick(object sender, EventArgs e)
        {
            SelectFolderDialog selectFolderDialog = new SelectFolderDialog();
            if (selectFolderDialog.ShowDialog(settingAreaLayout) == DialogResult.Ok)
            {
//                Debug.WriteLine("select directory: " + selectFolderDialog.Directory);
                AppData.ExcelFolderPath = selectFolderDialog.Directory;
                excelFolderLabel.Text = AppData.ExcelFolderPath;
            }
        }
    }
}