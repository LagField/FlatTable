using System;
using Eto.Drawing;
using Eto.Forms;

namespace FlatTable
{
    public class FlatTableForm : Form
    {
        private SettingControl settingControl;
        private ExcelFilesControl excelFilesControl;
        private FileProcesser fileProcesser;
        private TableLayout settingLayout;
        private TableLayout excelFileLayout;
        private TableLayout processButtonLayout;

        public FlatTableForm()
        {
            Title = "FlatTable";
            ClientSize = new Size(800, 400);

            settingControl = new SettingControl(OnExcelPathChanged);
            excelFilesControl = new ExcelFilesControl();
            fileProcesser = new FileProcesser();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            AppData.Init();

            settingLayout = settingControl.CreateSettingAreaLayout();
            excelFileLayout = excelFilesControl.CreateExcelFileLayout();

            Button processButton = new Button {Text = "开始生成", Size = new Size(100, 50)};
            processButton.Click += OnProcessButtonClick;
            processButtonLayout = new TableLayout
            {
                Rows =
                {
                    new TableRow(new TableCell {ScaleWidth = true}, new TableCell(processButton),
                        new TableCell {ScaleWidth = true}),
                }
            };

            ConstructeLayout();
        }

        private void ConstructeLayout()
        {
            Content = new TableLayout
            {
                Spacing = new Size(20, 0),

                Rows =
                {
                    new TableRow(excelFileLayout, new TableLayout
                    {
                        Rows =
                        {
                            settingLayout,
                            processButtonLayout,
                            new TableRow {ScaleHeight = true}
                        },
                    }),
                    new TableRow {ScaleHeight = true}
                }
            };
        }

        /// <summary>
        /// 开始处理所有文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="eventArgs"></param>
        private void OnProcessButtonClick(object sender, EventArgs eventArgs)
        {
            string[] checkedFilePath = excelFilesControl.GetCheckedFilePaths();
            if (checkedFilePath?.Length > 0)
            {
                fileProcesser.Process(checkedFilePath);
            }
        }

        private void OnExcelPathChanged()
        {
            excelFileLayout = excelFilesControl.CreateExcelFileLayout();
            ConstructeLayout();
        }
    }
}