using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Eto.Drawing;
using Eto.Forms;

namespace FlatTable
{
    public class ExcelFilesControl
    {
        private TableLayout scrollLayout;
        private List<CheckBox> fileCheckBoxList;
        private List<string> filePathList;

        public ExcelFilesControl()
        {
            fileCheckBoxList = new List<CheckBox>();
            filePathList = new List<string>();
        }

        public TableLayout CreateExcelFileLayout()
        {
            scrollLayout = new TableLayout();

            Scrollable fileScrollable = CreateFileScrollable();
            TableLayout buttonsLayout = CreateSelectButtonsLayout();

            scrollLayout.Rows.Add(fileScrollable);
            scrollLayout.Rows.Add(buttonsLayout);
            scrollLayout.Rows.Add(new TableRow {ScaleHeight = true});
            return scrollLayout;
        }

        private Scrollable CreateFileScrollable()
        {
            Scrollable scrollable = new Scrollable {Size = new Size(-1, 350)};
            TableLayout fileLayout = new TableLayout {Padding = new Padding(10, 10, 20, 10), Spacing = new Size(0, 10)};
            if (string.IsNullOrEmpty(AppData.ExcelFolderPath) || !Directory.Exists(AppData.ExcelFolderPath))
            {
                fileLayout.Rows.Add(new Panel {Size = new Size(200, 100)});
                fileLayout.Rows.Add(new Label {Text = "无法读取到Excel文件，请检查Excel路径是否正确"});
                fileLayout.Rows.Add(new TableRow {ScaleHeight = true});
            }
            else
            {
                ConstructExcelFileLayout(ref fileLayout);
            }

            fileLayout.Rows.Add(new TableRow {ScaleHeight = true});

            scrollable.Content = fileLayout;
            return scrollable;
        }

        private void ConstructExcelFileLayout(ref TableLayout fileLayout)
        {
            fileCheckBoxList.Clear();
            filePathList.Clear();
            string[] filePaths = Directory.GetFiles(AppData.ExcelFolderPath);
            for (int i = 0; i < filePaths.Length; i++)
            {
//                Debug.WriteLine("file:  " + fileNames[i]);
                string filePath = filePaths[i];
                string extension = Path.GetExtension(filePath);

                if (extension != ".xlsx")
                {
                    continue;
                }

                string fileName = Path.GetFileName(filePath);

                //忽略Excel 2016的临时文件
                if (fileName[0] == '~')
                {
                    continue;
                }

                CheckBox newFileCheckBox = new CheckBox {Text = fileName};
                fileCheckBoxList.Add(newFileCheckBox);
                filePathList.Add(filePath);
                fileLayout.Rows.Add(newFileCheckBox);
            }

            if (fileCheckBoxList.Count == 0)
            {
                fileLayout.Rows.Add(new Label {Text = "文件夹内找不到任何有效的excel文件", Size = new Size(200, 200)});
            }

            fileLayout.Rows.Add(new TableRow {ScaleHeight = true});
        }

        private TableLayout CreateSelectButtonsLayout()
        {
            TableLayout layout = new TableLayout {Padding = new Padding(0, 10, 0, 10), Spacing = new Size(50, 0)};

            Button selectAllButton = new Button {Text = "选择所有文件"};
            selectAllButton.Click += OnSelectAllFileClick;
            Button selectNonButton = new Button {Text = "取消选择所有文件"};
            selectNonButton.Click += OnSelectNonFileClick;

            layout.Rows.Add(new TableRow(selectAllButton, selectNonButton));
            layout.Rows.Add(new TableRow {ScaleHeight = true});

            return layout;
        }

        private void OnSelectAllFileClick(object sender, EventArgs e)
        {
            for (int i = 0; i < fileCheckBoxList.Count; i++)
            {
                fileCheckBoxList[i].Checked = true;
            }
        }
        
        private void OnSelectNonFileClick(object sender, EventArgs e)
        {
            for (int i = 0; i < fileCheckBoxList.Count; i++)
            {
                fileCheckBoxList[i].Checked = false;
            }
        }

        public string[] GetCheckedFilePaths()
        {
            if (filePathList?.Count > 0)
            {
                List<string> resultPathList = new List<string>();
                for (int i = 0; i < fileCheckBoxList.Count; i++)
                {
                    CheckBox checkBox = fileCheckBoxList[i];
                    //checked box
                    if (checkBox.Checked.HasValue && checkBox.Checked.Value)
                    {
                        resultPathList.Add(filePathList[i]);
                    }
                }

                return resultPathList.ToArray();
            }

            return null;
        }
    }
}