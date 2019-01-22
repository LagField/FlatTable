using System;
using Eto.Drawing;
using Eto.Forms;

namespace FlatTable
{
    public class FlatTableForm : Form
    {
        private SettingControl settingControl;
        
        public FlatTableForm()
        {
            Title = "FlatTable";
            ClientSize = new Size(600, 450);
            
            settingControl = new SettingControl();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            
            AppData.Init();

            TableLayout settingLayout = settingControl.CreateSettingAreaLayout();
            
            Content = new TableLayout
            {
                Rows =
                {
                    settingLayout,
                    new TableRow{ScaleHeight = true}
                }
            };
        }
    }
}