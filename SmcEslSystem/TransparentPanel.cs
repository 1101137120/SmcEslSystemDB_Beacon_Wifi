using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace SmcEslSystem
{
    public class TransparentPanel : Control
    {
        private Panel panel1;

        public TransparentPanel() { }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            //不进行背景的绘制  
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x00000020; //WS_EX_TRANSPARENT  
                return cp;
            }
        }

        protected override void OnPaint(System.Windows.Forms.PaintEventArgs e)
        {
            //绘制panel的背景图像  
            if (BackgroundImage != null) e.Graphics.DrawImage(this.BackgroundImage, new Point(0, 0));
        }

        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(200, 100);
            this.panel1.TabIndex = 0;
            this.ResumeLayout(false);

        }

        ////为控件添加自定义属性值num1  
        //private int num1 = 1;  

        //[Bindable(true), Category("自定义属性栏"), DefaultValue(1), Description("此处为自定义属性Attr1的说明信息！")]  
        //public int Attr1  
        //{  
        //    get { return num1; }  
        //    set { this.Invalidate(); }  
        //}  
    }

}
