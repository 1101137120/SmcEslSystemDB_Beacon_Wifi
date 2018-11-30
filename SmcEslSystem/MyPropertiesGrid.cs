using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SmcEslSystem
{


    public class MyPropertiesGridLabel
    {
        private string _Text;
        private Color _BackColor;
        private Point _Location;
        private int _Width;
        private int _Height;
        private bool _AutoSzie;
        private Font _Font;
        private Color _ForeColor;
        private Label _Control;
        

        
        public void SelectedObject(Label control)
        {
            this._Control = control;
        }



        [Category("字體調整"), DisplayName("內容")] 
        [Description("調整元件值")] 
        public  string Text
    {
        get { return this._Control.Text; }
        set
        {
            this._Control.Text = value;
        }
    }
        [Category("其他"), DisplayName("背景")]
        [Description("背景顏色調整(僅適用紅白黑三色)")]
        public Color BackColor
        {
            get { return this._Control.BackColor; }
            set
            {
                this._Control.BackColor = value;
            }
        }


        [Category("其他"), DisplayName("座標")]
        [Description("位於版面的X、Y座標")]
        public Point Location
        {
            get { return this._Control.Location; }
            set
            {
                this._Control.Location = value;
            }
        }

        [Category("文字框大小"), DisplayName("寬度")]
        [Description("文字框寬度調整")]
        public int Width
        {
            get { return this._Control.Width; }
            set
            {
                this._Control.Width = value;
            }
        }

        [Category("文字框大小"), DisplayName("高度")]
        [Description("文字框高度調整")]
        public int Height
        {
            get { return this._Control.Height; }
            set
            {
                this._Control.Height = value;
            }
        }

        [Category("文字框大小"), DisplayName("自動調整")]
        [Description("寬度高度依值自動調整")]
        public bool AutoSize
        {
            get { return this._Control.AutoSize; }
            set
            {
                this._Control.AutoSize = value;
            }
        }

        [Category("字體調整"), DisplayName("文字編輯")]
        [Description("文字樣式調整，大小、字型等等")]
        public Font Font
        {
            get {
                return this._Control.Font; }
            set
            {
                this._Control.Font = value;
            }
        }

        [Category("字體調整"), DisplayName("文字顏色")]
        [Description("調整文字顏色(僅適用紅白黑)")]
        public Color ForeColor
        {
            get { return this._Control.ForeColor; }
            set
            {
                this._Control.ForeColor = value;
            }
        }
        
        public MyPropertiesGridLabel() { }
    }


    public class MyPropertiesGridTextBox
    {
        private string _Text;
        private Color _BackColor;
        private Point _Location;
        private int _Width;
        private int _Height;
        private bool _AutoSzie;
        private Font _Font;
        private Color _ForeColor;
        private TextBox _Control;



        public void SelectedObject(TextBox control)
        {
            this._Control = control;
        }



        [Category("字體調整"), DisplayName("內容")]
        [Description("調整元件值")]
        public string Text
        {
            get { return this._Control.Text; }
            set
            {
                this._Control.Text = value;
            }
        }
        [Category("其他"), DisplayName("背景")]
        [Description("背景顏色調整(僅適用紅白黑三色)")]
        public Color BackColor
        {
            get { return this._Control.BackColor; }
            set
            {
                this._Control.BackColor = value;
            }
        }


        [Category("其他"), DisplayName("座標")]
        [Description("位於版面的X、Y座標")]
        public Point Location
        {
            get { return this._Control.Location; }
            set
            {
                this._Control.Location = value;
            }
        }

        [Category("文字框大小"), DisplayName("寬度")]
        [Description("文字框寬度調整")]
        public int Width
        {
            get { return this._Control.Width; }
            set
            {
                this._Control.Width = value;
            }
        }

        [Category("文字框大小"), DisplayName("高度")]
        [Description("文字框高度調整")]
        public int Height
        {
            get { return this._Control.Height; }
            set
            {
                this._Control.Height = value;
            }
        }

        [Category("文字框大小"), DisplayName("自動調整")]
        [Description("寬度高度依值自動調整")]
        public bool AutoSize
        {
            get { return this._Control.AutoSize; }
            set
            {
                this._Control.AutoSize = value;
            }
        }

        [Category("字體調整"), DisplayName("文字編輯")]
        [Description("文字樣式調整，大小、字型等等")]
        public Font Font
        {
            get
            {
                return this._Control.Font;
            }
            set
            {
                this._Control.Font = value;
            }
        }

        [Category("字體調整"), DisplayName("文字顏色")]
        [Description("調整文字顏色(僅適用紅白黑)")]
        public Color ForeColor
        {
            get { return this._Control.ForeColor; }
            set
            {
                this._Control.ForeColor = value;
            }
        }

        public MyPropertiesGridTextBox() { }
    }


    public class MyPropertiesGridPicBox
    {
        private string _Text;
        private Color _BackColor;
        private Point _Location;
        private int _Width;
        private int _Height;
        private bool _AutoSzie;
        private Font _Font;
        private Color _ForeColor;
        private PictureBox _Control;



        public void SelectedObject(PictureBox control)
        {
            this._Control = control;
        }



        [Category("字體調整"), DisplayName("內容")]
        [Description("調整元件值")]
        public string Text
        {
            get { return this._Control.Text; }
            set
            {
                this._Control.Text = value;
            }
        }
        [Category("其他"), DisplayName("背景")]
        [Description("背景顏色調整(僅適用紅白黑三色)")]
        public Color BackColor
        {
            get { return this._Control.BackColor; }
            set
            {
                this._Control.BackColor = value;
            }
        }


        [Category("其他"), DisplayName("座標")]
        [Description("位於版面的X、Y座標")]
        public Point Location
        {
            get { return this._Control.Location; }
            set
            {
                this._Control.Location = value;
            }
        }

        [Category("文字框大小"), DisplayName("寬度")]
        [Description("文字框寬度調整")]
        public int Width
        {
            get { return this._Control.Width; }
            set
            {
                this._Control.Width = value;
            }
        }

        [Category("文字框大小"), DisplayName("高度")]
        [Description("文字框高度調整")]
        public int Height
        {
            get { return this._Control.Height; }
            set
            {
                this._Control.Height = value;
            }
        }

        [Category("文字框大小"), DisplayName("自動調整")]
        [Description("寬度高度依值自動調整")]
        public bool AutoSize
        {
            get { return this._Control.AutoSize; }
            set
            {
                this._Control.AutoSize = value;
            }
        }

        [Category("字體調整"), DisplayName("文字編輯")]
        [Description("文字樣式調整，大小、字型等等")]
        public Font Font
        {
            get
            {
                return this._Control.Font;
            }
            set
            {
                this._Control.Font = value;
            }
        }

        [Category("字體調整"), DisplayName("文字顏色")]
        [Description("調整文字顏色(僅適用紅白黑)")]
        public Color ForeColor
        {
            get { return this._Control.ForeColor; }
            set
            {
                this._Control.ForeColor = value;
            }
        }

        public MyPropertiesGridPicBox() { }
    }

}
