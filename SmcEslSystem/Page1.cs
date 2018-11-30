using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SmcEslSystem
{
    class Page1
    {
        public string no;
        public  string barcode;
        public  string product_name;
        public  string Brand;
        public  string specification;
        public  string price;
        public  string Special_offer;
        public  string Web;
        public string BleAddress;
        public string usingAddress;
        public string onsale;
        public string HeadertextALL;
        public string UpdateTime;
        public string UpdateState;
        public string APLink;
        public string onSaleTimeS;
        public string onSaleTimeE;
        public string tryUpdateState;
        public string ProductStyle;
        public string actionName;
        public string ESLSize;
        public System.Windows.Forms.Timer TimerConnect;
        public Stopwatch TimerSeconds;
    }
    class Page
    {
        public string No;
        public string BeaconProduct;
        public string ProductName;
        public string APID;
        public string ESLID;
        public DateTime SBeaconTime;
        public DateTime EBeaconTime;
        public string salesDay;
        public string Comment;

    }


    class pictureboxBarcode : PictureBox
    {
        public string barcodedata;

    }

    class PicPage {
        public string  Tag;
        public string  Name;
        public string Text;
        public int Width;
        public int Height;
        public int LocationX;
        public int LocationY;
        public string FontName;
        //    Console.WriteLine("Name" + x.Name + "width" + x.Width + x.Height + "textBox1.Location" + x.Location + "x.font" + x.Font + " x.ForeColor" + x.ForeColor.A + "," + x.ForeColor.R + "," + x.ForeColor.G + "," + x.ForeColor.B + "x.Font.Style" + x.Font.Style + "x.BackColor" + x.BackColor.A + "," + x.BackColor.R + "," + x.BackColor.G + "," + x.BackColor.B);
        public int FontSize;
        public int FontStyle;
        public int ForeColorA;
        public int ForeColorR;
        public int ForeColorG;
        public int ForeColorB;
        public int BackColorA;
        public int BackColorR;
        public int BackColorG;
        public int BackColorB;
    }

    class Item
    {
         public string Name;
         public int Value;
        public Item(string name, int value)
        {
            Name = name; Value = value;
        }

        public override string ToString()
        {
            // Generates the text shown in the combo box
            return Name;
        }
    }

    class BackPage
{
    public string NewMateESL;
    public string OldMateProduct;
    public bool isBack;


}

    class OldEslPage
    {
        public string ESLID;
        public int dataGridRowIndex;
    }

    class OldESLFormat
    {
        public string FormatName;
        public string Type;
    }


}
