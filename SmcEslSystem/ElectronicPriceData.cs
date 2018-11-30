using System.Drawing;
using System.Windows.Forms;
using ZXing;                  // for BarcodeWriter
using SmcEslLib;
using System;
using System.Collections.Generic;
using System.Drawing.Text;

namespace SmcEslSystem
{
    class ElectronicPriceData
    {

        SmcDataToImage mSmcDataToImage = new SmcDataToImage();
        int bw = 212;
        int bh = 104;

        /*  public Bitmap setPage1(string fontName, string title, string brand, string specification,
              string price, string special_offer, string barcode, string qrcode)
          {
            // if (barcode.Equals("")) {
             //     MessageBox.Show("BARCODE不為空值");
             // }
                  Bitmap bmp = new Bitmap(212, 104);
              using (Graphics graphics = Graphics.FromImage(bmp))
              {
                  graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, bw, bh);
              }
              Console.WriteLine(fontName+ title+ brand+ specification+ price+ special_offer+ barcode+ qrcode);
              Bitmap bar = new Bitmap(120, 20);
              Bitmap bqr = new Bitmap(35, 35);
              BarcodeWriter barcode_w = new BarcodeWriter();       // 建立條碼物件
              barcode_w.Format = BarcodeFormat.EAN_13;            // 條碼類別.
              barcode_w.Options.Width = 120;
              barcode_w.Options.Height = 20;
              barcode_w.Options.PureBarcode = true;               // 顯示條碼字串
              if (!barcode.Equals(""))
              {
                  bar = barcode_w.Write(barcode);
                  //Console.WriteLine(bar.Width+"   "+bar.Height);
              }

              BarcodeWriter qr = new BarcodeWriter();
              qr.Format = BarcodeFormat.QR_CODE;
              qr.Options.Width = 35;
              qr.Options.Height = 35;
              qr.Options.Margin = 0;
              if (!qrcode.Equals(""))
              {
                  bqr = qr.Write(qrcode);
              }

              //標頭
              TextBox t1 = new TextBox();
              t1.Text = title;
              t1.Font = new Font(fontName, 12, FontStyle.Bold);
              t1.TextAlign = HorizontalAlignment.Center; //置中
              t1.BorderStyle = BorderStyle.FixedSingle;
              t1.Width = 206;
              t1.Height = 18;
              bmp = mSmcDataToImage.ConvertTextToImage(bmp, t1, Color.Black, 3, 8);

              if (special_offer.Equals("0"))
              {
                  Label L1 = new Label();
                  L1.Text = "品牌 : " + brand;
                  L1.Font = new Font(fontName, 9, FontStyle.Regular);
                  if (!brand.Equals(""))
                  {
                      bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, Color.Black, 3, 35);
                  }
                  L1.Text = "規格 : " + specification;
                  L1.Font = new Font(fontName, 9, FontStyle.Regular);
                  if (!specification.Equals("")) {
                      bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, Color.Black, 3, 55);
                  }

                  if (!qrcode.Equals(""))
                  {
                      bmp = mSmcDataToImage.ConvertImageToImage(bmp, bqr, 170, 55); //  QRcode
                  }
                  if (!barcode.Equals(""))
                  {
                      bmp = mSmcDataToImage.ConvertImageToImage(bmp, bar, 0, 80); //Barcode
                  }

                  L1.Text = price;
                  L1.Font = new Font(fontName, 36, FontStyle.Bold);
                  if (price.Length > 2)
                  {
                      if (!price.Equals(""))
                      {
                          L1.Font = new Font(fontName, 30, FontStyle.Bold);
                          bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, Color.Red, 95, 40); //價格
                      }
                  }else
                  {
                      if (!price.Equals(""))
                      {
                          bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, Color.Red, 105, 30); //價格
                      }
                  }

              }else
              {
                  Label L1 = new Label();
                  L1.Text = "品牌 : " + brand;
                  L1.Font = new Font(fontName, 9, FontStyle.Regular);
                  if (!brand.Equals(""))
                  {
                      bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, Color.Black, 3, 30);
                  }

                  L1.Text = "規格 : " + specification;
                  L1.Font = new Font(fontName, 9, FontStyle.Regular);
                  if (!specification.Equals(""))
                  {
                      bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, Color.Black, 3, 45);
                  }

                  L1.Text = "售價 : " + price;
                  L1.Font = new Font(fontName, 9, FontStyle.Regular);
                  if (!price.Equals(""))
                  {
                      bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, Color.Black, 3, 60);
                  }


                  L1.Text = special_offer;
                  L1.Font = new Font(fontName, 38, FontStyle.Bold);

                  if (special_offer.Length > 2)
                  {
                      if (!special_offer.Equals(""))
                      {
                          L1.Font = new Font(fontName, 34, FontStyle.Bold);
                          bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, Color.Red, 118, 40); //促銷價
                      }
                  }
                  else
                  {
                      if (!special_offer.Equals(""))
                      {
                          bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, Color.Red, 128, 40); //促銷價
                      }
                  }


                  L1.Text = " 促銷價 ";
                  L1.Font = new Font(fontName, 11, FontStyle.Regular);
                  L1.BackColor = Color.Black;
                  if (!special_offer.Equals(""))
                  {
                      bmp = mSmcDataToImage.ConvertBoxToImage(bmp, L1, Color.White, 135, 30);
                  }


                  if (!qrcode.Equals("")) 
                  {
                      bmp = mSmcDataToImage.ConvertImageToImage(bmp, bqr, 90, 42); //QRcode
                  }
                  if (!barcode.Equals(""))
                  {
                      bmp = mSmcDataToImage.ConvertImageToImage(bmp, bar, 0, 80); //Barcode
                  }

              }
              return bmp;
          }*/


        public Bitmap setPage1(string fontName, string title, string brand, string specification,
             string price, string special_offer, string barcode, string qrcode,string use_address, string headerAll, List<string> ESLFormat)
        {
            // if (barcode.Equals("")) {
            ////     MessageBox.Show("BARCODE不為空值");
            // }



            Bitmap bmp;
            if (ESLFormat[1] == "0")
                bmp = new Bitmap(212, 104);
            else if(ESLFormat[1] == "1")
                bmp = new Bitmap(296, 128);
            else if (ESLFormat[1] == "2")
                bmp = new Bitmap(400, 300);
            else
                bmp = new Bitmap(212, 104);


            using (Graphics graphics = Graphics.FromImage(bmp))
            {
                graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, bw, bh);
            }



            for (int i = 0; i < ESLFormat.Count; i++)
            {
                // Console.WriteLine("ESLFormat[i]" + ESLFormat[i]);

                if (ESLFormat[i] == "品名(最多10字)" || ESLFormat[i] == "品牌" || ESLFormat[i] == "規格" || ESLFormat[i] == "售價" || ESLFormat[i] == "促銷價"  || ESLFormat[i] == "ReaderID" || ESLFormat[i] == "貨架" || ESLFormat[i] == "特價" || ESLFormat[i] == "下架" || ESLFormat[i] == "UpTime" || ESLFormat[i] == "UpTimeE" || ESLFormat[i] == "use_address")
                {
                    if (ESLFormat[i] == "品牌")
                    {

                        Color backcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 13]), Convert.ToInt32(ESLFormat[i + 14]), Convert.ToInt32(ESLFormat[i + 15]), Convert.ToInt32(ESLFormat[i + 16]));
                        Color fontcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 9]), Convert.ToInt32(ESLFormat[i + 10]), Convert.ToInt32(ESLFormat[i + 11]), Convert.ToInt32(ESLFormat[i + 12]));
                        mytrydata(ESLFormat[i - 1], ESLFormat[i], brand, Convert.ToInt32(ESLFormat[i + 2]), Convert.ToInt32(ESLFormat[i + 3]), Convert.ToInt32(ESLFormat[i + 4]), Convert.ToInt32(ESLFormat[i + 5]), ESLFormat[i + 6], float.Parse(ESLFormat[i + 7]),ESLFormat[i + 8], fontcolor, backcolor, bmp);
                    }
                    else if (ESLFormat[i] == "品名(最多10字)")
                    {
                        //標頭
                        Color backcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 13]), Convert.ToInt32(ESLFormat[i + 14]), Convert.ToInt32(ESLFormat[i + 15]), Convert.ToInt32(ESLFormat[i + 16]));
                        Color fontcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 9]), Convert.ToInt32(ESLFormat[i + 10]), Convert.ToInt32(ESLFormat[i + 11]), Convert.ToInt32(ESLFormat[i + 12]));
                        mytrydata(ESLFormat[i - 1], ESLFormat[i], title, Convert.ToInt32(ESLFormat[i + 2]), Convert.ToInt32(ESLFormat[i + 3]), Convert.ToInt32(ESLFormat[i + 4]), Convert.ToInt32(ESLFormat[i + 5]), ESLFormat[i + 6], float.Parse(ESLFormat[i + 7]), ESLFormat[i + 8], fontcolor, backcolor, bmp);
                    }
                    else if (ESLFormat[i] == "規格")
                    {
                        Color backcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 13]), Convert.ToInt32(ESLFormat[i + 14]), Convert.ToInt32(ESLFormat[i + 15]), Convert.ToInt32(ESLFormat[i + 16]));
                        Color fontcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 9]), Convert.ToInt32(ESLFormat[i + 10]), Convert.ToInt32(ESLFormat[i + 11]), Convert.ToInt32(ESLFormat[i + 12]));
                        mytrydata(ESLFormat[i - 1], ESLFormat[i], specification, Convert.ToInt32(ESLFormat[i + 2]), Convert.ToInt32(ESLFormat[i + 3]), Convert.ToInt32(ESLFormat[i + 4]), Convert.ToInt32(ESLFormat[i + 5]), ESLFormat[i + 6], float.Parse(ESLFormat[i + 7]), ESLFormat[i + 8], fontcolor, backcolor, bmp);
                    }
                    else if (ESLFormat[i] == "use_address"&& ESLFormat[i-1] != "B" && ESLFormat[i - 1] != "use_address" && ESLFormat[i + 1] != "use_address")
                    {
                     //   Console.WriteLine("use_address" + use_address);
                        Color backcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 13]), Convert.ToInt32(ESLFormat[i + 14]), Convert.ToInt32(ESLFormat[i + 15]), Convert.ToInt32(ESLFormat[i + 16]));
                        Color fontcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 9]), Convert.ToInt32(ESLFormat[i + 10]), Convert.ToInt32(ESLFormat[i + 11]), Convert.ToInt32(ESLFormat[i + 12]));
                        mytrydata(ESLFormat[i - 1], ESLFormat[i], use_address, Convert.ToInt32(ESLFormat[i + 2]), Convert.ToInt32(ESLFormat[i + 3]), Convert.ToInt32(ESLFormat[i + 4]), Convert.ToInt32(ESLFormat[i + 5]), ESLFormat[i + 6], float.Parse(ESLFormat[i + 7]),ESLFormat[i + 8], fontcolor, backcolor, bmp);
                    }
                    else if (ESLFormat[i] == "售價")
                    {
                        if (ESLFormat[i - 1] != "Header" && ESLFormat[i - 1] != "售價")
                        {
                            Color backcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 13]), Convert.ToInt32(ESLFormat[i + 14]), Convert.ToInt32(ESLFormat[i + 15]), Convert.ToInt32(ESLFormat[i + 16]));
                            Color fontcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 9]), Convert.ToInt32(ESLFormat[i + 10]), Convert.ToInt32(ESLFormat[i + 11]), Convert.ToInt32(ESLFormat[i + 12]));
                            mytrydata(ESLFormat[i - 1], ESLFormat[i], price, Convert.ToInt32(ESLFormat[i + 2]), Convert.ToInt32(ESLFormat[i + 3]), Convert.ToInt32(ESLFormat[i + 4]), Convert.ToInt32(ESLFormat[i + 5]), ESLFormat[i + 6], float.Parse(ESLFormat[i + 7]),ESLFormat[i + 8], fontcolor, backcolor, bmp);
                        }
                    }
                    else if (ESLFormat[i] == "促銷價")
                    {
                        if (ESLFormat[i - 1] != "Header"&& ESLFormat[i - 1] != "促銷價") {
                            Color backcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 13]), Convert.ToInt32(ESLFormat[i + 14]), Convert.ToInt32(ESLFormat[i + 15]), Convert.ToInt32(ESLFormat[i + 16]));
                            Color fontcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 9]), Convert.ToInt32(ESLFormat[i + 10]), Convert.ToInt32(ESLFormat[i + 11]), Convert.ToInt32(ESLFormat[i + 12]));
                            mytrydata(ESLFormat[i - 1], ESLFormat[i], special_offer, Convert.ToInt32(ESLFormat[i + 2]), Convert.ToInt32(ESLFormat[i + 3]), Convert.ToInt32(ESLFormat[i + 4]), Convert.ToInt32(ESLFormat[i + 5]), ESLFormat[i + 6], float.Parse(ESLFormat[i + 7]), ESLFormat[i + 8], fontcolor, backcolor, bmp);

                        }

                    }
                }

                if (ESLFormat[i] == "Qrcode 網址")
                {
                    if (ESLFormat[i + 1] == "Qrcode 網址")
                    {
                        Bitmap bqr2 = new Bitmap(Convert.ToInt32(ESLFormat[i + 2]), Convert.ToInt32(ESLFormat[i + 3]));
                        BarcodeWriter qr2 = new BarcodeWriter();
                        qr2.Format = BarcodeFormat.QR_CODE;
                        qr2.Options.Width = Convert.ToInt32(ESLFormat[i + 2]);
                        qr2.Options.Height = Convert.ToInt32(ESLFormat[i + 3]);
                        qr2.Options.Margin = 0;
                        if (!qrcode.Equals(""))
                        {
                            bqr2 = qr2.Write(qrcode);
                        }

                        if (!qrcode.Equals(""))
                        {
                            Console.WriteLine("x:" + Convert.ToInt32(ESLFormat[i + 4]) + "    y:" + Convert.ToInt32(ESLFormat[i + 5]));
                            bmp = mSmcDataToImage.ConvertImageToImage(bmp, bqr2, Convert.ToInt32(ESLFormat[i + 4]), Convert.ToInt32(ESLFormat[i + 5])); //QRcode
                        }

                    }


                }
                if ( ESLFormat[i] == "ESL ID")
                {
                    if (ESLFormat[i -1 ] == "B")
                    {
                        Console.WriteLine("ESL ID");
                        Bitmap bar2 = new Bitmap(Convert.ToInt32(ESLFormat[i + 2]), Convert.ToInt32(ESLFormat[i + 3]));
                        Console.WriteLine("ESL ID1"+ ESLFormat[i]);
                        Console.WriteLine("ESL ID2" + ESLFormat[i+1]);
                        Console.WriteLine("ESL ID3" + ESLFormat[i+2]);
                        Console.WriteLine("ESL ID4" + ESLFormat[i+3]);
                        Console.WriteLine("ESL ID5" + ESLFormat[i+4]);

                        Console.WriteLine("ESL ID6" + ESLFormat[i+5]);
                        BarcodeWriter barcode_w2 = new BarcodeWriter();       // 建立條碼物件
                        barcode_w2.Format = BarcodeFormat.CODE_93;            // 條碼類別.
                        barcode_w2.Options.Width = Convert.ToInt32(ESLFormat[i + 2]);
                        barcode_w2.Options.Height = Convert.ToInt32(ESLFormat[i + 3]);
                        barcode_w2.Options.PureBarcode = true;               // 顯示條碼字串


                        if (!use_address.Equals(""))
                        {
                            if (use_address.Length > 12)
                            {
                                string[] ad = use_address.Split(',');
                                bar2 = barcode_w2.Write(ad[0]);
                            }
                            else
                            {
                                bar2 = barcode_w2.Write(use_address);
                            }
                            bmp = mSmcDataToImage.ConvertImageToImage(bmp, bar2, Convert.ToInt32(ESLFormat[i + 4]), Convert.ToInt32(ESLFormat[i + 5])); //Barcode
                            //Console.WriteLine(bar.Width+"   "+bar.Height);
                        }
                    }


                }
                if (ESLFormat[i] == "商品條碼")
                {
                    if (ESLFormat[i + 1] == "商品條碼")
                    {
                        //  Console.WriteLine(fontName + title + brand + specification + price + special_offer + barcode + qrcode);
                        Bitmap bar2 = new Bitmap(Convert.ToInt32(ESLFormat[i + 2]), Convert.ToInt32(ESLFormat[i + 3]));

                        BarcodeWriter barcode_w2 = new BarcodeWriter();       // 建立條碼物件
                        barcode_w2.Format = BarcodeFormat.CODE_93;            // 條碼類別.
                        barcode_w2.Options.Width = Convert.ToInt32(ESLFormat[i + 2]);
                        barcode_w2.Options.Height = Convert.ToInt32(ESLFormat[i + 3]);
                        barcode_w2.Options.PureBarcode = true;               // 顯示條碼字串
                        if (!barcode.Equals(""))
                        {
                            bar2 = barcode_w2.Write(barcode);
                            bmp = mSmcDataToImage.ConvertImageToImage(bmp, bar2, Convert.ToInt32(ESLFormat[i + 4]), Convert.ToInt32(ESLFormat[i + 5])); //Barcode
                            //Console.WriteLine(bar.Width+"   "+bar.Height);
                        }
                    }


                }

                if (ESLFormat[i] == "Header")
                {
                    Color backcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 14]), Convert.ToInt32(ESLFormat[i + 15]), Convert.ToInt32(ESLFormat[i + 16]), Convert.ToInt32(ESLFormat[i + 17]));
                    Color fontcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 10]), Convert.ToInt32(ESLFormat[i + 11]), Convert.ToInt32(ESLFormat[i + 12]), Convert.ToInt32(ESLFormat[i + 13]));
                    mytrydata(ESLFormat[i], ESLFormat[i + 1], ESLFormat[i + 2], Convert.ToInt32(ESLFormat[i + 3]), Convert.ToInt32(ESLFormat[i + 4]), Convert.ToInt32(ESLFormat[i + 5]), Convert.ToInt32(ESLFormat[i + 6]), ESLFormat[i + 7], float.Parse(ESLFormat[i + 8]), ESLFormat[i + 9], fontcolor, backcolor, bmp);
                    /*  L1.Text = " 促銷價 ";
                      L1.Font = new Font(fontName, 11, FontStyle.Regular);
                      L1.BackColor = Color.Black;
                      if (!special_offer.Equals(""))
                      {
                          bmp = mSmcDataToImage.ConvertBoxToImage(bmp, L1, Color.White, 135, 30);
                      }*/
                    /*    L1.Text = ESLFormat[i+2];
                         L1.TextAlign = ContentAlignment.MiddleCenter;
                         if (Convert.ToInt32(ESLFormat[i + 9]) == 0)
                             L1.Font = new Font(ESLFormat[i + 7], float.Parse(ESLFormat[i + 8]), FontStyle.Regular);
                         else if (Convert.ToInt32(ESLFormat[i + 9]) == 1)
                             L1.Font = new Font(ESLFormat[i + 7], float.Parse(ESLFormat[i + 8]), FontStyle.Bold);
                         else if (Convert.ToInt32(ESLFormat[i + 9]) == 2)
                             L1.Font = new Font(ESLFormat[i + 7], float.Parse(ESLFormat[i + 8]), FontStyle.Italic);
                         else if (Convert.ToInt32(ESLFormat[i + 9]) == 4)
                             L1.Font = new Font(ESLFormat[i + 7], float.Parse(ESLFormat[i + 8]), FontStyle.Underline);
                         else if (Convert.ToInt32(ESLFormat[i +9]) == 8)
                             L1.Font = new Font(ESLFormat[i + 7], float.Parse(ESLFormat[i + 8]), FontStyle.Strikeout);
                         else if (Convert.ToInt32(ESLFormat[i + 9]) == 3)
                             L1.Font = new Font(ESLFormat[i + 7], float.Parse(ESLFormat[i + 8]), FontStyle.Bold);

                         L1.BackColor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 14]), Convert.ToInt32(ESLFormat[i + 15]), Convert.ToInt32(ESLFormat[i + 16]), Convert.ToInt32(ESLFormat[i + 17]));

                         if (!ESLFormat[i + 1].Equals(""))
                         {
                             Color fontcolor = Color.FromArgb(Convert.ToInt32(ESLFormat[i + 10]), Convert.ToInt32(ESLFormat[i + 11]), Convert.ToInt32(ESLFormat[i + 12]), Convert.ToInt32(ESLFormat[i + 13]));
                             bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, fontcolor, Convert.ToInt32(ESLFormat[i + 5]), Convert.ToInt32(ESLFormat[i + 6]));
                         }*/
                }




            }

            return bmp;
        }

        public Bitmap writeIDimage(string barcode)
        {
            Bitmap bmp = new Bitmap(212, 104);
            using (Graphics graphics = Graphics.FromImage(bmp))
            {
                graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, bw, bh);
            }
            //標頭
            TextBox t1 = new TextBox();
            t1.Text = barcode;
            t1.Font = new Font("Calibri", 20, FontStyle.Bold);
            t1.TextAlign = HorizontalAlignment.Center; //置中
            t1.BorderStyle = BorderStyle.FixedSingle;
            t1.Width = 206;
            t1.Height = 50;
            bmp = mSmcDataToImage.ConvertTextToImage(bmp, t1, Color.Red, 3, 8);
            try
            {
                Console.WriteLine("barcode"+ barcode);
                Bitmap bar = new Bitmap(90, 40);
                BarcodeWriter barcode_w = new BarcodeWriter();       // 建立條碼物件
                barcode_w.Format = BarcodeFormat.CODE_93;            // 條碼類別.
                barcode_w.Options.Width = 90;
                barcode_w.Options.Height = 50;
                barcode_w.Options.PureBarcode = true;               // 顯示條碼字串
                if (!barcode.Equals(""))
                {
                    bar = barcode_w.Write(barcode);
                    //Console.WriteLine(bar.Width+"   "+bar.Height);
                }

                bmp = mSmcDataToImage.ConvertImageToImage(bmp, bar, 30, 40); //Barcode
            }

            catch (Exception ex)
            {
                throw;
            }

            return bmp;
        }
        private void mytrydata(string Tag, string name, string text, int width, int height, int x, int y, string fontname, float fontsize, string fontstyle, Color fontcolor, Color backcolor, Bitmap bmp) {
            if (Tag == "T")
            {
                TextBox L1 = new TextBox();
                L1.Text = text;
                L1.AutoSize = false;
                L1.TextAlign = HorizontalAlignment.Center;
                L1.Dock = DockStyle.Fill;
                L1.Font = new Font(fontname, fontsize);
                if (fontstyle.Contains("Regular"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Regular);
                if (fontstyle.Contains("Bold"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Bold);
                if (fontstyle.Contains("Italic"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Italic);
                if (fontstyle.Contains("Underline"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Underline);
                if (fontstyle.Contains("Strikeout"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Strikeout);

                L1.Width = width;
                L1.Height = height;
                if (!name.Equals(""))
                {
                    //       Console.WriteLine("55555555555" + backcolor);

                    bmp = mSmcDataToImage.ConvertTextToImage(bmp, L1, fontcolor, x, y);
                }
            }
            else if (Tag == "L")
            {
                Label L1 = new Label();
                if (name != "售價" && name != "促銷價")
                {
                    L1.Text = name + ":" + text;

                }
                else
                {
                    L1.Text = text;
                    L1.TextAlign = ContentAlignment.MiddleCenter;
                }

                L1.AutoSize = false;
               // L1.TextAlign = ContentAlignment.MiddleCenter;
                L1.Dock = DockStyle.Fill;
                L1.Font = new Font(fontname, fontsize);
                if (fontstyle.Contains("Regular"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Regular);
                if (fontstyle.Contains("Bold"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Bold);
                if (fontstyle.Contains("Italic"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Italic);
                if (fontstyle.Contains("Underline"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Underline);
                if (fontstyle.Contains("Strikeout"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Strikeout);

                L1.Width = width;
                L1.Height = height;
                L1.BackColor = backcolor;
                if (!name.Equals(""))
                {
                    //  Console.WriteLine("55555555555" + backcolor);
                    if (backcolor != Color.FromArgb(255, 255, 255))
                    {
                        Console.WriteLine("NONOONOONOOWhite");
                        L1.TextAlign = ContentAlignment.MiddleCenter;
                        bmp = ConvertBoxToImage(bmp, L1, fontcolor, x, y);
                    }
                    else
                    {
                        Console.WriteLine("Color.WhiteColor.WhiteColor.WhiteColor.White");
                        bmp = ConvertTextToImage(bmp, L1, fontcolor, x, y);

                    }




                }
            }
            else
            {
                Label L1 = new Label();
                L1.Text = text;


                if (text != "售價:")
                {

                    L1.TextAlign = ContentAlignment.MiddleCenter;

                }

                L1.AutoSize = false;
                L1.Dock = DockStyle.Fill;
                L1.Font = new Font(fontname, fontsize);
                if (fontstyle.Contains("Regular"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Regular);
                if (fontstyle.Contains("Bold"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Bold);
                if (fontstyle.Contains("Italic"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Italic);
                if (fontstyle.Contains("Underline"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Underline);
                if (fontstyle.Contains("Strikeout"))
                    L1.Font = new Font(L1.Font, L1.Font.Style | FontStyle.Strikeout);


                L1.Width = width;
                L1.Height = height;
                L1.BackColor = backcolor;
                if (!name.Equals(""))
                {
                    //  Console.WriteLine("55555555555" + backcolor);
                    if (backcolor != Color.FromArgb(255,255,255))
                    {
                        Console.WriteLine("NONOONOONOOWhite");
                        L1.TextAlign = ContentAlignment.MiddleCenter;
                        bmp = ConvertBoxToImage(bmp, L1, fontcolor, x, y);
                    }
                    else
                    {

                        Console.WriteLine("Color.WhiteColor.WhiteColor.WhiteColor.White");
                        bmp = ConvertTextToImage(bmp, L1, fontcolor, x, y);

                    }




                }
            }
          

        }


        public Bitmap ConvertBoxToImage(Bitmap mbmp, Label label, Color textcolor, int x, int y)
        {
            using (Graphics graphic = Graphics.FromImage(mbmp))
            {
                graphic.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;
                SolidBrush solidBrush = new SolidBrush(label.BackColor);
                StringFormat stringFormat = new StringFormat();
                stringFormat.Alignment = StringAlignment.Center;
                stringFormat.LineAlignment = StringAlignment.Center;
                Rectangle rectangle = new Rectangle(x,y, label.Width, label.Height);
                graphic.FillRectangle(solidBrush, rectangle);
                graphic.DrawString(label.Text, label.Font, new SolidBrush(textcolor), (float)(x+ label.Width/2), (float)(y+ label.Height/2), stringFormat);
                graphic.Flush();
                graphic.Dispose();
            }
            return mbmp;
        }

        public Bitmap ConvertImageToImage(Bitmap mbmp, Bitmap img, int x, int y)
        {
            for (int i = 0; i < img.Width; i++)
            {
                for (int j = 0; j < img.Height; j++)
                {
                    Color pixel = img.GetPixel(i, j);
                    if ((pixel.R + pixel.B + pixel.G) / 3 < 180)
                    {
                        mbmp.SetPixel(i + x, j + y, Color.FromArgb(0, 0, 0));
                    }
                }
            }
            return mbmp;
        }


        public Bitmap setESLimage_29(string MacAddress, string battery)
        {
            Bitmap bmp = new Bitmap(296, 128);
            using (Graphics graphics = Graphics.FromImage(bmp))
            {
                graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, 296, 128);
            }

            Bitmap bar = new Bitmap(290, 50);
            BarcodeWriter barcode_w = new BarcodeWriter();
            barcode_w.Format = BarcodeFormat.CODE_93;
            barcode_w.Options.Width = 296;
            barcode_w.Options.Height = 40;
            barcode_w.Options.PureBarcode = true;
            bar = barcode_w.Write(MacAddress);
            bmp = mSmcDataToImage.ConvertImageToImage(bmp, bar, 4, 85); //QRcode

            TextBox t1 = new TextBox();
            t1.Text = "SMC ESL  " + battery + " V";
            t1.Font = new Font("Cambria", 26, FontStyle.Bold);
            t1.TextAlign = HorizontalAlignment.Center; //置中
            t1.BorderStyle = BorderStyle.FixedSingle;
            t1.Width = 280;
            t1.Height = 25;
            bmp = mSmcDataToImage.ConvertTextToImage(bmp, t1, Color.Red, 1, 0);

            String StrName = String.Format("{0}", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            t1.Text = StrName;
            t1.Font = new Font("Cambria", 15, FontStyle.Bold);
            t1.TextAlign = HorizontalAlignment.Center; //置中
            t1.BorderStyle = BorderStyle.FixedSingle;
            t1.Width = 280;
            t1.Height = 25;
            bmp = mSmcDataToImage.ConvertTextToImage(bmp, t1, Color.Red, 2, 35);


            t1.Text = MacAddress;
            t1.Font = new Font("Cambria", 20, FontStyle.Bold);
            t1.TextAlign = HorizontalAlignment.Center; //置中
            t1.BorderStyle = BorderStyle.FixedSingle;
            t1.Width = 280;
            t1.Height = 25;
            bmp = mSmcDataToImage.ConvertTextToImage(bmp, t1, Color.Black, 2, 53);




            return bmp;
        }


        public Bitmap setESLimage_42(string MacAddress, string battery)
        {
            Bitmap bmp = new Bitmap(400, 300);
            using (Graphics graphics = Graphics.FromImage(bmp))
            {
                graphics.FillRectangle(new SolidBrush(Color.White), 0, 0, 400, 300);
            }



            TextBox t1 = new TextBox();
            t1.Text = "SMC ESL  " + battery + " V";
            t1.Font = new Font("Cambria", 38, FontStyle.Bold);
            t1.TextAlign = HorizontalAlignment.Center; //置中
            t1.BorderStyle = BorderStyle.FixedSingle;
            t1.Width = 380;
            t1.Height = 40;
            bmp = mSmcDataToImage.ConvertTextToImage(bmp, t1, Color.Red, 1, 0);

            String StrName = String.Format("{0}", DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            t1.Text = StrName;
            t1.Font = new Font("Cambria", 20, FontStyle.Bold);
            t1.TextAlign = HorizontalAlignment.Center; //置中
            t1.BorderStyle = BorderStyle.FixedSingle;
            t1.Width = 380;
            t1.Height = 40;
            bmp = mSmcDataToImage.ConvertTextToImage(bmp, t1, Color.Red, 2, 60);


            t1.Text = MacAddress;
            t1.Font = new Font("Cambria", 26, FontStyle.Bold);
            t1.TextAlign = HorizontalAlignment.Center; //置中
            t1.BorderStyle = BorderStyle.FixedSingle;
            t1.Width = 380;
            t1.Height = 40;
            bmp = mSmcDataToImage.ConvertTextToImage(bmp, t1, Color.Black, 2, 85);

            Bitmap bar = new Bitmap(400, 80);
            BarcodeWriter barcode_w = new BarcodeWriter();
            barcode_w.Format = BarcodeFormat.CODE_93;
            barcode_w.Options.Width = 400;
            barcode_w.Options.Height = 80;
            barcode_w.Options.PureBarcode = true;
            bar = barcode_w.Write(MacAddress);
            bmp = mSmcDataToImage.ConvertImageToImage(bmp, bar, 4, 210); //QRcode


            Bitmap qr = new Bitmap(400, 70);

            barcode_w.Format = BarcodeFormat.QR_CODE;
            barcode_w.Options.Width = 400;
            barcode_w.Options.Height = 70;
            barcode_w.Options.PureBarcode = true;
            bar = barcode_w.Write("http://www.smartchip.com.tw/");
            bmp = mSmcDataToImage.ConvertImageToImage(bmp, bar, 4, 130); //QRcode



            return bmp;
        }


        public Bitmap ConvertTextToImage(Bitmap mbmp, Label label, Color textcolor, int x, int y)
        {
            using (Graphics graphic = Graphics.FromImage(mbmp))
            {
                StringFormat stringFormat = new StringFormat();
          //      Console.WriteLine("label.TextAlign"+ label.TextAlign.ToString());
                if (label.TextAlign == ContentAlignment.MiddleCenter)
                {
               //     Console.WriteLine("UUUUUUUUYYYYYYYY");
                    stringFormat.Alignment = StringAlignment.Center;
                 //   stringFormat.LineAlignment = StringAlignment.Center;
                    graphic.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;
                    graphic.DrawString(label.Text, label.Font, new SolidBrush(textcolor), (float)(x+ label.Width/2), (float)y, stringFormat);
                }
                else
                {
                    graphic.TextRenderingHint = TextRenderingHint.SingleBitPerPixelGridFit;
                    graphic.DrawString(label.Text, label.Font, new SolidBrush(textcolor), (float)x, (float)y);
                }
                graphic.Flush();
                graphic.Dispose();
            }
            return mbmp;
        }



    }
}
