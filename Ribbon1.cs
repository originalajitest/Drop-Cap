using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;

namespace Drop_VC
{
    public partial class Ribbon1
    {

        //string content = "table of content";
        int cutOff = 13;
        bool fontReq = false;
        bool[] fontReqStat;
        string fontReq1 = "";
        string fontReq2 = "";
        string fontReq3 = "";
        string fontReq4 = "";
        bool div = false;
        int lineDropped = 3;
        int wDmargin = 1;
        string fontDropped = "";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            titleFont1.Text = "";
            titleFont2.Text = "";
            titleFont3.Text = "";
            titleFont4.Text = "";
            titleSize.SelectedItemIndex = 12;
            divider.Checked = false;
            fontDrop.Text = "";
            lineDrop.SelectedItemIndex = 2;
            margin.Checked = false;


            fontReqStat = new bool[4];
            for (int i = 0; i < 4; i++)
            {
                fontReqStat[i] = false;
            }
        }

        // Main drop function
        private void drop_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("Hello");
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraphs paras = doc.Paragraphs;

            //MessageBox.Show(paras[1].Range.Font.Name);

            int i = 1, skip;
            bool drop = false;
            string txt, rep;

            while (true)
            {
                if (paras[i].Range.Text.Length > 7) // 7 to allow for footers and other special stuff that cant be dropped
                {
                    if (isTitle(paras[i]))
                    {
                        /*
                        if (paras[i].Range.Text.Length < 20)
                        {
                            MessageBox.Show("Found a title, Text:" + paras[i].Range.Text.Substring(0, paras[i].Range.Text.Length));
                        }
                        else
                        {
                            MessageBox.Show("Found a title, Text:" + paras[i].Range.Text.Substring(0, 20));
                        }
                        */

                        drop = true;

                        if (div) i++;

                    }
                    else if (drop)
                    {
                        //MessageBox.Show("Found a paragraph to drop, Text: " + paras[i].Range.Text.Substring(0, 20));
                        skip = specialChars(paras[i]);

                        txt = paras[i].Range.Text;
                        paras[i].Range.Text = txt.Substring(skip);
                        dropTxt(paras[i]);
                        rep = txt.Substring(0, skip + 1);
                        paras[i].Range.Text = rep;

                        drop = false;
                        i++;
                    }
                }

                if (i == doc.Paragraphs.Count) break;
                i++;
            }

            /*
                string txt = para.Range.Text;
                string temp = txt.Substring(1);
                para.Range.Text = temp;
                char x = txt.ToCharArray()[0];
                MessageBox.Show("Removed first char: " + x + " in int form " + (int)x);

                Word.DropCap dropCap = paras[1].DropCap;
                dropCap.Position = Word.WdDropPosition.wdDropNormal;
                dropCap.LinesToDrop = 3;
                dropCap.DistanceFromText = 1;
                dropCap.Enable();

                MessageBox.Show("Dropped");

                string rep = txt.Substring(0,2);
                paras[1].Range.Text = rep;
            */

            //MessageBox.Show("Done?");
        }


        // Getting number of non alphabetic characters from the front of the paragraph
        private int specialChars(Word.Paragraph para)
        {
            char[] txt = para.Range.Text.ToCharArray();

            int i = 0, temp;
            while (true)
            {
                temp = txt[i];
                if (temp > 90) temp -= 32;
                //MessageBox.Show("i: " + i + " Char: " + txt[i] + " Int: " + (int)txt[i] + " Used Int: " + temp);
                if (temp >= 65 && temp <= 90) break;
                i++;
                if (i == txt.Length) return 0;
            }
            return i;
        }

        // Is checking if the paragraph is a title
        private bool isTitle(Word.Paragraph para)
        {

            string txt = para.Range.Text;

            //if (txt.ToLower().Contains(content)) return false;// Is table of content
            // Had to remove table of content as it broke when multiple font check
            // Drop was passed along until it ended on table of content

            /*
            int style = (int) ((WdBuiltinStyle) para.get_Style());// Is breaking
            if (style == -2 || style == -3 || style == -63 || (style >= -28 && style <= -20)) return true;
            */

            int size = (int) para.Range.Font.Size;

            if (size >= cutOff)
            {
                if (fontReq)
                {
                    if (fontReqStat[0] && para.Range.Font.Name.Contains(fontReq1)) return true;
                    if (fontReqStat[1] && para.Range.Font.Name.Contains(fontReq2)) return true;
                    if (fontReqStat[2] && para.Range.Font.Name.Contains(fontReq3)) return true;
                    if (fontReqStat[3] && para.Range.Font.Name.Contains(fontReq4)) return true;
                    return false;
                }
                else return true;
            }

            //MessageBox.Show("Size is " + size + " Cutoff is " + cutOff);
            return false;
        }



        // Does the Drop Cap
        private void dropTxt(Word.Paragraph para)
        {
            //MessageBox.Show(para.Range.Text);
            Word.DropCap dropCap = para.DropCap;
            dropCap.Position = (Word.WdDropPosition) wDmargin;
            dropCap.LinesToDrop = lineDropped;
            dropCap.DistanceFromText = 1;
            if (fontDropped != "") dropCap.FontName = fontDropped;
            dropCap.Enable();
        }



        // Function to remove all drop cap, doesn't need to check just do it on all.
        private void rmv_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("Remove");

            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraphs paras = doc.Paragraphs;
            Word.DropCap dropCap;

            int i = 1;
            while (true)
            {
                // At 1 it does it for all page number and that breaks as they cannot have dropCap
                if (paras[i].Range.Text.Length > 7)
                {
                    //MessageBox.Show("Paragraph " + i + ". With length " + paras[i].Range.Text.Length);
                    dropCap = paras[i].DropCap;
                    if (dropCap != null) dropCap.Clear();
                }
                if (i == doc.Paragraphs.Count) break;
                i++;
            }

        }

        private void fontUpdater()
        {
            if (fontReqStat[0] || fontReqStat[1] || fontReqStat[2] || fontReqStat[3]) fontReq = true;
            else fontReq = false;
            //if (fontReq)MessageBox.Show("True");
            //else MessageBox.Show("False");
        }


        // Title function
        private void titleFont1_Change(object sender, RibbonControlEventArgs e)
        {
            RibbonEditBox obj = (RibbonEditBox)sender;
            fontReq1 = obj.Text;
            if (fontReq1 == "") fontReqStat[0] = false;
            else fontReqStat[0] = true;
            fontUpdater();
        }
        private void titleFont2_Change(object sender, RibbonControlEventArgs e)
        {
            RibbonEditBox obj = (RibbonEditBox)sender;
            fontReq2 = obj.Text;
            if (fontReq2 == "") fontReqStat[1] = false;
            else fontReqStat[1] = true;
            fontUpdater();
        }
        private void titleFont3_Change(object sender, RibbonControlEventArgs e)
        {
            RibbonEditBox obj = (RibbonEditBox)sender;
            fontReq3 = obj.Text;
            if (fontReq3 == "") fontReqStat[2] = false;
            else fontReqStat[2] = true;
            fontUpdater();
        }
        private void titleFont4_Change(object sender, RibbonControlEventArgs e)
        {
            RibbonEditBox obj = (RibbonEditBox)sender;
            fontReq4 = obj.Text;
            if (fontReq4 == "") fontReqStat[3] = false;
            else fontReqStat[3] = true;
            fontUpdater();
        }

        private void titleSize_change(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDown obj = (RibbonDropDown) sender;
            cutOff = int.Parse(obj.SelectedItem.Label);
        }

        private void divFunc(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox obj = (RibbonCheckBox) sender;
            div = obj.Checked;
        }




        // Drop Cap functions
        private void font_Drop_Change(object sender, RibbonControlEventArgs e)
        {
            RibbonEditBox obj = (RibbonEditBox) sender;
            fontDropped = obj.Text;
        }

        private void lineDrop_Change(object sender, RibbonControlEventArgs e)
        {
            RibbonDropDown obj = (RibbonDropDown) sender;
            lineDropped = int.Parse(obj.SelectedItem.Label);
        }

        private void margin_Change(object sender, RibbonControlEventArgs e)
        {
            RibbonCheckBox obj = (RibbonCheckBox) sender;
            if (obj.Checked) wDmargin = 2;
            else wDmargin = 1;
        }
    }
}
