using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace Drop_VC
{
    public partial class Ribbon1
    {

        string content = "table of content";
        int cutOff = 13;
        string fontReq = "";
        bool div = false;
        int lineDropped = 3;
        int wDmargin = 1;
        string fontDropped = "";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            titleFont.Text = "";
            titleSize.SelectedItemIndex = 12;
            divider.Checked = false;
            fontDrop.Text = "";
            lineDrop.SelectedItemIndex = 2;
            margin.Checked = false;
        }

        // Main drop function
        private void drop_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("Hello");
            Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Paragraphs paras = doc.Paragraphs;

            int i = 1, skip;
            bool drop = false;
            string txt, rep;

            while (true)
            {
                if (paras[i].Range.Text.Length > 1)
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
            }
            return i;
        }

        // Is checking if the paragraph is a title
        private bool isTitle(Word.Paragraph para)
        {

            string txt = para.Range.Text;
            
            if (txt.ToLower().Contains(content)) return false;// Is table of content

            int size = (int) para.Range.Font.Size;

            if (fontReq != "")
            {
                if (para.Range.Font.Name.Contains(fontReq))
                {
                    return (size >= cutOff);
                }
                else return false;
            } else if (size >= cutOff) return true;

            //MessageBox.Show("Size is " + size + " Cutoff is " + cutOff);
            return false;
        }



        // Does the Drop Cap
        private void dropTxt(Word.Paragraph para)
        {
            Word.DropCap dropCap = para.DropCap;
            if (fontDropped != "") dropCap.FontName = fontDropped;
            dropCap.Position = (Word.WdDropPosition) wDmargin;
            dropCap.LinesToDrop = lineDropped;
            dropCap.DistanceFromText = 1;
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
                if (paras[i].Range.Text.Length > 10)
                {
                    //MessageBox.Show("Paragraph " + i + ". With length " + paras[i].Range.Text.Length);
                    dropCap = paras[i].DropCap;
                    if (dropCap != null) dropCap.Clear();
                }
                if (i == doc.Paragraphs.Count) break;
                i++;
            }

        }


        // Title function
        private void titleFont_Change(object sender, RibbonControlEventArgs e)
        {
            RibbonEditBox obj = (RibbonEditBox) sender;
            fontReq = obj.Text;
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
