using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Drawing;


// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace ExcelAddInDemo
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        private IRibbonUI ribbon;
        private List<string> fontNames;
        private Excel.Range copiedRange;

        public Ribbon()
        {
            fontNames = new List<string>();
            foreach (FontFamily font in FontFamily.Families)
            {
                fontNames.Add(font.Name);
            }
        }

        public void OnCutButtonClick(IRibbonControl control)
        {
            CutSelection();
        }

        public void OnCopyButtonClick(IRibbonControl control)
        {
            CopySelection();
        }

        public void OnPasteButtonClick(IRibbonControl control)
        {
            PasteSelection();
        }
        private void CopySelection()
        {
            try
            {
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                Excel.Range selectedRange = excelApp.Selection as Excel.Range;

                if (selectedRange != null)
                {
                    selectedRange.Copy();
                    copiedRange = selectedRange;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No range selected for copying.");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Exception during copy: " + ex.Message);
            }
        }

        private void CutSelection()
        {
            try
            {
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                Excel.Range selectedRange = excelApp.Selection as Excel.Range;

                if (selectedRange != null)
                {
                    selectedRange.Cut();
                    copiedRange = selectedRange;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No range selected for cutting.");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Exception during cut: " + ex.Message);
            }
        }

        private void PasteSelection()
        {
            try
            {
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                Excel.Range selectedRange = excelApp.Selection as Excel.Range;

                if (selectedRange != null && copiedRange != null)
                {
                    selectedRange.PasteSpecial(Excel.XlPasteType.xlPasteAll);
                    excelApp.CutCopyMode = 0;  // Clear the clipboard mode
                    copiedRange = null;        // Reset the copied range after pasting
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No range selected for pasting or no data available.");
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                System.Diagnostics.Debug.WriteLine("COMException during paste: " + ex.Message);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Exception during paste: " + ex.Message);
            }
        }


        // Font Methods
        public void OnBoldButtonClick(IRibbonControl control)
        {
            ToggleBold();
        }

        public void OnItalicButtonClick(IRibbonControl control)
        {
            ToggleItalic();
        }

        public void OnUnderlineButtonClick(IRibbonControl control)
        {
            ToggleUnderline();
        }

        public void OnFontSizeChange(IRibbonControl control, string text)
        {
            if (float.TryParse(text, out float size))
            {
                SetFontSize(size);
            }
        }

        public void OnFontNameChange(IRibbonControl control, string selectedId, int selectedIndex)
        {
            SetFontName(fontNames[selectedIndex]);
        }

        public void OnFontColorButtonClick(IRibbonControl control)
        {
            SetFontColor();
        }

        public int GetFontNameCount(IRibbonControl control)
        {
            return fontNames.Count;
        }

        public string GetFontNameLabel(IRibbonControl control, int index)
        {
            return fontNames[index];
        }

        private void ToggleBold()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;

            if (selectedRange != null)
            {
                Excel.Font font = selectedRange.Font;
                font.Bold = !font.Bold;
            }
        }

        private void ToggleItalic()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;

            if (selectedRange != null)
            {
                Excel.Font font = selectedRange.Font;
                font.Italic = !font.Italic;
            }
        }

        private void ToggleUnderline()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;

            if (selectedRange != null)
            {
                Excel.Font font = selectedRange.Font;
                if ((int)font.Underline == (int)Excel.XlUnderlineStyle.xlUnderlineStyleNone)
                {
                    font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle; // Apply underline
                }
                else
                {
                    font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone; // Remove underline
                }
            }
        }


        private void SetFontSize(float size)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;

            if (selectedRange != null)
            {
                Excel.Font font = selectedRange.Font;
                font.Size = size;
            }
        }

        private void SetFontName(string name)
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;

            if (selectedRange != null)
            {
                Excel.Font font = selectedRange.Font;
                font.Name = name;
            }
        }

        private void SetFontColor()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Range selectedRange = excelApp.Selection as Excel.Range;

            if (selectedRange != null)
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    Excel.Font font = selectedRange.Font;
                    font.Color = ColorTranslator.ToOle(colorDialog.Color);
                }
            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelAddInDemo.Ribbon.xml");
        }

        

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
