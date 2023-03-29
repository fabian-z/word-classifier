using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace word_classifier
{
    public partial class ThisAddIn
    {
        internal enum Classification
        {
            None,
            White,
            Green,
            Amber,
            Red
        }

        private static string classificationProperty = "Classification";

        // User control
        private UserControl _usr;
        // Custom task pane
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;

        private static readonly Dictionary<Classification, string> classificationHeader = new Dictionary<Classification, string> {
            { Classification.White, "Classified: TLP White" },
            { Classification.Green, "Classified: TLP Green" },
            { Classification.Amber, "Classified: TLP Amber" },
            { Classification.Red, "Classified: TLP Red" },
        };

        private static readonly Dictionary<Classification, Microsoft.Office.Interop.Word.WdColor> classificationTextColor = new Dictionary<Classification, Microsoft.Office.Interop.Word.WdColor> {
            { Classification.White, Word.WdColor.wdColorAutomatic },
            { Classification.Green, Word.WdColor.wdColorGreen },
            { Classification.Amber, Word.WdColor.wdColorOrange },
            { Classification.Red, Word.WdColor.wdColorRed },
        };

        private static readonly Dictionary<Classification, string> classificationProperties = new Dictionary<Classification, string> {
            { Classification.White, "TLP:WHITE" },
            { Classification.Green, "TLP:GREEN" },
            { Classification.Amber, "TLP:AMBER" },
            { Classification.Red, "TLP:RED" },
        };

        internal void ToggleTaskPane()
        {
            if (_taskPane != null)
            {
                _taskPane.Visible = !_taskPane.Visible;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);


            //Create an instance of the user control
            _usr = new TaskPane();
            // Connect the user control and the custom task pane 
            _taskPane = CustomTaskPanes.Add(_usr, "Classifier TLP");
            _taskPane.Width = 300;
            _taskPane.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            Classification classification = GetClassification();
            if (classification == Classification.None)
            {
                Cancel = true;
                MessageBox.Show("Please classify document before saving. See the Add-In / Classifier Toolbar Menu.", "Classifier prevented saving", MessageBoxButtons.OK, MessageBoxIcon.Error);
            } else
            {
                // re-apply changes as defense in depth
                Classify(classification);
            }
        }

        Classification GetClassification()
        {
            string classificationString = ReadDocumentProperty(classificationProperty);
            if (!classificationProperties.ContainsValue(classificationString))
            {
                return Classification.None;
            } else
            {
                Classification key = classificationProperties.FirstOrDefault(x => x.Value == classificationString).Key;
                return key;
            }
        }

        internal void Classify(Classification classification)
        {
            SetDocumentProperty(classificationProperty, classificationProperties[classification]);
            foreach (Word.Section section in this.Application.ActiveDocument.Sections)
            {
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Font.Color = classificationTextColor[classification];
                headerRange.Font.Size = 16;
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Text = classificationHeader[classification];
            }
        }

        private string ReadDocumentProperty(string propertyName)
        {
            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)this.Application.ActiveDocument.CustomDocumentProperties;

            foreach (Office.DocumentProperty prop in properties)
            {
                if (prop.Name == propertyName)
                {
                    return prop.Value.ToString();
                }
            }
            return null;
        }

        private void SetDocumentProperty(string propertyName, string value)
        {
            Office.DocumentProperties properties;
            properties = (Office.DocumentProperties)this.Application.ActiveDocument.CustomDocumentProperties;

            if (ReadDocumentProperty(propertyName) != null)
            {
                properties[propertyName].Delete();
            }

            properties.Add(propertyName, false,
                Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                value);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }


        #endregion
    }
}
