using System;
using System.Drawing;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using VstoMultiLanguageDemo.Properties;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Linq;

namespace VstoMultiLanguageDemo
{
    [ComVisible(true)]
    public class RibbonHandler : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        // remember ribbon to save
        public void OnRibbonLoad(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void Invalidate()
        {
            _ribbon.Invalidate();
        }

        public void OnCommand(Office.IRibbonControl control)
        {
            MessageBox.Show(string.Format(Resources.MessageBoxText, control.Id), Resources.MessageBoxCaption);
        }

        // callback to get the label for the ribbon button from resoures
        public string GetRibbonLabel(Office.IRibbonControl control) 
            => Resources.ResourceManager.GetString($"{control.Id}_Label", Resources.Culture);

        // callback to get the description for the ribbon button from resoures
        public string GetRibbonDescription(Office.IRibbonControl control) 
            => Resources.ResourceManager.GetString($"{control.Id}_Description", Resources.Culture);

        // callback to get the image for the ribbon button from resoures
        public Bitmap GetRibbonImage(Office.IRibbonControl control) 
            => (Bitmap)Resources.ResourceManager.GetObject(control.Id, Resources.Culture);

        #region IIRibbonExtensibility
        public string GetCustomUI(string ribbonId)
            => Resources.Ribbon;
        #endregion
    }

    public partial class ThisAddIn 
    {
        private RibbonHandler _ribbonHandler = new RibbonHandler();

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() 
            => _ribbonHandler;

        private void InternalStartup()
        {
            var lcid = Application.LanguageSettings.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI];
            var culture = new CultureInfo(lcid);

            var languages = new[] { "de", "ru" };
            if (languages.Any(language => language == culture.TwoLetterISOLanguageName))
            {
                Resources.Culture = culture;
                _ribbonHandler.Invalidate();
            }
        }

    }
}
