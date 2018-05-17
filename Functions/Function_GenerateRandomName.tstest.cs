using Telerik.TestingFramework.Controls.KendoUI;
using Telerik.WebAii.Controls.Html;
using Telerik.WebAii.Controls.Xaml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

using ArtOfTest.Common.UnitTesting;
using ArtOfTest.WebAii.Core;
using ArtOfTest.WebAii.Controls.HtmlControls;
using ArtOfTest.WebAii.Controls.HtmlControls.HtmlAsserts;
using ArtOfTest.WebAii.Design;
using ArtOfTest.WebAii.Design.Execution;
using ArtOfTest.WebAii.ObjectModel;
using ArtOfTest.WebAii.Silverlight;
using ArtOfTest.WebAii.Silverlight.UI;

namespace VPS_Web
{

    public class Function_GenerateRandomName : BaseWebAiiTest
    {
        #region [ Dynamic Pages Reference ]

        private Pages _pages;

        /// <summary>
        /// Gets the Pages object that has references
        /// to all the elements, frames or regions
        /// in this project.
        /// </summary>
        public Pages Pages
        {
            get
            {
                if (_pages == null)
                {
                    _pages = new Pages(Manager.Current);
                }
                return _pages;
            }
        }

        #endregion
        
        // Add your test methods here...
        public string RandomName = null;
        
        [CodedStep(@"New Coded Step")]
        public void Function_GenerateRandomName_CodedStep()
        {
            StringBuilder builder = new StringBuilder(); 
            builder.Clear();
            Random random = new Random();  
            char ch;  
            string type = "VPS_";
            type = type.ToUpper();
                        
            builder.Append(type);         
            
            for (int i = 0; i < 15; i++)  
            {  
                ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));  
                builder.Append(ch);  
            }  
            string randomString = builder.ToString();
            //SetExtractedValue("RandomName",randomString);  
            SetExtractedValue("RandomNameShort",randomString.Remove(13));    
        }
    }
}
