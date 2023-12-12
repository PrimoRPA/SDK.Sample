using LTools.Common.Model;
using LTools.Common.UIElements;
using LTools.Office;
using LTools.SDK;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace Primo.Samples.OfficeSDK
{
    public class WFSampleOutlook : LTools.Office.SDK.PrimoComponentOutlook<WFSampleOutlookBase>
    {
        private string prop1;

        [LTools.Common.Model.Serialization.StoringProperty]
        [LTools.Common.Model.Studio.ValidateReturnScript(DataType = typeof(string))]
        [Category("SDK"), DisplayName("Value")]
        public string Prop1
        {
            get { return this.prop1; }
            set { this.prop1 = value; this.InvokePropertyChanged(this, "Prop1"); }
        }

        public WFSampleOutlook(IWFContainer container) : base(container)
        {
            sdkComponentHelp = "SDK help";
            sdkComponentName = "Office SDK Outlook Test";
            sdkProperties = new List<LTools.Common.Helpers.WFHelper.PropertiesItem>()
            {
                new LTools.Common.Helpers.WFHelper.PropertiesItem() { PropName = "Prop1", PropertyType = LTools.Common.Helpers.WFHelper.PropertiesItem.PropertyTypes.SCRIPT, EditorType = ScriptEditorTypes.NONE, DataType = typeof(string), ToolTip = "SDK Tooltip1", IsReadOnly = false }
            };
            InitClass(container);
            this.Prop1 = this.IsNoCode("Prop1") ? "test text" : "\"test text\"";
        }

        public override ExecutionResult SimpleAction(ScriptingData sd)
        {
            string p1 = GetPropertyValue<string>(this.Prop1, "Prop1", sd);

            var newMail = (Outlook.MailItem)this.Driver.InteropOutlookApp.CreateItem(Outlook.OlItemType.olMailItem);
            newMail.To = "test@test.ru";
            newMail.Subject = p1;
            newMail.Send();

            return new ExecutionResult() { SuccessMessage = "Done", IsSuccess = true };
        }
    }
}
