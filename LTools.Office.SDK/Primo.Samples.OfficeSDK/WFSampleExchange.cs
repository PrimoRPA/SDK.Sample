using LTools.Common.Model;
using LTools.Common.UIElements;
using LTools.Office;
using LTools.SDK;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Primo.Samples.OfficeSDK
{
    public class WFSampleExchange : LTools.Office.SDK.PrimoComponentExchange<WFSampleExchangeBase>
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

        public WFSampleExchange(IWFContainer container) : base(container)
        {
            sdkComponentHelp = "SDK help";
            sdkComponentName = "Office SDK Exchange Test";
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

            var email = new EmailMessage(this.Driver.InteropExchangeApp);
            email.ToRecipients.Add("test@test.ru");
            email.Subject = "test";
            email.Body = p1;
            email.Send();

            return new ExecutionResult() { SuccessMessage = "Done", IsSuccess = true };
        }
    }
}
