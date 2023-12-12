using DevExpress.XtraRichEdit.API.Native.Implementation;
using LTools.Common.Model;
using LTools.Common.UIElements;
using LTools.SDK;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Primo.Samples.OfficeSDK
{
    public class WFSampleWord : LTools.Office.SDK.PrimoComponentWord<WFSampleWordBase>
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

        public WFSampleWord(IWFContainer container) : base(container)
        {
            sdkComponentHelp = "SDK help";
            sdkComponentName = "Office SDK Word Test";
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

            switch (this.Driver.InteropType)
            {
                case LTools.Office.Model.InteropTypes.DX:
                    this.Driver.DXCurrentDoc.AppendText(p1);
                    break;
                case LTools.Office.Model.InteropTypes.Interop:
                    this.Driver.InteropWordApp.ActiveDocument.Characters.Last.Select();
                    this.Driver.InteropWordApp.Selection.Collapse();
                    this.Driver.InteropWordApp.Selection.TypeText(p1);
                    break;
            }

            return new ExecutionResult() { SuccessMessage = "Done", IsSuccess = true };
        }
    }
}
