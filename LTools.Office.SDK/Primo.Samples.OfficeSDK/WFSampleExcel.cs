using LTools.Common.Model;
using LTools.Common.UIElements;
using LTools.SDK;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace Primo.Samples.OfficeSDK
{
    public class WFSampleExcel : LTools.Office.SDK.PrimoComponentExcel<WFSampleExcelBase>
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

        public WFSampleExcel(IWFContainer container) : base(container)
        {
            sdkComponentHelp = "SDK help";
            sdkComponentName = "Office SDK Excel Test";
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
                    this.Driver.DXExcelApp.Worksheets[0].Cells["A1"].Value = p1;
                    break;
                case LTools.Office.Model.InteropTypes.Interop:
                    (this.Driver.InteropExcelApp.Worksheets[1] as Excel.Worksheet).Range["A1"].Cells[1, 1] = p1;
                    break;
            }

            return new ExecutionResult() { SuccessMessage = "Done", IsSuccess = true };
        }
    }
}
