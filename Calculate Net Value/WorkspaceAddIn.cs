using System.AddIn;
using System.Drawing;
using System.Windows.Forms;
using RightNow.AddIns.AddInViews;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Text;
using System.Data;
using System.ComponentModel;
using System.Linq;
using System;
using Calculate_Net_Value.RightNowService;

////////////////////////////////////////////////////////////////////////////////
//
// File: WorkspaceAddIn.cs
//
// Comments:
//
// Notes: 
//
// Pre-Conditions: 
//
////////////////////////////////////////////////////////////////////////////////
namespace Calculate_Net_Value
{
    public class WorkspaceAddIn : Panel, IWorkspaceComponent2
    {
        /// <summary>
        /// The current workspace record context.
        /// </summary>
        private IRecordContext _recordContext;
        public static IGlobalContext _globalContext { get; private set; }
        public static IIncident _incidentRecord;
        private static IGenericObject _partsRecord;
        RightNowConnectService _rnConnectService;

        /// <summary>
        /// Default constructor.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        public WorkspaceAddIn(bool inDesignMode, IRecordContext RecordContext, IGlobalContext GlobalContext)
        {
            _recordContext = RecordContext;
            _globalContext = GlobalContext;
            _rnConnectService = RightNowConnectService.GetService(_globalContext);
        }

        #region IAddInControl Members

        /// <summary>
        /// Method called by the Add-In framework to retrieve the control.
        /// </summary>
        /// <returns>The control, typically 'this'.</returns>
        public Control GetControl()
        {
            return this;
        }

        #endregion

        #region IWorkspaceComponent2 Members

        /// <summary>
        /// Sets the ReadOnly property of this control.
        /// </summary>
        public bool ReadOnly { get; set; }

        /// <summary>
        /// Method which is called when any Workspace Rule Action is invoked.
        /// </summary>
        /// <param name="ActionName">The name of the Workspace Rule Action that was invoked.</param>
        public void RuleActionInvoked(string ActionName)
        {
            
            string[] pResults;
            string[] lResults;
            string[] oResults;
            
            decimal totallabor=0;
            decimal totalpartsCost=0;
            decimal totalpartsPrice = 0;
            decimal totalother=0;
            decimal net=0;
            decimal approvalamt = 0;
            

            try
            {
                _incidentRecord = (IIncident)_recordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
                pResults = _rnConnectService.GetParts(_incidentRecord.ID.ToString());
             
                lResults = _rnConnectService.GetLabor(_incidentRecord.ID.ToString());
           
                totallabor = Labor(lResults);
           
                oResults = _rnConnectService.GetOther(_incidentRecord.ID.ToString());
             
                totalother = OtherCharges(oResults);
                totalpartsPrice = PartsPrice(pResults);
                totalpartsCost = PartsCost(pResults);
               // MessageBox.Show("action" + ActionName);
               /// MessageBox.Show("totalpartsCost" + totallabor);

               
                switch (ActionName)
                {
                    case "NetReg":                        
                        net = totalpartsPrice + totallabor + totalother;
                        approvalamt = net;
                        break;
                    case "NetWCD":
                        approvalamt = totalpartsCost + totallabor + totalother;
                        net = totalpartsPrice + totallabor + totalother;
                        break;
                    
                      
                        
                    default:
                        break;
                }

                

                SetIncidentField("c", "Net Amount", Math.Round(net, 0).ToString());
                SetIncidentField("CO", "net_amount_text", Math.Round(net, 2).ToString());

                SetIncidentField("c", "Approval_Amount", Math.Round(approvalamt, 0).ToString());



            }
            catch (Exception e)
            { MessageBox.Show("error in net calculation"); }

        }

        /// <summary>
        /// Method which is called when any Workspace Rule Condition is invoked.
        /// </summary>
        /// <param name="ConditionName">The name of the Workspace Rule Condition that was invoked.</param>
        /// <returns>The result of the condition.</returns>
        public string RuleConditionInvoked(string ConditionName)
        {
            return string.Empty;
        }


        public decimal OtherCharges(string[] rs)
        {
            List<string> oc = new List<string>();
            
            string amt;
            decimal damt;
            decimal total = 0;
            

            if (rs != null && rs.Length > 0)
            {
                oc = rs.ToList();
                total = 0;
               
                foreach (string p in oc)
                {         
                    amt = p.Split('~')[0];

                    
                    if ((amt != null) && (amt != "" )) 
                    {
                        amt = Regex.Replace(amt, "[^.0-9]", "");
                        damt = Convert.ToDecimal(amt);
                        total += damt;
                    }
                }
            }

            return total;
        }



        public decimal Labor(string[] rs)
        {
            List<string> lb = new List<string>();
            string qty;
            string cost;
            decimal dcost;
            decimal total = 0;
            decimal iqty;

            if (rs != null && rs.Length > 0)
            {
                lb = rs.ToList();
                total = 0;
                foreach (string p in lb)
                {
                    qty = p.Split('~')[1];
                    cost = p.Split('~')[0];


                    if ((qty != null) && (cost != null) && (qty != "") && (cost != ""))
                    {
                        qty = Regex.Replace(qty, "[^.0-9]", "");
                        cost = Regex.Replace(cost, "[^.0-9]", "");
                        iqty = Convert.ToDecimal(qty);
                        dcost = Convert.ToDecimal(cost);

                    total += iqty * dcost;
                    }
                }
            }

            return total;
        }



        public decimal PartsCost(string[] rs)
        {
            List<string> parts = new List<string>();
            string qty;
            string cost;
            decimal dcost;
            decimal total = 0;
            decimal iqty;

            if (rs != null && rs.Length > 0)
            {
                parts = rs.ToList();
                total = 0;
                
                foreach (string p in parts)
                {
                    qty = p.Split('~')[2];
                    
                    cost = p.Split('~')[0];
                    if ((qty != null) && (cost != null) && (qty != "") && (cost != ""))
                    {
                        qty = Regex.Replace(qty, "[^.0-9]", "");
                        cost = Regex.Replace(cost, "[^.0-9]", "");
                        iqty = Convert.ToDecimal(qty);
                        dcost = Convert.ToDecimal(cost);

                        total += iqty * dcost;
                    }
                }
            }

            return total; 
        }


        public decimal PartsPrice(string[] rs)
        {
            List<string> parts = new List<string>();
            string qty;
            string pr;
            decimal dpr;
            decimal total = 0;
            decimal iqty;

            if (rs != null && rs.Length > 0)
            {

                parts = rs.ToList();
                total = 0;
                foreach (string p in parts)
                {
                    qty = p.Split('~')[2];
                    pr = p.Split('~')[1];
                    if ((qty != null) && (pr != null) && (qty != "") && (pr != ""))
                    {
                        iqty = Convert.ToDecimal(qty);
                        dpr = Convert.ToDecimal(pr);
                        total += iqty * dpr;
                    } 

                }
            }

            return total;
        }


        public string GetFieldValue(string fieldName, IGenericObject _obj)
        {
            IList<IGenericField> fields = _obj.GenericFields;
            if (null != fields)
            {
                foreach (IGenericField field in fields)
                {
                    if (field.Name.Equals(fieldName))
                    {
                        if (field.DataValue.Value != null)
                            return field.DataValue.Value.ToString();
                    }
                }
            }
            return "";
        }
        /// <summary>
        /// Method which is use to set value to a field using record Context 
        /// </summary>
        /// <param name="fieldName">field name</param>
        /// <param name="value">value of field</param>
        public void SetFieldValue(string fieldName, string value, IGenericObject _obj)
        {
            IList<IGenericField> fields = _obj.GenericFields;
            if (null != fields)
            {
                foreach (IGenericField field in fields)
                {
                    if (field.Name.Equals(fieldName))
                    {
                        switch (field.DataType)
                        {
                            case RightNow.AddIns.Common.DataTypeEnum.STRING:
                                field.DataValue.Value = value;
                                break;
                        }
                    }
                }
            }
            return;
        }

        public static void SetIncidentField(string pkgName, string fieldName, string value)
        {

            if (pkgName == "c")
            {
               
                IList<ICfVal> incCustomFields = _incidentRecord.CustomField;
                int fieldID = GetCustomFieldID(fieldName);
                

                foreach (ICfVal val in incCustomFields)
                {
                    if (val.CfId == fieldID)
                    {
                        

                        switch (val.DataType)
                        {

                            case 1:                    
                                break; 
                            case 3: //Integer//
                                
                                if (value.Trim() == "" || value.Trim() == null)
                                {
                                    val.ValInt = null;
                                }

                                else
                                {
                                    val.ValInt = Convert.ToInt32(value);
                                  
                                }
                                
                                break;        
                        default:
                                break; 

                        }

                    }
                }
            }
            else
            {
               
                IList<ICustomAttribute> incCustomAttributes = _incidentRecord.CustomAttributes;

                foreach (ICustomAttribute val in incCustomAttributes)
                {
                    if (val.PackageName == pkgName)
                    {
                        if (val.GenericField.Name == pkgName + "$" + fieldName)
                        {
                           
                            
                            switch (val.GenericField.DataType)
                            {
                                case RightNow.AddIns.Common.DataTypeEnum.BOOLEAN:
                                    if (value == "1" || value.ToLower() == "true")
                                    {
                                        val.GenericField.DataValue.Value = true;
                                    }
                                    else if (value == "0" || value.ToLower() == "false")
                                    {
                                        val.GenericField.DataValue.Value = false;
                                    }
                                    break;
                                case RightNow.AddIns.Common.DataTypeEnum.INTEGER:
                                    if (value.Trim() == "" || value.Trim() == null)
                                    {
                                        val.GenericField.DataValue.Value = null;
                                    }
                                    else
                                    {
                                        val.GenericField.DataValue.Value = Convert.ToInt32(value);
                                    }
                                    break;
                                case RightNow.AddIns.Common.DataTypeEnum.STRING:
                                   
                                    val.GenericField.DataValue.Value = value;
                                    break;
                            }
                        }
                    }
                }
            }
            return;
        }

        /// </summary>
        /// <param name="fieldName">Custom Field Name</param>
        public static int GetCustomFieldID(string fieldName)
        {
            IList<IOptlistItem> CustomFieldOptList = _globalContext.GetOptlist((int)RightNow.AddIns.Common.OptListID.CustomFields);//92 returns an OptList of custom fields in a hierarchy
            foreach (IOptlistItem CustomField in CustomFieldOptList)
            {
             
                
                if (CustomField.Label == fieldName)//Custom Field Name
                {

                    
                    return (int)CustomField.ID;//Get Custom Field ID
                }
            }
            return -1;
        }

        #endregion
    }

    [AddIn("Workspace Factory AddIn", Version = "1.0.0.0")]
    public class WorkspaceAddInFactory : IWorkspaceComponentFactory2
    {
        #region IWorkspaceComponentFactory2 Members
        static public IGlobalContext _globalContext;

        /// <summary>
        /// Method which is invoked by the AddIn framework when the control is created.
        /// </summary>
        /// <param name="inDesignMode">Flag which indicates if the control is being drawn on the Workspace Designer. (Use this flag to determine if code should perform any logic on the workspace record)</param>
        /// <param name="RecordContext">The current workspace record context.</param>
        /// <returns>The control which implements the IWorkspaceComponent2 interface.</returns>
        public IWorkspaceComponent2 CreateControl(bool inDesignMode, IRecordContext RecordContext)
        {
            return new WorkspaceAddIn(inDesignMode, RecordContext, _globalContext);
        }

        #endregion

        #region IFactoryBase Members

        /// <summary>
        /// The 16x16 pixel icon to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public Image Image16
        {
            get { return Properties.Resources.AddIn16; }
        }

        /// <summary>
        /// The text to represent the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Text
        {
            get { return "Calculate Net Amount"; }
        }

        /// <summary>
        /// The tooltip displayed when hovering over the Add-In in the Ribbon of the Workspace Designer.
        /// </summary>
        public string Tooltip
        {
            get { return "Calculate Net Amount"; }
        }

        #endregion

        #region IAddInBase Members

        /// <summary>
        /// Method which is invoked from the Add-In framework and is used to programmatically control whether to load the Add-In.
        /// </summary>
        /// <param name="GlobalContext">The Global Context for the Add-In framework.</param>
        /// <returns>If true the Add-In to be loaded, if false the Add-In will not be loaded.</returns>
        public bool Initialize(IGlobalContext GlobalContext)
        {
            _globalContext = GlobalContext;
            return true;
        }

        #endregion
    }
}