using System;  
using System.Web.UI;  
  
public partial class index : Page  
{  
    protected void Page_Load(object sender, EventArgs e)  
    {  
        if (IsPostBack || Request.HttpMethod == "POST")  
        {  
            string action = Request.QueryString["action"];  
            if (!string.IsNullOrEmpty(action))  
            {  
                ManageComPlusObject(action);  
            }  
        }  
    }  
  
    private void ManageComPlusObject(string action)  
    {  
        try  
        {  
            // COM+ object name  
            var comPlusObject = "LebbermanREDCom";  
            var comPlus = new Activator.CreateInstance(Type.GetTypeFromProgID(comPlusObject));  
  
            if (action.Equals("stop", StringComparison.OrdinalIgnoreCase))  
            {  
                // Assuming Stop method exists  
                comPlus.GetType().InvokeMember("Stop", System.Reflection.BindingFlags.InvokeMethod, null, comPlus, null);  
                Response.Write("LebbermanREDCom stopped successfully.");  
            }  
            else if (action.Equals("restart", StringComparison.OrdinalIgnoreCase))  
            {  
                // Assuming Stop and Start methods exist  
                comPlus.GetType().InvokeMember("Stop", System.Reflection.BindingFlags.InvokeMethod, null, comPlus, null);  
                comPlus.GetType().InvokeMember("Start", System.Reflection.BindingFlags.InvokeMethod, null, comPlus, null);  
                Response.Write("LebbermanREDCom restarted successfully.");  
            }  
            else  
            {  
                Response.Write("Invalid action.");  
            }  
        }  
        catch (Exception e)  
        {  
            Response.Write("Error: " + e.Message);  
        }  
    }  
}  
