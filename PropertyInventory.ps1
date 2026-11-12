# Read COM property categories inside C# to avoid PowerShell COM enumeration quirks.
if (-not ('WordPdfPropertyInventory' -as [type])) {
 Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Reflection;
public static class WordPdfPropertyInventory {
    static object Get(object obj, string name, params object[] args) {
        return obj.GetType().InvokeMember(name, BindingFlags.GetProperty, null, obj, args);
    }
    public static Dictionary<string,object> Read(object document) {
        var result = new Dictionary<string,object>();
        try {
            object builtins = Get(document, "BuiltInDocumentProperties");
            int count = Convert.ToInt32(Get(builtins,"Count"));
            var names = new List<string>();
            for(int i=1;i<=count;i++) {
                try {
                    object prop=Get(builtins,"Item",i);
                    object value=Get(prop,"Value");
                    if(value!=null && !String.IsNullOrWhiteSpace(Convert.ToString(value)))
                        names.Add(Convert.ToString(Get(prop,"Name")));
                } catch { /* Unset built-in properties commonly have no readable value. */ }
            }
            object custom=Get(document,"CustomDocumentProperties");
            result["customPropertiesFound"]=Convert.ToInt32(Get(custom,"Count"));
            names.Sort(StringComparer.OrdinalIgnoreCase);
            result["sourcePropertyTypesFound"]=String.Join("; ",names);
            result["propertyInventoryMethod"]="Word COM property inventory";
            result["propertyInventoryAvailable"]=true;
        } catch(Exception ex) {
            result["propertyInventoryAvailable"]=false;
            result["propertyInventoryError"]=ex.GetType().Name;
        }
        return result;
    }
}
'@
}
