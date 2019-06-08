using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace CSInterface
{
    [ComVisible(true)]
    [Guid("F0239BF9-0A6E-49A6-8853-BADE1B95E66F")]
    public interface ICSWrapper
    {
        object InvokeMemberMethod(string assemblyName, string methodName,bool hasParas, object[] paras);
        object InvokeStaticMethod(string assemblyName, string methodName,bool hasParas, object[] paras);
        object CreateInstance(string assemblyName, string className,bool hasParas,object[] paras);
        object InvokeObjectMethod(object obj, string methodName, bool hasParas, object[] paras);
        string GetAppDirectory();
    }

}
