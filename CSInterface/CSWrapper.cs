using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace CSInterface
{
    [ComVisible(true)]
    [Guid("35A5CE1E-551C-41EC-81D4-005318550119")]
    [ProgId("MyLib.MyClassa")]
    public class CSWrapper:ICSWrapper
    {
        public CSWrapper()
        {

        }
        /// <summary>
        /// 加载程序集，assemblyName可以是程序集名称，可以是相对路径、绝对路径(相对于CSInterface的路径)
        /// </summary>
        /// <param name="assemblyName">可以是相对路径、绝对路径</param>
        /// <returns></returns>
        private Assembly LoadAssembly(string assemblyName)
        {
            Assembly asm = null;
            if (assemblyName.Contains(".dll"))
            {
                //说明是路径
                if (!Path.IsPathRooted(assemblyName))
                {
                    assemblyName = Path.Combine(GetAppDirectory(), assemblyName);
                }
                asm = Assembly.LoadFrom(assemblyName);
            }
            else
            {
                asm = Assembly.Load(assemblyName);
            }
            return asm;
        }
        public object InvokeMemberMethod(string assemblyName, string methodName, bool hasParas,object[] paras)
        {
            Assembly asm = LoadAssembly(assemblyName);
            foreach (var type in asm.GetTypes())
            {
                if (!type.IsClass)
                    continue;
                //获取类下面的方法
                var method = type.GetMethod(methodName);
                if (method != null)
                {
                    var obj = Activator.CreateInstance(type);
                    if (!hasParas)
                    {
                        return method.Invoke(obj,null);
                    }
                    else 
                    {
                        return method.Invoke(obj, paras);
                    }
                }
            }
            return null;
        }
        public object InvokeStaticMethod(string assemblyName, string methodName,bool hasParas, object[] paras)
        {
            Assembly asm = LoadAssembly(assemblyName);
            
            foreach (var type in asm.GetTypes())
            {
                if (!type.IsClass)
                    continue;
                //获取类下面的方法
                var method = type.GetMethod(methodName, BindingFlags.InvokeMethod|BindingFlags.Public| BindingFlags.Static);
                if (method != null)
                {
                    if (!hasParas)
                    {
                        return method.Invoke(null,null);
                    }
                    else 
                    {
                        return method.Invoke(null, paras);
                    }
                }
            }
            return null;
        }
        public object CreateInstance(string assemblyName, string className,bool hasParas,object[] paras)
        {
            Assembly asm = LoadAssembly(assemblyName);
            foreach (var type in asm.GetTypes())
            {
                if (!type.IsClass)
                    continue;
                //获取类下面的方法
                if (type.Name == className || type.FullName == className)
                {
                    if (hasParas)
                    {
                        return Activator.CreateInstance(type, paras);
                    }
                    else
                    {
                        return Activator.CreateInstance(type);
                    }
                }
            }
            return null;
        }
        public string GetAppDirectory()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }
    }
}
