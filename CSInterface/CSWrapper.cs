using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace CSInterface
{
    [ComVisible(true)]
    [Guid("35A5CE1E-551C-41EC-81D4-005318550119")]
    public class CSWrapper : ICSWrapper
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
            if (assemblyName.ToLower().Contains(".dll") || assemblyName.ToLower().Contains(".exe"))
            {
                //说明是路径
                if (!Path.IsPathRooted(assemblyName))
                {
                    //在CSInterface的路径下寻找dll
                    assemblyName = Path.Combine(GetAppDirectory(), assemblyName);
                    if (!File.Exists(assemblyName))
                    {
                        //在CSInterface的目录下寻找dll，不存在则到程序运行的目录下寻找
                        assemblyName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, assemblyName);
                    }
                }
                SetWorkDir(Path.GetDirectoryName(assemblyName));
                asm = Assembly.LoadFrom(assemblyName);
            }
            else
            {
                SetWorkDir(Path.GetDirectoryName(assemblyName));
                asm = Assembly.Load(assemblyName);
            }
            return asm;
        }
        private void SetWorkDir(string dir)
        {
            Process.GetCurrentProcess().StartInfo.WorkingDirectory = dir;
        }
        public object InvokeMemberMethod(string assemblyName, string methodName, bool hasParas, object[] paras)
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
                        return method.Invoke(obj, null);
                    }
                    else
                    {
                        return method.Invoke(obj, paras);
                    }
                }
            }
            return null;
        }
        public object InvokeStaticMethod(string assemblyName, string methodName, bool hasParas, object[] paras)
        {
            Assembly asm = LoadAssembly(assemblyName);

            foreach (var type in asm.GetTypes())
            {
                if (!type.IsClass)
                    continue;
                //获取类下面的方法
                var method = type.GetMethod(methodName, BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Static);
                if (method != null)
                {
                    if (!hasParas)
                    {
                        return method.Invoke(null, null);
                    }
                    else
                    {
                        return method.Invoke(null, paras);
                    }
                }
            }
            return null;
        }
        public object CreateInstance(string assemblyName, string className, bool hasParas, object[] paras)
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
        public object InvokeObjectMethod(object obj, string methodName, bool hasParas, object[] paras)
        {
            MethodInfo method = null;
            if (hasParas)
            {
                Type[] types = new Type[paras.Length];
                for (int i = 0; i < paras.Length; i++)
                {
                    types[i] = paras[i].GetType();
                }
                method = obj.GetType().GetMethod(methodName, types);
                return method.Invoke(obj, paras);
            }
            else
            {
                method = obj.GetType().GetMethod(methodName, new Type[] { });
                return method.Invoke(obj, null);
            }
        }
        public string GetAppDirectory()
        {
            return Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }
        public string TestString(string str)
        {
            return str;
        }
    }
}
