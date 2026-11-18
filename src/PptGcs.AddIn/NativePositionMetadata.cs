using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Accessibility;

namespace PptGcs.AddIn
{
    // Read-only inspection of the observed Chinese Format Shape / Position UI.
    // No coordinate origin is inferred from an absent or rounded field value.
    internal static class NativePositionMetadata
    {
        public static object Capture(IAccessible field)
        {
            var nodes=Nodes(field);if(nodes==null)return null;
            return new {scope="设置形状格式/位置",nodes,originInterpretation="top-left-supported-center-unverified"};
        }
        public static bool TryReadOrigin(IAccessible field,string fieldName,out string origin)
        {
            origin=null;var nodes=Nodes(field);if(nodes==null)return false;
            bool found=false;
            foreach(var node in nodes)
            {
                if(node.name==fieldName){found=true;continue;}
                if(!found)continue;
                if((node.name=="水平位置"||node.name=="垂直位置")&&node.name!=fieldName)return false;
                if(node.role==46&&node.name=="从"){origin=node.value as string;return true;}
            }
            return false;
        }
        private static List<Node> Nodes(IAccessible field)
        {
            IAccessible group=FindGroup(field,"位置");if(group==null)return null;
            var nodes=new List<Node>();var seen=new HashSet<string>();
            Read(group,0,new int[0],0,nodes,seen);
            return nodes;
        }
        public static bool HasPaneGroup(IAccessible field,string groupName){return FindGroup(field,groupName)!=null;}
        private static IAccessible FindGroup(IAccessible field,string groupName)
        {
            IAccessible group=null,current=field;bool formatPane=false;
            for(int depth=0;current!=null&&depth<9;depth++)
            {
                string name=Name(current,0);
                if(name==groupName&&group==null)group=current;
                if(name=="设置形状格式")formatPane=true;
                try{current=current.accParent as IAccessible;}catch{break;}
            }
            return formatPane?group:null;
        }
        private static void Read(IAccessible accessible,int child,int[] path,int depth,List<Node> nodes,HashSet<string> seen)
        {
            if(nodes.Count>=100||depth>5)return;
            IntPtr unknown=Marshal.GetIUnknownForObject(accessible);string identity;
            try{identity=unknown.ToInt64()+"/"+child;}finally{Marshal.Release(unknown);}
            if(!seen.Add(identity))return;
            int? role=null,state=null;string rawValue=null;
            try{object value=accessible.get_accRole(child);if(value is int)role=(int)value;}catch{}
            try{object value=accessible.get_accState(child);if(value is int)state=(int)value;}catch{}
            try{rawValue=accessible.get_accValue(child);}catch{}
            object safeValue=null;double numeric;
            if(Double.TryParse(rawValue,NumberStyles.Float,CultureInfo.CurrentCulture,out numeric)||
                Double.TryParse(rawValue,NumberStyles.Float,CultureInfo.InvariantCulture,out numeric))
            {if(!Double.IsNaN(numeric)&&!Double.IsInfinity(numeric))safeValue=numeric;}
            // Combo/list choices are native UI settings. Other nonnumeric field
            // values are omitted, so document text cannot enter this trace.
            else if(role==46||role==33||role==34||role==45)safeValue=rawValue;
            nodes.Add(new Node {path=path,name=Name(accessible,child),role=role,state=state,value=safeValue});
            if(child!=0||depth==5||nodes.Count>=100)return;
            int count;try{count=Math.Min(accessible.accChildCount,100-nodes.Count);}catch{return;}
            if(count<=0)return;
            var children=new object[count];int obtained;
            if(AccessibleChildren(accessible,0,count,children,out obtained)<0)return;
            for(int index=0;index<obtained&&nodes.Count<100;index++)
            {
                var childPath=new int[path.Length+1];Array.Copy(path,childPath,path.Length);childPath[path.Length]=index;
                IAccessible element=children[index] as IAccessible;
                if(element!=null)Read(element,0,childPath,depth+1,nodes,seen);
                else if(children[index] is int)
                {
                    int id=(int)children[index];if(id==0)continue;
                    try{element=accessible.get_accChild(id) as IAccessible;}catch{}
                    if(element!=null)Read(element,0,childPath,depth+1,nodes,seen);
                    else Read(accessible,id,childPath,depth+1,nodes,seen);
                }
            }
        }
        private static string Name(IAccessible accessible,int child)
        {try{return accessible.get_accName(child);}catch{return null;}}
        private sealed class Node {public int[] path;public string name;public int? role,state;public object value;}
        [DllImport("oleacc.dll")]
        private static extern int AccessibleChildren([MarshalAs(UnmanagedType.Interface)]IAccessible container,int start,int count,
            [Out,MarshalAs(UnmanagedType.LPArray,ArraySubType=UnmanagedType.Struct,SizeParamIndex=2)]object[] children,out int obtained);
    }
}
