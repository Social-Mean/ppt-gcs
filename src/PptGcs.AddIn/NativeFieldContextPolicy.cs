using System;

namespace PptGcs.AddIn
{
    public enum NativeFieldContextState { Matched,Different,Unavailable }

    public sealed class NativeFieldContextUnavailableException:InvalidOperationException
    {
        public NativeFieldContextUnavailableException(string message,Exception inner=null):base(message,inner){}
    }

    public static class NativeFieldContextPolicy
    {
        public static NativeFieldContextState Observe(Func<bool> matches)
        {
            if(matches==null)throw new ArgumentNullException("matches");
            try{return matches()?NativeFieldContextState.Matched:NativeFieldContextState.Different;}
            catch(Exception){return NativeFieldContextState.Unavailable;}
        }
        public static bool CanCommit(NativeFieldContextState state)
        {
            if(state==NativeFieldContextState.Unavailable)
                throw new NativeFieldContextUnavailableException("暂时无法读取 PowerPoint 编辑上下文，本次字段提交已取消。");
            return state==NativeFieldContextState.Matched;
        }
        public static bool RetainCachedField(NativeFieldContextState state)
        {return state!=NativeFieldContextState.Different;}
        public static bool CanVerifyScope(NativeFieldContextState state,bool scopeKnown)
        {return state==NativeFieldContextState.Matched&&scopeKnown;}
        public static T CaptureIdentified<T>(Func<T> capture)
        {
            if(capture==null)throw new ArgumentNullException("capture");
            try{return capture();}
            catch(Exception error){throw new NativeFieldContextUnavailableException("已识别原生字段，但暂时无法取得其编辑上下文，本次输入已取消。",error);}
        }
    }
}
