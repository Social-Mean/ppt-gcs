using System;

namespace PptGcs.AddIn
{
    public static class NativeFieldDirtyPolicy
    {
        // PowerPoint's RichEdit flag distinguishes an untouched rounded display
        // from deliberately retyping the same display value. Text equality does
        // not identify the user's edit and must not participate in this decision.
        public static bool RequiresPreapply(long? modifiedRaw)
        {
            if(!modifiedRaw.HasValue)throw new InvalidOperationException("无法读取原生字段的修改状态，操作已取消。");
            return modifiedRaw.Value!=0;
        }
    }
}
