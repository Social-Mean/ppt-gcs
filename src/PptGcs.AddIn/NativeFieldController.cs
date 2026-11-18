using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using Accessibility;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Pre-apply the native size or position field's edit before PowerPoint
    // commits it. The subsequent same-value native commit must remain a no-op;
    // post-change observers cannot merge their writes into that native Undo.
    public sealed class NativeFieldController : IDisposable
    {
        private static readonly Guid AccessibleId=new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");
        private readonly dynamic application;
        private readonly Action<string,string> log;
        private readonly Action normalizeIdentities;
        private readonly uint thread;
        private readonly string widthName,heightName,widthLabel,heightLabel,regionalUnit;
        private readonly WindowHook callback,returnedCallback,fieldMouseCallback;
        private readonly JavaScriptSerializer json=new JavaScriptSerializer {MaxJsonLength=4*1024*1024};
        private IntPtr hook,returnedHook,fieldMouseHook,fieldMouseWindow;
        private Field active,pendingUnitRefresh;
        private MouseClick mouseClick;
        private int pendingUnitKey;
        private readonly HashSet<int> blockedMouseUps=new HashSet<int>();
        private bool applying,inspecting,tracingDeactivation,refreshingMouseHook,fieldMouseBusy,disposed;
        private int previewMouseDiagnostics,actionMouseDiagnostics;
        public bool IsApplying {get{return applying||fieldMouseBusy;}}
        public string GetState()
        {
            RequireThread();Field field=active;NativeFieldContextState context=field==null?NativeFieldContextState.Different:ContextState(field);
            return json.Serialize(new {applying,inspecting,disposed,fieldMouseInstalled=fieldMouseHook!=IntPtr.Zero,
                fieldMouseWindow=fieldMouseWindow.ToInt64(),fieldMouseBusy,field=field==null?null:
                (object)new {name=Name(field),hwnd=field.Handle.ToInt64(),context=context.ToString(),field.Origin,
                    field.Rejected,field.LastCommitSource,field.ScopeKnown,field.ScopeError,field.RootNativeIds,field.RelatedIds,
                    modified=context==NativeFieldContextState.Matched?ModifiedState(field):null}});
        }

        public NativeFieldController(object application,Action<string,string> log,Action normalizeIdentities)
        {
            if(application==null)throw new ArgumentNullException("application");
            if(log==null)throw new ArgumentNullException("log");
            if(normalizeIdentities==null)throw new ArgumentNullException("normalizeIdentities");
            this.application=application;this.log=log;this.normalizeIdentities=normalizeIdentities;thread=GetCurrentThreadId();
            widthName=(string)this.application.CommandBars.GetScreentipMso("ShapeWidth");
            heightName=(string)this.application.CommandBars.GetScreentipMso("ShapeHeight");
            widthLabel=NormalizeLabel((string)this.application.CommandBars.GetLabelMso("ShapeWidth"));
            heightLabel=NormalizeLabel((string)this.application.CommandBars.GetLabelMso("ShapeHeight"));
            if(String.IsNullOrWhiteSpace(widthName)||String.IsNullOrWhiteSpace(heightName)||widthName==heightName)
                throw new InvalidOperationException("PowerPoint did not expose distinct native shape dimension field names.");
            var measurement=new StringBuilder(8);
            regionalUnit=GetLocaleInfoEx(null,0x0000000D,measurement,measurement.Capacity)>0&&measurement.ToString()=="1"?"in":"cm";
            callback=BeforeWindow;hook=SetWindowsHookEx(4,callback,IntPtr.Zero,thread);
            if(hook==IntPtr.Zero)throw new InvalidOperationException("Cannot observe native shape dimension field commits.");
            returnedCallback=AfterWindow;fieldMouseCallback=OnFieldMouse;
            returnedHook=SetWindowsHookEx(12,returnedCallback,IntPtr.Zero,thread);
            if(returnedHook==IntPtr.Zero){UnhookWindowsHookEx(hook);hook=IntPtr.Zero;throw new InvalidOperationException("Cannot observe completed native field focus changes.");}
        }
        public bool HandleKeyboard(int key,bool down)
        {
            RequireThread();if(disposed||applying||fieldMouseBusy||(key!=13&&key!=9))return false;
            try
            {
                if(!down){if(key==pendingUnitKey)RefreshCommittedUnit();return false;}
                pendingUnitRefresh=null;Field field=GetField(GetFocus(),false);if(field==null)return false;
                bool accepted=Commit(field,key==9?"tab":"enter");if(accepted){pendingUnitRefresh=field;pendingUnitKey=key;}return !accepted;
            }
            catch(Exception error){SafeLog("native-field-key-error",new {error=error.ToString()});return error is NativeFieldContextUnavailableException;}
        }
        public bool HandleMouseInput(int kind,int x,int y,IntPtr hwnd)
        {
            RequireThread();if(disposed)return false;
            if(blockedMouseUps.Remove(kind))return true;
            int up=MatchingMouseUp(kind);if(up==0)return false;
            if(applying){blockedMouseUps.Add(up);return true;}
            try
            {
                Field field=FocusedMouseField(GetFocus());
                if(field==null)return false;
                if(hwnd==field.Handle||IsChild(field.Handle,hwnd))return false;
                if(Commit(field,"pointer-before-focus-loss"))return false;
                // Keep the field focused on rejection; its matching release must
                // not reach PowerPoint after we consumed the initiating press.
                blockedMouseUps.Add(up);return true;
            }
            catch(Exception error)
            {
                SafeLog("native-field-pointer-error",new {kind,x,y,error=error.ToString()});
                if(!(error is NativeFieldContextUnavailableException))return false;
                blockedMouseUps.Add(up);return true;
            }
        }
        // Lifecycle only: no geometry polling or idle solve/writeback.
        public void Tick()
        {
            RequireThread();if(disposed||applying||inspecting||fieldMouseBusy||refreshingMouseHook)return;
            if(active!=null&&(!IsWindow(active.Handle)||!NativeFieldContextPolicy.RetainCachedField(ContextState(active))))
            {active=null;if(blockedMouseUps.Count==0)RemoveFieldMouseHook();}
            ReleaseInactiveMouseHook();
        }
        private IntPtr AfterWindow(int code,IntPtr wParam,IntPtr lParam)
        {
            if(code==0&&!disposed&&!applying&&!inspecting&&!refreshingMouseHook)
            {
                var message=(ReturnMessage)Marshal.PtrToStructure(lParam,typeof(ReturnMessage));
                if(message.Message==0x7)
                {
                    refreshingMouseHook=true;
                    try
                    {
                        Field field=FocusedMouseField(message.Window);
                        if(field!=null&&GetFocus()==message.Window)RefreshFieldMouseHook(message.Window);
                    }
                    catch(NativeFieldContextUnavailableException error)
                    {
                        // Identification succeeded; retain mouse ownership even
                        // if its COM context must be retried on the actual down.
                        try{if(GetFocus()==message.Window)RefreshFieldMouseHook(message.Window);}
                        catch(Exception installError){SafeLog("native-field-mouse-focus-error",new {hwnd=message.Window.ToInt64(),error=installError.ToString()});}
                        SafeLog("native-field-mouse-context-unavailable",new {hwnd=message.Window.ToInt64(),message=error.Message});
                    }
                    catch(Exception error){SafeLog("native-field-mouse-focus-error",new {hwnd=message.Window.ToInt64(),error=error.ToString()});}
                    finally{refreshingMouseHook=false;}
                }
            }
            return CallNextHookEx(returnedHook,code,wParam,lParam);
        }
        private void RefreshFieldMouseHook(IntPtr handle)
        {
            RemoveFieldMouseHook();fieldMouseWindow=handle;
            previewMouseDiagnostics=actionMouseDiagnostics=0;
            // Office can install its own hook while setting focus. Install ours
            // after that handler returns, ahead of older thread mouse hooks.
            fieldMouseHook=SetWindowsHookEx(7,fieldMouseCallback,IntPtr.Zero,thread);
            if(fieldMouseHook==IntPtr.Zero){fieldMouseWindow=IntPtr.Zero;throw new InvalidOperationException("Cannot install the focused native field mouse hook.");}
            SafeLog("native-field-mouse-installed",new {hwnd=handle.ToInt64(),thread});
        }
        private IntPtr OnFieldMouse(int code,IntPtr message,IntPtr data)
        {
            IntPtr next=fieldMouseHook;
            return TryHandleFieldMouse(code,message.ToInt32(),data)?new IntPtr(1):CallNextHookEx(next,code,message,data);
        }
        private bool TryHandleFieldMouse(int code,int kind,IntPtr data)
        {
            if(code<0||disposed||fieldMouseWindow==IntPtr.Zero)return false;
            if(code!=0&&code!=3)return false;
            bool? release=MouseRelease(code,kind);if(release.HasValue){ReleaseInactiveMouseHook();return release.Value;}
            int up=MatchingMouseUp(kind);if(up==0)return false;
            var mouse=(MouseMessage)Marshal.PtrToStructure(data,typeof(MouseMessage));
            bool? cached=CachedMouseDown(code,kind,mouse);
            if(cached.HasValue){TraceFieldMouse(code,kind);return cached.Value;}
            if(blockedMouseUps.Contains(up))return true;
            if(fieldMouseBusy||applying){if(mouseClick!=null)mouseClick.Consume=true;blockedMouseUps.Add(up);return true;}
            if(GetFocus()!=fieldMouseWindow)return false;
            var click=new MouseClick {Owner=fieldMouseWindow,Target=mouse.Window,Kind=kind,Up=up,X=mouse.Position.X,Y=mouse.Position.Y,InProgress=true};
            mouseClick=click;
            fieldMouseBusy=true;
            try
            {
                long? modified=ReadModifiedRaw(fieldMouseWindow);
                TraceFieldMouse(code,kind);
                // Both the first PeekMessage and the eventual removal share
                // this result. Never re-solve after native formatting changes
                // the field's text/modified flag during the same physical down.
                bool consumed=modified!=0&&HandleMouseInput(kind,mouse.Position.X,mouse.Position.Y,mouse.Window);
                return CompleteMouseClick(click,consumed);
            }
            catch(Exception error)
            {
                SafeLog("native-field-mouse-error",new {code,kind,owner=fieldMouseWindow.ToInt64(),error=error.ToString()});
                return CompleteMouseClick(click,true);
            }
            finally{click.InProgress=false;fieldMouseBusy=false;ReleaseInactiveMouseHook();}
        }
        private bool? CachedMouseDown(int code,int kind,MouseMessage mouse)
        {
            MouseClick click=mouseClick;if(click==null)return null;
            if(click.PhysicalReleased&&code==3)
            {blockedMouseUps.Remove(click.Up);mouseClick=null;return null;}
            if(click.Owner!=fieldMouseWindow||click.Kind!=kind||click.Target!=mouse.Window||click.X!=mouse.Position.X||click.Y!=mouse.Position.Y)return null;
            if(click.InProgress){click.Consume=true;blockedMouseUps.Add(click.Up);}
            return click.Consume;
        }
        private bool CompleteMouseClick(MouseClick click,bool consumed)
        {
            click.Consume|=consumed;click.InProgress=false;
            if(click.Released){blockedMouseUps.Remove(click.Up);if(Object.ReferenceEquals(mouseClick,click))mouseClick=null;}
            else if(click.Consume)blockedMouseUps.Add(click.Up);
            return click.Consume;
        }
        private bool? MouseRelease(int code,int kind)
        {
            MouseClick click=mouseClick;bool matches=click!=null&&click.Up==kind;
            if(!matches&&!blockedMouseUps.Contains(kind))return null;
            bool consume=blockedMouseUps.Contains(kind)||matches&&(click.Consume||click.InProgress);
            if(matches){click.PhysicalReleased=true;if(click.InProgress)click.Consume=true;}
            if(code==0)
            {blockedMouseUps.Remove(kind);if(matches){click.Released=true;if(!click.InProgress)mouseClick=null;}}
            else if(consume)blockedMouseUps.Add(kind);
            return consume;
        }
        private Field FocusedMouseField(IntPtr focus)
        {
            // Capture the already identified field before any COM getter can
            // pump focus loss. Commit installs its writer guard before checking
            // the document/window/slide identity through COM.
            Field field=active;
            return field!=null&&field.Handle==focus?field:GetField(focus,false);
        }
        private void TraceFieldMouse(int code,int kind)
        {
            if(code==3){if(previewMouseDiagnostics++>=4)return;}else if(actionMouseDiagnostics++>=4)return;
            Field field=active;
            SafeLog("native-field-mouse-observed",new {code,kind,focus=GetFocus().ToInt64(),owner=fieldMouseWindow.ToInt64(),
                active=field==null?0:field.Handle.ToInt64(),modified=ReadModifiedRaw(fieldMouseWindow),applying,readOnly=true});
        }
        private void ReleaseInactiveMouseHook()
        {if(blockedMouseUps.Count==0&&fieldMouseWindow!=IntPtr.Zero&&GetFocus()!=fieldMouseWindow)RemoveFieldMouseHook();}
        private void RemoveFieldMouseHook()
        {if(fieldMouseHook!=IntPtr.Zero)UnhookWindowsHookEx(fieldMouseHook);fieldMouseHook=fieldMouseWindow=IntPtr.Zero;if(blockedMouseUps.Count==0)mouseClick=null;}
        private IntPtr BeforeWindow(int code,IntPtr wParam,IntPtr lParam)
        {
            if(code==0&&!disposed&&!inspecting)
            {
                var message=(CallMessage)Marshal.PtrToStructure(lParam,typeof(CallMessage));
                try
                {
                    if(IsDeactivation(message))TraceDeactivation(message);
                    if(!applying&&message.Message==0x7)GetField(message.Window,true);
                    else if(!applying&&message.Message==0x8)
                    {
                        Field field=active!=null&&active.Handle==message.Window?active:GetField(message.Window,false);
                        if(field!=null)
                        {
                            // PowerPoint can commit before WM_KILLFOCUS reaches
                            // this hook. Never turn this into a follower writer:
                            // that would be a separate native Undo item.
                            VerifyFocusLoss(field);
                            if(field.Rejected)RestoreField(field);
                        }
                        if(pendingUnitRefresh!=null&&pendingUnitRefresh.Handle==message.Window)pendingUnitRefresh=null;
                        if(active!=null&&active.Handle==message.Window)active=null;
                        ReleaseInactiveMouseHook();
                    }
                }
                catch(Exception error){SafeLog("native-field-hook-error",new {message=message.Message,error=error.ToString()});}
            }
            return CallNextHookEx(hook,code,wParam,lParam);
        }
        private Field GetField(IntPtr handle,bool capture)
        {
            if(handle==IntPtr.Zero)return null;
            if(active!=null&&active.Handle==handle)
            {
                NativeFieldContextState context=ContextState(active);
                if(context==NativeFieldContextState.Unavailable||!capture&&context==NativeFieldContextState.Matched)return active;
                if(context==NativeFieldContextState.Different)active=null;
                else if(capture){active.ScopeKnown=false;active.RelatedIds=null;active.RootNativeIds=null;active.ScopeError="正在重新读取字段焦点范围。";}
            }
            if(inspecting)return null;
            inspecting=true;
            try
            {
            var className=new StringBuilder(128);GetClassName(handle,className,className.Capacity);
            string windowClass=className.ToString();
            if(windowClass.IndexOf("RICHEDIT",StringComparison.OrdinalIgnoreCase)<0&&windowClass!="Edit")return null;
            Guid iid=AccessibleId;object raw;
            if(AccessibleObjectFromWindow(handle,0xFFFFFFFC,ref iid,out raw)<0)return null;
            var accessible=raw as IAccessible;if(accessible==null)return null;
            string name=accessible.get_accName(0),origin=null;NativeFieldProperty property;
            if(windowClass=="RICHEDIT60W"&&(name==widthName||name==widthLabel&&IsNativePane(handle)&&NativePositionMetadata.HasPaneGroup(accessible,"大小")))property=NativeFieldProperty.Width;
            else if(windowClass=="RICHEDIT60W"&&(name==heightName||name==heightLabel&&IsNativePane(handle)&&NativePositionMetadata.HasPaneGroup(accessible,"大小")))property=NativeFieldProperty.Height;
            else if(windowClass=="RICHEDIT60W"&&(name=="水平位置"||name=="垂直位置")&&IsNativePane(handle))
            {
                property=name=="水平位置"?NativeFieldProperty.Left:NativeFieldProperty.Top;
                origin=NativeFieldContextPolicy.CaptureIdentified(()=>{string value;NativePositionMetadata.TryReadOrigin(accessible,name,out value);return value;});
            }
            else{LogUnknownField(handle,windowClass,name,accessible);return null;}
            string text=Text(handle);
            Field field=CreateIdentifiedField(handle,property,origin,text);
            field.InitialModifiedRaw=ReadModifiedRaw(handle);
            active=field;
            CaptureFocusScope(field);
            if(ContextState(field)==NativeFieldContextState.Matched)field.RecoveryText=RecoveryText(field);
            SafeLog("native-field-focus",new {field=Name(field),text,unit=field.Unit,origin,capture,modified=ModifiedState(field),
                field.ScopeKnown,field.ScopeError,field.RootNativeIds,field.RelatedIds});
            return field;
            }
            finally{inspecting=false;}
        }
        private Field CreateIdentifiedField(IntPtr handle,NativeFieldProperty property,string origin,string text)
        {
            return NativeFieldContextPolicy.CaptureIdentified(()=>
            {
                dynamic ownerWindow=application.ActiveWindow;
                if(ownerWindow==null)throw new InvalidOperationException("ActiveWindow is unavailable.");
                object presentation=(object)ownerWindow.Presentation;
                if(presentation==null)throw new InvalidOperationException("Window.Presentation is unavailable.");
                return new Field {Handle=handle,Property=property,Origin=origin,InitialText=text,Unit=NativeFieldValue.CaptureUnit(text,regionalUnit),
                    Presentation=presentation,Window=(object)ownerWindow,SlideId=(int)ownerWindow.View.Slide.SlideID};
            });
        }
        private bool Commit(Field field,string source)
        {
            field.Rejected=false;
            applying=true;
            LayoutContext context=null;SelectionMoveFrame before=null;Dictionary<int,object> extras=null;
            try
            {
                if(!NativeFieldContextPolicy.CanCommit(ContextState(field)))return true;
                long? modifiedRaw=ReadModifiedRaw(field.Handle);
                SafeLog("native-field-modify-before",new {source,field=Name(field),modified=ModifiedState(field,modifiedRaw)});
                if(!NativeFieldDirtyPolicy.RequiresPreapply(modifiedRaw))
                {
                    field.LastCommitSource=null;
                    SafeLog("native-field-unmodified",new {source,field=Name(field),modifiedRaw,writes=false});
                    return true;
                }
                dynamic window=field.Window,selection=window.Selection;
                int type=(int)selection.Type;
                if((int)window.ViewType!=9 || (type!=2&&type!=3))return true;
                RequireMatchedContext(field);normalizeIdentities();context=ReadCapturedContext(field);
                if(context.Constraints.Count==0)return true;
                var nativeToId=context.Shapes.ToDictionary(pair=>(int)((dynamic)pair.Value).Id,pair=>pair.Key);
                if((bool)selection.HasChildShapeRange)
                {
                    var children=new List<object>();
                    for(int i=1;i<=(int)selection.ChildShapeRange.Count;i++)
                        children.Add((object)selection.ChildShapeRange[i]);
                    return CommitChildren(context,field,children.ToArray(),source);
                }
                var selected=new List<object>();var driverIds=new List<string>();extras=new Dictionary<int,object>();
                dynamic range=selection.ShapeRange;
                for(int i=1;i<=(int)range.Count;i++)
                {
                    object shape=(object)range[i];selected.Add(shape);int nativeId=(int)((dynamic)shape).Id;string id;
                    if(nativeToId.TryGetValue(nativeId,out id))driverIds.Add(id);else extras.Add(nativeId,shape);
                }
                if(driverIds.Count==0)return true;
                field.RelatedIds=driverIds.ToArray();field.ScopeKnown=true;field.ScopeError=null;
                RefreshOrigin(field);
                field.RecoveryText=RecoveryText(field);
                if(!KnownOrigin(field))throw UnknownOrigin();
                double points=RequestedPoints(field);
                before=new SelectionMoveFrame(PowerPointLayout.Snapshot(context),extras.Values.Select(ReadPosition));
                var byId=before.Geometry.ToDictionary(shape=>shape.Id,StringComparer.Ordinal);var edits=new List<GeometryEdit>();
                foreach(string id in driverIds)
                {
                    ShapeGeometry original=byId[id];dynamic native=context.Shapes[id];bool locked=(int)native.LockAspectRatio!=0;
                    SelectionMovePosition target=NativeFieldTarget.Create(ReadPosition((object)native),field.Property,points,locked);
                    edits.Add(new GeometryEdit(new ShapeGeometry(id,target.Left,target.Top,target.Width,target.Height),
                        !NativeFieldTarget.IsDimension(field.Property)?GeometryEditMode.Move:
                        locked?GeometryEditMode.Exact:field.Property==NativeFieldProperty.Width?GeometryEditMode.NativeWidth:GeometryEditMode.NativeHeight));
                }
                var solved=LayoutSolver.Solve(before.Geometry,context.Constraints,edits);
                var connected=LayoutSolver.ConnectedConstraints(context.Constraints,driverIds);
                var requested=new SelectionMoveFrame(solved,extras.Values.Select(shape=>NativeFieldTarget.Create(ReadPosition(shape),field.Property,points,(int)((dynamic)shape).LockAspectRatio!=0)));
                if(before.Matches(requested)){field.LastCommitSource=source;return true;}
                if(LayoutSolver.MaxResidual(requested.Geometry,connected)>LayoutSolver.PositionTolerance)
                    throw new ConstraintConflictException("原生大小或位置字段不能满足当前约束。");
                SafeLog("native-field-before",new {source,field=Name(field),text=Text(field.Handle),points,modified=ModifiedState(field),
                    selectedNativeIds=selected.Select(shape=>(int)((dynamic)shape).Id).ToArray(),before=before.Geometry,beforeExtras=before.Extras});
                RequireMatchedContext(field);application.StartNewUndoEntry();
                try
                {
                    RequireMatchedContext(field);
                    WriteFrame(context,extras,requested);
                    var actual=ReadFrame(context,extras);
                    if(!actual.Matches(requested)||LayoutSolver.MaxResidual(actual.Geometry,connected)>LayoutSolver.PositionTolerance)
                        throw new InvalidOperationException("PowerPoint 未保留完整的原生大小或位置约束操作。");
                    field.RecoveryText=RecoveryText(field);
                    field.LastCommitSource=source;
                    SafeLog("native-field-committed",new {source,field=Name(field),points,geometry=actual.Geometry,extras=actual.Extras,modified=ModifiedState(field)});
                    AcknowledgePositionCommit(field);
                    return true;
                }
                catch
                {
                    WriteFrame(context,extras,before);
                    if(!ReadFrame(context,extras).Matches(before))throw new InvalidOperationException("无法恢复原生字段操作前的全部对象。");
                    throw;
                }
            }
            catch(Exception error)
            {field.Rejected=true;field.LastCommitSource=null;SafeLog("native-field-rejected",new {source,field=Name(field),text=Text(field.Handle),message=error.Message,error=error.ToString()});return false;}
            finally{applying=false;}
        }
        private void RestoreField(Field field)
        {
            if(ContextState(field)!=NativeFieldContextState.Matched)return;
            applying=true;
            try
            {
                string recovery=field.RecoveryText;
                if(recovery==null)throw new InvalidOperationException("无法取得原生字段对应对象的精确尺寸。");
                bool restored=SetWindowText(field.Handle,recovery);
                SafeLog("native-field-focus-loss-restored",new {field=Name(field),text=recovery,restored,mixed=recovery.Length==0,
                    verification=recovery.Length==0?"mixed-blank-no-commit-requires-native-regression":"requires-native-conflict-focus-loss-regression"});
                if(!restored)throw new InvalidOperationException("无法在失焦提交前恢复原生字段。");
            }
            finally{applying=false;}
        }
        private bool CommitChildren(LayoutContext context,Field field,object[] children,string source)
        {
            var edit=GroupChildNativeEdit.Capture(context,children,GeometryEditMode.Move,false);if(edit==null)return true;
            field.RelatedIds=edit.Plan.GroupIds.ToArray();field.ScopeKnown=true;field.ScopeError=null;RefreshOrigin(field);field.RecoveryText=RecoveryText(field);
            if(!KnownOrigin(field))throw UnknownOrigin();
            double points=RequestedPoints(field);
            var targets=edit.Children.Select(child=>NativeFieldTarget.Create(ReadPosition(child),field.Property,points,(int)((dynamic)child).LockAspectRatio!=0)).ToArray();
            if(GroupChildEditPlan.SameChildren(edit.Plan.Initial.Extras,targets)){field.LastCommitSource=source;return true;}
            edit.ThrowIfProvablyImpossible(targets);
            SafeLog("native-field-child-before",new {source,field=Name(field),points,groupIds=edit.Plan.GroupIds,childrenBefore=edit.Plan.Initial.Extras,before=edit.Plan.Initial.Geometry,modified=ModifiedState(field)});
            RequireMatchedContext(field);application.StartNewUndoEntry();
            // Office's resulting group bounds are authoritative. As in the owned
            // child pointer path, solve external relations after this native
            // child write, within the same atomic rollback and native Undo unit.
            RequireMatchedContext(field);SelectionMoveFrame actual=edit.Apply(targets);
            field.RecoveryText=RecoveryText(field);
            field.LastCommitSource=source;
            SafeLog("native-field-child-committed",new {source,field=Name(field),points,groupIds=edit.Plan.GroupIds,
                children=actual.Extras,geometry=actual.Geometry,
                residualPt=LayoutSolver.MaxResidual(actual.Geometry,LayoutSolver.ConnectedConstraints(context.Constraints,edit.Plan.GroupIds))});
            AcknowledgePositionCommit(field);
            return true;
        }
        private void AcknowledgePositionCommit(Field field)
        {
            if(NativeFieldTarget.IsDimension(field.Property) || GetFocus()!=field.Handle)return;
            // Replace the pending user edit with the verified position before
            // letting PowerPoint process the same Enter, Tab or mouse action.
            double resolved;
            bool restored=NativeFieldValue.TryParsePoints(field.RecoveryText,"pt",CultureInfo.CurrentCulture,Double.NegativeInfinity,out resolved)&&
                SetWindowText(field.Handle,NativeFieldValue.FormatPoints(resolved,field.Unit??regionalUnit,CultureInfo.CurrentCulture));
            UIntPtr result;
            bool delivered=ReadEditMessage(field.Handle,0x00B9,IntPtr.Zero,IntPtr.Zero,0x23,1000,out result)!=IntPtr.Zero;
            SafeLog("native-position-acknowledged",new {field=Name(field),restored,delivered,modified=ReadModifiedRaw(field.Handle),text=Text(field.Handle)});
        }
        private void VerifyFocusLoss(Field field)
        {
            NativeFieldContextState state=ContextState(field);
            if(!NativeFieldContextPolicy.CanVerifyScope(state,field.ScopeKnown))
            {
                SafeLog("native-field-focus-loss-unverified",new {field=Name(field),context=state.ToString(),field.ScopeKnown,field.ScopeError,
                    field.RootNativeIds,field.RelatedIds,residualPt=(double?)null,writes=false});
                return;
            }
            try
            {
                LayoutContext context=ReadCapturedContext(field);
                var scope=LayoutSolver.ConnectedConstraints(context.Constraints,field.RelatedIds);
                double residual=LayoutSolver.MaxResidual(context.Geometry,scope);
                SafeLog(residual>LayoutSolver.PositionTolerance?"native-field-focus-loss-drift":"native-field-focus-loss-observed",
                    new {field=Name(field),preCommitSource=field.LastCommitSource,field.Rejected,scopeKnown=true,residualPt=residual,modified=ModifiedState(field),
                        owned=field.LastCommitSource!=null,geometry=context.Geometry,writes=false});
            }
            catch(Exception error){SafeLog("native-field-focus-loss-read-error",new {error=error.ToString(),residualPt=(double?)null,writes=false});}
        }
        private static int MatchingMouseUp(int kind)
        {
            switch(kind)
            {
                case 0x201:case 0x203:return 0x202;
                case 0x204:case 0x206:return 0x205;
                case 0x207:case 0x209:return 0x208;
                case 0x20B:case 0x20D:return 0x20C;
                case 0xA1:case 0xA3:return 0xA2;
                case 0xA4:case 0xA6:return 0xA5;
                case 0xA7:case 0xA9:return 0xA8;
                case 0xAB:case 0xAD:return 0xAC;
                default:return 0;
            }
        }
        private void RefreshCommittedUnit()
        {
            Field field=pendingUnitRefresh;pendingUnitRefresh=null;
            if(field==null||field!=active||ContextState(field)!=NativeFieldContextState.Matched||!IsWindow(field.Handle))return;
            // Key-UP is after the native Enter/Tab commit. Never treat the typed
            // suffix in the pre-commit buffer as PowerPoint's display unit.
            string displayUnit=NativeFieldValue.CaptureUnit(Text(field.Handle),null);
            if(displayUnit!=null)field.Unit=displayUnit;
            SafeLog("native-field-display-unit",new {field=Name(field),unit=field.Unit});
        }
        private static string RecoveryText(Field field)
        {
            try
            {
                if(!KnownOrigin(field))return null;
                dynamic selection=((dynamic)field.Window).Selection;int type=(int)selection.Type;
                if(type!=2&&type!=3)return null;
                dynamic range=(bool)selection.HasChildShapeRange?selection.ChildShapeRange:selection.ShapeRange;
                if((int)range.Count==0)return null;
                double first=ReadValue((object)range[1],field.Property);
                for(int index=2;index<=(int)range.Count;index++)
                {
                    double value=ReadValue((object)range[index],field.Property);
                    if(value!=first)return String.Empty;
                }
                return (first-OriginOffset(field)).ToString("R",CultureInfo.CurrentCulture)+" pt";
            }
            catch{return null;}
        }
        private void LogUnknownField(IntPtr handle,string windowClass,string name,IAccessible accessible)
        {
            double number;string text=Text(handle);
            object numeric=Double.TryParse(text,NumberStyles.Float,CultureInfo.CurrentCulture,out number)||
                Double.TryParse(text,NumberStyles.Float,CultureInfo.InvariantCulture,out number)?(object)number:null;
            if(numeric!=null&&(Double.IsNaN(number)||Double.IsInfinity(number)))numeric=null;
            var ancestors=new List<object>();IAccessible parent=accessible;
            for(int depth=0;depth<5;depth++)
            {
                try{parent=parent.accParent as IAccessible;if(parent==null)break;ancestors.Add(new {name=parent.get_accName(0)});}
                catch{break;}
            }
            var windowAncestors=new List<string>();
            for(IntPtr current=GetParent(handle);current!=IntPtr.Zero&&windowAncestors.Count<5;current=GetParent(current))
            {var value=new StringBuilder(128);GetClassName(current,value,value.Capacity);windowAncestors.Add(value.ToString());}
            SafeLog("native-field-unrecognized",new {hwnd=handle.ToInt64(),windowClass,name,numeric,ancestors,windowAncestors,
                positionGroup=NativePositionMetadata.Capture(accessible)});
        }
        private static double ReadValue(object shapeObject,NativeFieldProperty property)
        {
            dynamic shape=shapeObject;
            switch(property){case NativeFieldProperty.Width:return (double)shape.Width;case NativeFieldProperty.Height:return (double)shape.Height;
                case NativeFieldProperty.Left:return (double)shape.Left;default:return (double)shape.Top;}
        }
        private static bool KnownOrigin(Field field)
        {return NativeFieldTarget.IsDimension(field.Property)||field.Origin=="左上角"||field.Origin=="居中";}
        private static InvalidOperationException UnknownOrigin(){return new InvalidOperationException("此位置参考原点尚未验证，请先将“从”设为“左上角”。");}
        private static string Name(Field field){return field.Property.ToString().ToLowerInvariant();}
        private static double RequestedPoints(Field field)
        {
            double points;
            if(!NativeFieldValue.TryParsePoints(Text(field.Handle),field.Unit,CultureInfo.CurrentCulture,
                NativeFieldTarget.IsDimension(field.Property)?LayoutSolver.MinimumSize:Double.NegativeInfinity,out points))
                throw new InvalidOperationException("请输入有效数值；可按系统单位输入或使用 cm、mm、in、pt，宽高必须至少为 1 磅。");
            return points+OriginOffset(field);
        }
        private static double OriginOffset(Field field)
        {
            if(NativeFieldTarget.IsDimension(field.Property)||field.Origin=="左上角")return 0;
            if(!KnownOrigin(field))throw UnknownOrigin();
            dynamic document=field.Presentation;
            return NativeFieldTarget.AbsolutePosition(field.Property,NativePositionOrigin.SlideCenter,0,
                (double)document.PageSetup.SlideWidth,(double)document.PageSetup.SlideHeight);
        }
        private static bool IsNativePane(IntPtr handle)
        {
            string[] classes={"NetUICtrlNotifySink","NetUIHWND","NUIPane","MsoWorkPane","MsoCommandBar"};
            IntPtr current=GetParent(handle);
            foreach(string expected in classes)
            {
                var value=new StringBuilder(128);GetClassName(current,value,value.Capacity);
                if(value.ToString()!=expected)return false;current=GetParent(current);
            }
            return true;
        }
        private static void RefreshOrigin(Field field)
        {
            if(NativeFieldTarget.IsDimension(field.Property))return;
            field.Origin=null;Guid iid=AccessibleId;object raw;
            if(!IsNativePane(field.Handle)||AccessibleObjectFromWindow(field.Handle,0xFFFFFFFC,ref iid,out raw)<0)return;
            var accessible=raw as IAccessible;if(accessible==null)return;
            string expected=field.Property==NativeFieldProperty.Left?"水平位置":"垂直位置";
            if(accessible.get_accName(0)!=expected)return;
            string origin;if(NativePositionMetadata.TryReadOrigin(accessible,expected,out origin))field.Origin=origin;
        }
        private static string NormalizeLabel(string label)
        {
            var value=new StringBuilder();
            foreach(char character in label??String.Empty)
                if(Char.GetUnicodeCategory(character)!=UnicodeCategory.Format)value.Append(character);
            return value.ToString().Trim().TrimEnd(':','：').Trim();
        }
        private static object ModifiedState(Field field)
        {return ModifiedState(field,ReadModifiedRaw(field.Handle));}
        private static object ModifiedState(Field field,long? raw)
        {
            string text=Text(field.Handle);
            return new {initialRaw=field.InitialModifiedRaw,currentRaw=raw,
                initialModified=field.InitialModifiedRaw.HasValue?(bool?)(field.InitialModifiedRaw.Value!=0):null,
                currentModified=raw.HasValue?(bool?)(raw.Value!=0):null,
                initialText=field.InitialText,currentText=text,textChanged=!String.Equals(field.InitialText,text,StringComparison.Ordinal),
                readOnlyQuery=true};
        }
        private static long? ReadModifiedRaw(IntPtr handle)
        {
            uint owner;
            if(handle==IntPtr.Zero||!IsWindow(handle)||GetWindowThreadProcessId(handle,out owner)!=GetCurrentThreadId()||owner!=GetCurrentProcessId())return null;
            var name=new StringBuilder(128);GetClassName(handle,name,name.Capacity);if(name.ToString()!="RICHEDIT60W")return null;
            // EM_GETMODIFY distinguishes a retyped value from an untouched field.
            UIntPtr result;
            if(ReadEditMessage(handle,0x00B8,IntPtr.Zero,IntPtr.Zero,0x23,1000,out result)==IntPtr.Zero||!IsWindow(handle))return null;
            return unchecked((long)result.ToUInt64());
        }
        private static bool IsDeactivation(CallMessage message)
        {return (message.Message==0x86&&message.WParam==IntPtr.Zero)||(message.Message==0x6&&(message.WParam.ToInt64()&65535)==0)||
            (message.Message==0x1C&&message.WParam==IntPtr.Zero);}
        private void TraceDeactivation(CallMessage message)
        {
            Field field=active;if(field==null||tracingDeactivation)return;
            tracingDeactivation=true;
            try
            {
                IntPtr root=GetAncestor(field.Handle,2);if(root==IntPtr.Zero||GetAncestor(message.Window,2)!=root)return;
                if(applying)
                {
                    SafeLog("native-field-deactivate-during-apply",new {message=message.Message,wParam=message.WParam.ToInt64(),lParam=message.LParam.ToInt64(),
                        target=WindowInfo(message.Window),root=WindowInfo(root),focus=WindowInfo(GetFocus()),field=Name(field),applying=true,layoutRead=false,writes=false});
                    return;
                }
                NativeFieldContextState state=ContextState(field);
                if(state!=NativeFieldContextState.Matched)
                {
                    SafeLog("native-field-deactivate-context",new {message=message.Message,field=Name(field),context=state.ToString(),layoutRead=false,writes=false});
                    return;
                }
                var selected=new List<object>();var ids=new List<string>();var scope=new List<object>();
                LayoutContext context=ReadCapturedContext(field);dynamic selection=((dynamic)field.Window).Selection;
                if((int)selection.Type==2||(int)selection.Type==3)
                {
                    dynamic range=(bool)selection.HasChildShapeRange?selection.ChildShapeRange:selection.ShapeRange;
                    var byNative=context.Shapes.ToDictionary(pair=>(int)((dynamic)pair.Value).Id,pair=>pair.Key);
                    for(int index=1;index<=(int)range.Count;index++)
                    {
                        object shape=(object)range[index];object owner=ShapeCapabilities.TopLevel(shape);int nativeRoot=(int)((dynamic)owner).Id;string id;
                        if(byNative.TryGetValue(nativeRoot,out id))ids.Add(id);
                        selected.Add(new {geometry=ReadPosition(shape),rootNativeId=nativeRoot});
                    }
                }
                foreach(var relation in LayoutSolver.ConnectedConstraints(context.Constraints,ids))scope.Add(new {relation.Id,kind=relation.Kind.ToString()});
                SafeLog("native-field-deactivate-before",new {message=message.Message,wParam=message.WParam.ToInt64(),lParam=message.LParam.ToInt64(),
                    target=WindowInfo(message.Window),root=WindowInfo(root),focus=WindowInfo(GetFocus()),field=Name(field),
                    modified=ModifiedState(field),applying,selected,scope,writes=false});
            }
            catch(Exception error){SafeLog("native-field-deactivate-read-error",new {message=message.Message,error=error.ToString(),writes=false});}
            finally{tracingDeactivation=false;}
        }
        private static object WindowInfo(IntPtr handle)
        {
            var name=new StringBuilder(128);GetClassName(handle,name,name.Capacity);uint process;uint owner=GetWindowThreadProcessId(handle,out process);
            return new {hwnd=handle.ToInt64(),className=name.ToString(),thread=owner,process};
        }
        private static SelectionMovePosition ReadPosition(object shapeObject)
        {dynamic shape=shapeObject;return new SelectionMovePosition((int)shape.Id,(double)shape.Left,(double)shape.Top,(double)shape.Width,(double)shape.Height);}
        private static SelectionMoveFrame ReadFrame(LayoutContext context,Dictionary<int,object> extras)
        {return new SelectionMoveFrame(PowerPointLayout.Snapshot(context),extras.Values.Select(ReadPosition));}
        private static void WriteFrame(LayoutContext context,Dictionary<int,object> extras,SelectionMoveFrame frame)
        {
            PowerPointLayout.Apply(context,frame.Geometry);
            foreach(var value in frame.Extras)
                WritePosition(extras[value.NativeId],value);
        }
        private static void WritePosition(object shapeObject,SelectionMovePosition value)
        {
            dynamic shape=shapeObject;int aspect=(int)shape.LockAspectRatio;
            if((int)shape.Id!=value.NativeId)throw new InvalidOperationException("原生字段编辑的对象标识发生变化。");
            bool resized=Math.Abs((double)shape.Width-value.Width)>.0001||Math.Abs((double)shape.Height-value.Height)>.0001;
            try
            {
                if(resized&&aspect!=0)shape.LockAspectRatio=0;
                if(Math.Abs((double)shape.Width-value.Width)>.0001)shape.Width=(float)value.Width;
                if(Math.Abs((double)shape.Height-value.Height)>.0001)shape.Height=(float)value.Height;
                if(Math.Abs((double)shape.Left-value.Left)>.0001)shape.Left=(float)value.Left;
                if(Math.Abs((double)shape.Top-value.Top)>.0001)shape.Top=(float)value.Top;
            }
            finally{if(resized&&aspect!=0)shape.LockAspectRatio=aspect;}
        }
        private NativeFieldContextState ContextState(Field field)
        {
            return NativeFieldContextPolicy.Observe(()=>
            {
                if((int)application.Presentations.Count==0)return false;
                object presentation=(object)application.ActivePresentation;
                if(presentation==null)throw new InvalidOperationException("ActivePresentation is unavailable.");
                if(!Same(presentation,field.Presentation))return false;
                dynamic window=application.ActiveWindow;
                if(window==null)throw new InvalidOperationException("ActiveWindow is unavailable.");
                return Same((object)window,field.Window)&&(int)window.ViewType==9&&(int)window.View.Slide.SlideID==field.SlideId;
            });
        }
        private void RequireMatchedContext(Field field)
        {if(!NativeFieldContextPolicy.CanCommit(ContextState(field)))throw new InvalidOperationException("字段所属的演示文稿、窗口或幻灯片已改变。");}
        private LayoutContext ReadCapturedContext(Field field)
        {
            RequireMatchedContext(field);LayoutContext context=PowerPointLayout.ReadWindow(field.Window);
            if(!Same(context.Presentation,field.Presentation)||!Same(context.Window,field.Window)||context.SlideId!=field.SlideId)
                throw new InvalidOperationException("读取字段上下文时文稿、窗口或幻灯片发生改变。");
            RequireMatchedContext(field);return context;
        }
        private void CaptureFocusScope(Field field)
        {
            field.ScopeKnown=false;field.RelatedIds=null;field.RootNativeIds=null;field.ScopeError=null;
            try
            {
                LayoutContext context=ReadCapturedContext(field);dynamic selection=((dynamic)field.Window).Selection;
                if((int)selection.Type!=2&&(int)selection.Type!=3)throw new InvalidOperationException("无法确定字段对应的所选对象。");
                dynamic range=(bool)selection.HasChildShapeRange?selection.ChildShapeRange:selection.ShapeRange;
                if((int)range.Count==0)throw new InvalidOperationException("字段没有可确定的所选对象。");
                var byNative=context.Shapes.ToDictionary(pair=>(int)((dynamic)pair.Value).Id,pair=>pair.Key);
                var roots=new HashSet<int>();var ids=new HashSet<string>(StringComparer.Ordinal);
                for(int index=1;index<=(int)range.Count;index++)
                {
                    int root=(int)((dynamic)ShapeCapabilities.TopLevel((object)range[index])).Id;roots.Add(root);string id;
                    if(byNative.TryGetValue(root,out id))ids.Add(id);
                }
                RequireMatchedContext(field);field.RootNativeIds=roots.ToArray();field.RelatedIds=ids.ToArray();field.ScopeKnown=true;
            }
            catch(Exception error){field.ScopeError=error.Message;}
        }
        private static bool Same(object first,object second)
        {IntPtr a=Marshal.GetIUnknownForObject(first);try{IntPtr b=Marshal.GetIUnknownForObject(second);try{return a==b;}finally{Marshal.Release(b);}}finally{Marshal.Release(a);}}
        private static string Text(IntPtr handle){var value=new StringBuilder(512);GetWindowText(handle,value,value.Capacity);return value.ToString();}
        private void SafeLog(string kind,object value){try{log(kind,json.Serialize(value));}catch{}}
        private void RequireThread(){if(GetCurrentThreadId()!=thread)throw new InvalidOperationException("Native field edits must run on PowerPoint's UI thread.");}
        public void Dispose(){RequireThread();if(disposed)return;disposed=true;RemoveFieldMouseHook();if(returnedHook!=IntPtr.Zero)UnhookWindowsHookEx(returnedHook);returnedHook=IntPtr.Zero;if(hook!=IntPtr.Zero)UnhookWindowsHookEx(hook);hook=IntPtr.Zero;active=null;pendingUnitRefresh=null;blockedMouseUps.Clear();}
        private sealed class Field {public IntPtr Handle;public NativeFieldProperty Property;public bool Rejected,ScopeKnown;public int SlideId;public long? InitialModifiedRaw;public string Unit,Origin,InitialText,RecoveryText,LastCommitSource,ScopeError;public string[] RelatedIds;public int[] RootNativeIds;public object Presentation,Window;}
        private sealed class MouseClick {public IntPtr Owner,Target;public int Kind,Up,X,Y;public bool InProgress,Consume,PhysicalReleased,Released;}
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]private delegate IntPtr WindowHook(int code,IntPtr wParam,IntPtr lParam);
        [StructLayout(LayoutKind.Sequential)]private struct CallMessage {public IntPtr LParam,WParam;public uint Message;public IntPtr Window;}
        [StructLayout(LayoutKind.Sequential)]private struct ReturnMessage {public IntPtr Result,LParam,WParam;public uint Message;public IntPtr Window;}
        [StructLayout(LayoutKind.Sequential)]private struct Point {public int X,Y;}
        [StructLayout(LayoutKind.Sequential)]private struct MouseMessage {public Point Position;public IntPtr Window;public uint HitTest;public UIntPtr ExtraInfo;}
        [DllImport("kernel32.dll")]private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll")]private static extern uint GetCurrentProcessId();
        [DllImport("kernel32.dll",CharSet=CharSet.Unicode)]private static extern int GetLocaleInfoEx(string locale,uint type,StringBuilder data,int size);
        [DllImport("user32.dll")]private static extern IntPtr GetFocus();
        [DllImport("user32.dll")]private static extern IntPtr GetParent(IntPtr handle);
        [DllImport("user32.dll")]private static extern IntPtr GetAncestor(IntPtr handle,uint flags);
        [DllImport("user32.dll")]private static extern bool IsWindow(IntPtr handle);
        [DllImport("user32.dll")]private static extern bool IsChild(IntPtr parent,IntPtr child);
        [DllImport("user32.dll")]private static extern uint GetWindowThreadProcessId(IntPtr handle,out uint process);
        [DllImport("user32.dll",EntryPoint="SendMessageTimeoutW",CharSet=CharSet.Unicode,SetLastError=true)]private static extern IntPtr ReadEditMessage(IntPtr handle,uint message,IntPtr wParam,IntPtr lParam,uint flags,uint timeout,out UIntPtr result);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern int GetClassName(IntPtr hwnd,StringBuilder value,int size);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern int GetWindowText(IntPtr hwnd,StringBuilder value,int size);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern bool SetWindowText(IntPtr hwnd,string value);
        [DllImport("user32.dll",SetLastError=true)]private static extern IntPtr SetWindowsHookEx(int type,WindowHook callback,IntPtr module,uint thread);
        [DllImport("user32.dll")]private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")]private static extern IntPtr CallNextHookEx(IntPtr hook,int code,IntPtr wParam,IntPtr lParam);
        [DllImport("oleacc.dll")]private static extern int AccessibleObjectFromWindow(IntPtr hwnd,uint objectId,ref Guid iid,[MarshalAs(UnmanagedType.Interface)]out object result);
    }
}
