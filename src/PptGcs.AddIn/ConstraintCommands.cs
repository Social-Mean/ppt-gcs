using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Ribbon, task pane and radial menu all use these atomic document commands.
    public sealed class ConstraintCommands
    {
        private readonly dynamic application;
        private readonly Action<object> normalizeIdentities;
        private readonly Action<string,string> log;
        private string freedomKey;
        private Dictionary<string,GeometryFreedom> freedomByShape;
        public ConstraintCommands(object application, Action<object> normalizeIdentities, Action<string,string> log)
        { this.application=application; this.normalizeIdentities=normalizeIdentities; this.log=log; }

        public string Add(string kind, double? value, double? second, string containerId=null, string parameterId=null)
        {
            NormalizeCurrent();
            EditPlan plan=PlanAdd(ParseKind(kind),value,second,containerId,parameterId);
            Commit(plan);
            return plan.RelationId;
        }

        public IList<ShapeGeometry> Preview(string kind)
        {
            var option=BuildModel().Options.Single(o=>o.Kind==kind);
            var plan=PlanAdd(ParseKind(kind),option.DefaultValue,option.DefaultSecond);
            var before=plan.Context.Geometry.ToDictionary(s=>s.Id);
            return plan.Solved.Where(s=>LayoutSolver.NeedsUpdate(before[s.Id],s)).ToList();
        }

        public void Update(string id, double? value, double? second, string parameterId=null)
        {
            Change(id,(document,relation)=>{
                var parameter=ParameterCommands.Find(document,parameterId);
                return new LayoutConstraint(relation.Id,relation.SlideId,relation.Kind,relation.ShapeIds,
                    parameter?.Value??value,second,relation.Enabled,parameter?.Id);
            });
        }
        public void SetEnabled(string id,bool enabled)
        { Change(id,(document,relation)=>relation.WithEnabled(enabled)); }
        public void Delete(string id) { Change(id,(document,relation)=>null); }

        public void DeleteForSelection()
        {
            NormalizeCurrent();
            LayoutContext context=ReadCurrent();
            var ids=new HashSet<string>(SelectedShapes(context.Window).Select(PowerPointLayout.ReadId).Where(id=>id!=null));
            if(ids.Count==0) throw new InvalidOperationException("请选择有约束的对象。");
            var next=context.Document.WithConstraints(context.Document.Constraints.Where(r=>r.SlideId!=context.SlideId || !r.ShapeIds.Any(ids.Contains)));
            if(next.ToXml()==context.Document.ToXml()) return;
            Commit(SolvePlan(new EditPlan {Context=context,Document=next}));
        }

        private void Change(string id,Func<ConstraintDocument,LayoutConstraint,LayoutConstraint> change)
        {
            NormalizeCurrent();
            LayoutContext context=ReadCurrent();
            string normalized=ShapeIdentity.Normalize(id);
            LayoutConstraint current=context.Document.Constraints.SingleOrDefault(r=>r.Id==normalized && r.SlideId==context.SlideId);
            if(current==null) throw new InvalidOperationException("该约束已不存在或不属于当前页。");
            LayoutConstraint replacement=change(context.Document,current);
            var next=context.Document.WithConstraints(context.Document.Constraints.Select(r=>r.Id==current.Id?replacement:r).Where(r=>r!=null));
            if(next.ToXml()==context.Document.ToXml()) return;
            Commit(SolvePlan(new EditPlan {Context=context,Document=next,Attempted=replacement}));
        }

        private EditPlan PlanAdd(ConstraintKind kind,double? value,double? second,string containerId=null,string parameterId=null)
        {
            LayoutContext context=ReadCurrent();
            if((bool)((dynamic)context.Window).Selection.HasChildShapeRange)
                throw new InvalidOperationException("请先退出组合内部编辑，并选中组合整体建立约束。");
            return PlanAdd(context,SelectedShapes(context.Window),kind,value,second,containerId,parameterId);
        }

        private static EditPlan PlanAdd(LayoutContext context,List<object> selected,ConstraintKind kind,double? value,
            double? second,string containerId=null,string parameterId=null)
        {
            int minimum=LayoutConstraint.GetMinimumShapeCount(kind), maximum=LayoutConstraint.GetMaximumShapeCount(kind);
            if(selected.Count<minimum || (maximum>0 && selected.Count>maximum))
                throw new InvalidOperationException(maximum==minimum ? "此约束需要选中 "+minimum+" 个对象。" : "请选中至少 "+minimum+" 个对象。");
            foreach(object shape in selected) PowerPointLayout.ValidateShape(shape);
            if(RelativeAnchor.UsesContainer(kind))
            {
                object frame=selected.SingleOrDefault(shape=>UiKey(context.SlideId,(int)((dynamic)shape).Id)==containerId);
                if(frame==null)throw new InvalidOperationException("请在已选对象中指定参照容器。");
                selected.Remove(frame);selected.Insert(0,frame);
            }
            if(ConstraintCatalog.IsHorizontalGap(kind)) selected=selected.OrderBy(s=>(double)((dynamic)s).Left).ToList();
            if(ConstraintCatalog.IsVerticalGap(kind)) selected=selected.OrderBy(s=>(double)((dynamic)s).Top).ToList();
            var tags=new List<NewIdentity>();
            var ids=new List<string>();
            foreach(object shape in selected)
            {
                string id=PowerPointLayout.ReadId(shape);
                if(id==null)
                {
                    id=Guid.NewGuid().ToString("D").ToUpperInvariant();
                    tags.Add(new NewIdentity { Shape=shape, Id=id, Before=(string)((dynamic)shape).Tags[PowerPointLayout.TagName],
                        OwnerBefore=(string)((dynamic)shape).Tags[ShapeIdentityManager.OwnerTagName] });
                    context.Shapes.Add(id,shape);
                    context.Geometry.Add(PowerPointLayout.Geometry(shape,id));
                }
                ids.Add(id);
            }
            string members=LayoutConstraint.MembershipKey(kind,ids);
            if(context.Document.Constraints.Any(r=>r.SlideId==context.SlideId && r.Kind==kind &&
                LayoutConstraint.MembershipKey(kind,r.ShapeIds)==members))
                throw new InvalidOperationException("这些对象已经存在“"+ConstraintCatalog.Label(kind)+"”关系，请在窗格中修改或启用它。");
            var parameter=ParameterCommands.Find(context.Document,parameterId);
            var relation=new LayoutConstraint(Guid.NewGuid().ToString("D"),context.SlideId,kind,ids,parameter?.Value??value,second,parameterId:parameter?.Id);
            var next=context.Document.WithConstraints(context.Document.Constraints.Concat(new[]{relation}));
            return SolvePlan(new EditPlan {Context=context,Document=next,Attempted=relation,NewIdentities=tags,RelationId=relation.Id});
        }

        internal IReadOnlyList<CanvasConstraintOffer> FindSuggestions(IReadOnlyDictionary<int,ShapeGeometry> baseline)
        {
            var context=ReadCurrent();
            if(((dynamic)context.Window).Selection.HasChildShapeRange)return Array.Empty<CanvasConstraintOffer>();
            var shapes=SuggestionShapes(context);
            var edited=shapes.Where(item=>baseline.ContainsKey(item.Key) && baseline[item.Key].Id==item.Value.Id &&
                LayoutSolver.NeedsUpdate(baseline[item.Key],item.Value)).Select(item=>item.Value.Id).ToList();
            if(edited.Count==0)return Array.Empty<CanvasConstraintOffer>();
            var byId=shapes.ToDictionary(item=>item.Value.Id,item=>item.Key,StringComparer.Ordinal);
            var current=context.Document.Constraints.Where(item=>item.SlideId==context.SlideId)
                .Select(item=>item.BindPageSize(context.PageWidth,context.PageHeight)).ToList();
            return ConstraintSuggestions.Find(context.SlideId,shapes.Values.ToList(),edited,current).Select(offer=>
                new CanvasConstraintOffer(offer.Kind,byId[offer.AnchorId],offer.MemberIds.Select(id=>byId[id]).ToArray(),
                    offer.MemberIds.Select(id=>shapes[byId[id]]).ToArray())).ToList().AsReadOnly();
        }

        internal string AcceptSuggestion(CanvasConstraintOffer offer)
        {
            var context=ReadCurrent();
            dynamic slide=context.Slide;
            var byNativeId=new Dictionary<int,object>();
            for(int index=1;index<=(int)slide.Shapes.Count;index++)
            {
                object shape=(object)slide.Shapes[index];
                int nativeId=(int)((dynamic)shape).Id;
                if(offer.MemberIds.Contains(nativeId) && IsSupported(shape))byNativeId.Add(nativeId,shape);
            }
            if(byNativeId.Count!=offer.MemberIds.Count)return null;
            var selected=offer.MemberIds.Select(id=>byNativeId[id]).ToList();
            var geometry=offer.MemberIds.Select(id=>PowerPointLayout.Geometry(byNativeId[id],
                PowerPointLayout.ReadId(byNativeId[id])??SuggestionIdentity(context.SlideId,id))).ToList();
            if(geometry.Where((shape,index)=>shape.Id!=offer.Geometry[index].Id || LayoutSolver.NeedsUpdate(shape,offer.Geometry[index])).Any() ||
                ConstraintSuggestions.Error(offer.Kind,geometry)>ConstraintSuggestions.Tolerance || ConstraintSuggestions.AlreadyRelated(
                    context.Document.Constraints.Where(item=>item.SlideId==context.SlideId),offer.Kind,geometry.Select(shape=>shape.Id).ToArray()))return null;
            EditPlan plan;
            try {plan=PlanAdd(context,selected,offer.Kind,null,null);}
            catch(DraftConflict) {return null;}
            Commit(plan);
            return plan.RelationId;
        }

        internal static string SuggestionIdentity(int slideId,int shapeId)
        {
            // Preview-only identities cannot collide with the UUIDv4 tags
            // generated when the user actually creates a relationship.
            return new Guid(slideId,0,0x7000,BitConverter.GetBytes((long)shapeId)).ToString("D").ToUpperInvariant();
        }

        private static Dictionary<int,ShapeGeometry> SuggestionShapes(LayoutContext context)
        {
            var result=new Dictionary<int,ShapeGeometry>();
            var byId=context.Geometry.ToDictionary(item=>item.Id,StringComparer.Ordinal);
            dynamic slide=context.Slide;
            for(int index=1;index<=(int)slide.Shapes.Count;index++)
            {
                object shape=(object)slide.Shapes[index];
                int nativeId=(int)((dynamic)shape).Id;
                string id=PowerPointLayout.ReadId(shape);
                ShapeGeometry geometry;
                if(id!=null && byId.TryGetValue(id,out geometry))result.Add(nativeId,geometry);
                else if(id==null && IsSupported(shape))result.Add(nativeId,PowerPointLayout.Geometry(shape,SuggestionIdentity(context.SlideId,nativeId)));
            }
            return result;
        }

        private static EditPlan SolvePlan(EditPlan plan)
        {
            var context=plan.Context;
            var active=PowerPointLayout.BindRelations(context,plan.Document);
            var missing=active.SelectMany(r=>r.ShapeIds).Where(id=>!context.Shapes.ContainsKey(id)).Distinct().ToList();
            if(missing.Count>0) throw new InvalidOperationException("存在引用缺失或不支持对象的约束，请先在约束窗格禁用或删除该关系。");
            try {plan.Solved=LayoutSolver.Normalize(context.Geometry,active);return plan;}
            catch(ConstraintConflictException error)
            {
                var witness=LayoutDiagnostics.ExplainConflict(context.Geometry,active,plan.Attempted);
                var existing=new HashSet<string>(context.Document.Constraints.Where(item=>item.Enabled).Select(item=>item.Id),StringComparer.Ordinal);
                var choices=witness.Select(item=>new ConflictRelation(item.Id,DescribeRelation(context,item),existing.Contains(item.Id) && item.Id!=plan.Attempted?.Id)).ToList();
                string attempted=plan.Attempted==null ? null : DescribeRelation(context,plan.Attempted);
                throw new DraftConflict(plan,attempted,choices,witness,error);
            }
        }

        public void ResolveConflict(DraftConflict conflict,IReadOnlyList<string> disableIds)
        {
            if(conflict==null)throw new ArgumentNullException("conflict");
            var plan=conflict.Plan;
            var allowed=new HashSet<string>(conflict.Choices.Where(choice=>choice.CanDisable).Select(choice=>choice.Id),StringComparer.Ordinal);
            if(disableIds==null || disableIds.Count==0 || disableIds.Any(id=>!allowed.Contains(id)))
                throw new InvalidOperationException("请选择当前冲突中可以停用的关系。");
            var current=ReadCurrent();
            if(!SameIdentity(current.Window,plan.Context.Window) || !SameIdentity(current.Presentation,plan.Context.Presentation) ||
                current.SlideId!=plan.Context.SlideId || current.Document.ToXml()!=plan.Context.Document.ToXml() ||
                current.PageWidth!=plan.Context.PageWidth || current.PageHeight!=plan.Context.PageHeight)
                throw new InvalidOperationException("演示文稿中的约束已改变，请重新提交编辑。");
            // Retain the intended targets even if selection changes. Re-read
            // the slide so the retry minimizes movement from current geometry.
            dynamic slide=current.Slide;
            foreach(var tag in plan.NewIdentities)
            {
                int nativeId=(int)((dynamic)tag.Shape).Id;
                object shape=null;
                for(int i=1;i<=(int)slide.Shapes.Count;i++)
                    if((int)slide.Shapes[i].Id==nativeId){shape=(object)slide.Shapes[i];break;}
                if(shape==null || !SameIdentity(shape,tag.Shape) || (string)((dynamic)shape).Tags[PowerPointLayout.TagName]!=tag.Before ||
                    (string)((dynamic)shape).Tags[ShapeIdentityManager.OwnerTagName]!=tag.OwnerBefore)
                    throw new InvalidOperationException("参与对象已经改变，请重新建立约束。");
                PowerPointLayout.ValidateShape(shape);
                current.Shapes.Add(tag.Id,shape);
                current.Geometry.Add(PowerPointLayout.Geometry(shape,tag.Id));
            }
            var disabled=new HashSet<string>(disableIds,StringComparer.Ordinal);
            var next=plan.Document.WithConstraints(plan.Document.Constraints.Select(item=>disabled.Contains(item.Id)
                ? item.WithEnabled(false) : item));
            Commit(SolvePlan(new EditPlan {Context=current,Document=next,Attempted=plan.Attempted,
                NewIdentities=plan.NewIdentities,RelationId=plan.RelationId}));
        }

        private static bool SameIdentity(object first,object second)
        {
            IntPtr left=Marshal.GetIUnknownForObject(first);
            try{IntPtr right=Marshal.GetIUnknownForObject(second);try{return left==right;}finally{Marshal.Release(right);}}
            finally{Marshal.Release(left);}
        }

        internal static string DescribeRelation(LayoutContext context,LayoutConstraint relation)
        {
            string name=ConstraintCatalog.Label(relation.Kind);
            var values=new[]{relation.Value,relation.SecondaryValue}.Where(value=>value.HasValue)
                .Select(value=>value.Value.ToString("0.###",CultureInfo.CurrentCulture));
            if(relation.Value.HasValue)name+=" "+String.Join(" / ",values)+(relation.Kind==ConstraintKind.AspectRatio ? "" : " pt");
            return name+"（"+String.Join(" → ",relation.ShapeIds.Select(id=>(string)((dynamic)context.Shapes[id]).Name))+"）";
        }

        private void Commit(EditPlan plan)
        {
            dynamic presentation=plan.Context.Presentation;
            string documentToken=(string)presentation.Tags[ShapeIdentityManager.DocumentTagName];
            string rawConstraintTag=(string)presentation.Tags[NativeConstraintStore.TagName];
            application.StartNewUndoEntry();
            try
            {
                foreach(var tag in plan.NewIdentities) ((dynamic)tag.Shape).Tags.Add(PowerPointLayout.TagName,tag.Id);
                PowerPointLayout.WriteDocument(plan.Context.Presentation,plan.Document);
                PowerPointLayout.Apply(plan.Context,plan.Solved);
                normalizeIdentities(plan.Context.Presentation);
                var actual=PowerPointLayout.Snapshot(plan.Context);
                if(LayoutSolver.MaxResidual(actual,PowerPointLayout.BindRelations(plan.Context,plan.Document))>LayoutSolver.PositionTolerance)
                    throw new InvalidOperationException("PowerPoint 几何读回不满足约束。");
                RefreshSelectionBounds(plan.Context);
                log("constraint-command","{\"slideId\":"+plan.Context.SlideId+",\"constraints\":"+plan.Document.Constraints.Count+"}");
            }
            catch(Exception failure)
            {
                var failures=new List<Exception>{failure};
                Action<Action> restore=action=>{try{action();}catch(Exception error){failures.Add(error);}};
                restore(()=>PowerPointLayout.Apply(plan.Context,plan.Context.Geometry));
                restore(()=>RestoreTag(plan.Context.Presentation,NativeConstraintStore.TagName,rawConstraintTag));
                foreach(var tag in plan.NewIdentities)
                {
                    restore(()=>RestoreTag(tag.Shape,PowerPointLayout.TagName,tag.Before));
                    restore(()=>RestoreTag(tag.Shape,ShapeIdentityManager.OwnerTagName,tag.OwnerBefore));
                }
                restore(()=>RestoreTag(plan.Context.Presentation,ShapeIdentityManager.DocumentTagName,documentToken));
                restore(()=>RefreshSelectionBounds(plan.Context));
                if(failures.Count>1)throw new AggregateException("操作失败，恢复过程中也出现错误。",failures);
                throw;
            }
            finally { application.StartNewUndoEntry(); }
        }

        internal static void RestoreTag(object target,string key,string value)
        { if(String.IsNullOrEmpty(value)) ((dynamic)target).Tags.Delete(key); else ((dynamic)target).Tags.Add(key,value); }

        private void NormalizeCurrent()
        {
            if((int)application.Presentations.Count==0)throw new InvalidOperationException("请先打开演示文稿。");
            normalizeIdentities((object)application.ActivePresentation);
        }

        internal static void RefreshSelectionBounds(LayoutContext context)
        {
            var selected=new HashSet<int>(SelectedShapes(context.Window).Select(s=>(int)((dynamic)s).Id));
            if(selected.Count==0)return;
            dynamic slide=context.Slide;dynamic window=context.Window;
            var indices=new List<object>();
            for(int i=1;i<=(int)slide.Shapes.Count;i++)if(selected.Contains((int)slide.Shapes[i].Id))indices.Add(i);
            window.Selection.Unselect();
            if(indices.Count>0)slide.Shapes.Range(indices.ToArray()).Select();
        }

        private LayoutContext ReadCurrent()
        {
            if((int)application.Presentations.Count==0) throw new InvalidOperationException("请先打开演示文稿。");
            dynamic window=application.ActiveWindow;
            if(window==null || (int)window.ViewType!=9) throw new InvalidOperationException("请切换到普通编辑视图。");
            dynamic slide=window.View.Slide;
            object presentation=window.Presentation;
            var context=PowerPointLayout.NewContext(presentation,(object)slide,(object)window,PowerPointLayout.ReadDocument(presentation));
            for(int i=1;i<=(int)slide.Shapes.Count;i++)
            {
                object shape=(object)slide.Shapes[i];
                string id=PowerPointLayout.ReadId(shape);
                if(id==null || !IsSupported(shape)) continue;
                if(context.Shapes.ContainsKey(id)) throw new InvalidOperationException("对象编号重复，无法确定约束对象。");
                context.Shapes.Add(id,shape); context.Geometry.Add(PowerPointLayout.Geometry(shape,id));
            }
            return context;
        }

        public ConstraintPaneModel BuildModel() { return BuildModel(false); }
        public ConstraintPaneModel BuildPaneModel() { return BuildModel(true); }

        private ConstraintPaneModel BuildModel(bool includeFreedom)
        {
            LayoutContext context=ReadCurrent();
            dynamic slide=context.Slide;
            var selected=SelectedShapes(context.Window);
            var selectedNative=new HashSet<int>(selected.Select(s=>(int)((dynamic)s).Id));
            var defaults=selected.Where(IsSupported).ToDictionary(s=>(int)((dynamic)s).Id,s=>PowerPointLayout.Geometry(s,PowerPointLayout.ReadId(s)??Guid.NewGuid().ToString("D")));
            var model=new ConstraintPaneModel {PageLabel="第 "+(int)slide.SlideIndex+" 页 · "+(string)slide.Name,Options=ConstraintCatalog.Options(defaults.Values.ToList())};
            var parameterById=context.Document.Parameters.ToDictionary(item=>item.Id,StringComparer.Ordinal);
            if(includeFreedom)
            {
                var counts=context.Document.Constraints.Where(item=>item.ParameterId!=null)
                    .GroupBy(item=>item.ParameterId,StringComparer.Ordinal).ToDictionary(group=>group.Key,group=>group.Count(),StringComparer.Ordinal);
                foreach(var item in context.Document.Parameters)
                    model.Parameters.Add(new ConstraintPaneParameter {Id=item.Id,Name=item.Name,Value=item.Value,
                        References=counts.ContainsKey(item.Id)?counts[item.Id]:0});
            }
            var names=new Dictionary<string,string>(StringComparer.Ordinal);
            var uiKeys=new Dictionary<string,string>(StringComparer.Ordinal);
            var geometryById=context.Geometry.ToDictionary(s=>s.Id,StringComparer.Ordinal);
            var relations=context.Document.Constraints.Where(r=>r.SlideId==context.SlideId).ToList();
            if(includeFreedom)UpdateFreedomCache(context,defaults.Values.Where(shape=>geometryById.ContainsKey(shape.Id)));
            for(int i=1;i<=(int)slide.Shapes.Count;i++)
            {
                dynamic shape=slide.Shapes[i]; string id=PowerPointLayout.ReadId((object)shape);
                string uiKey=UiKey(context.SlideId,(int)shape.Id);
                int nativeId=(int)shape.Id;
                ShapeGeometry geometry,unmarked;
                GeometryFreedom freedom=null;
                if(includeFreedom && defaults.TryGetValue(nativeId,out unmarked))
                {
                    if(id==null)freedom=LayoutDiagnostics.AnalyzeShape(new[]{unmarked},Array.Empty<LayoutConstraint>(),unmarked.Id);
                    else freedomByShape.TryGetValue(id,out freedom);
                }
                model.Objects.Add(new ConstraintPaneObject {Id=uiKey,Name=(string)shape.Name,Selected=selectedNative.Contains((int)shape.Id),
                    Geometry=id!=null && geometryById.TryGetValue(id,out geometry)?geometry:null,
                    Freedom=freedom,
                    UnsupportedReason=ShapeCapabilities.UnsupportedReason((object)shape),
                    RelationCount=id==null?0:relations.Count(r=>r.ShapeIds.Contains(id))});
                if(id!=null) {names[id]=(string)shape.Name;uiKeys[id]=uiKey;}
            }
            foreach(var relation in relations)
            {
                string status;
                if(relation.ShapeIds.Any(id=>!context.Shapes.ContainsKey(id))) status="对象缺失或不支持";
                else if(!relation.Enabled) status="已禁用";
                else status=LayoutSolver.MaxResidual(context.Geometry,new[]{relation.BindPageSize(context.PageWidth,context.PageHeight)})<=LayoutSolver.PositionTolerance?"已满足":"未满足";
                model.Relations.Add(new ConstraintPaneRelation {Id=relation.Id,Kind=relation.Kind.ToString(),Label=ConstraintCatalog.Label(relation.Kind),
                    MemberIds=relation.ShapeIds.Select(id=>uiKeys.ContainsKey(id)?uiKeys[id]:id).ToList(),
                    MemberNames=relation.ShapeIds.Select(id=>names.ContainsKey(id)?names[id]:"（对象缺失）").ToList(),
                    Value=relation.Value,Second=relation.SecondaryValue,Enabled=relation.Enabled,Status=status,
                    ParameterId=relation.ParameterId,ParameterName=relation.ParameterId==null?null:parameterById[relation.ParameterId].Name});
            }
            model.Message="已选 "+selected.Count+" 个对象 · 当前页 "+relations.Count+" 条关系";
            return model;
        }

        private void UpdateFreedomCache(LayoutContext context,IEnumerable<ShapeGeometry> selected)
        {
            var ids=selected.Select(shape=>shape.Id).OrderBy(id=>id,StringComparer.Ordinal).ToArray();
            var connected=LayoutSolver.ConnectedConstraints(context.Constraints,ids);
            var members=new HashSet<string>(connected.SelectMany(relation=>relation.ShapeIds).Concat(ids),StringComparer.Ordinal);
            var geometry=context.Geometry.Where(shape=>members.Contains(shape.Id)).ToList();
            var key=new StringBuilder().Append(context.SlideId).Append('|').Append(String.Join(",",ids));
            key.AppendFormat(CultureInfo.InvariantCulture,"|{0:R},{1:R}",context.PageWidth,context.PageHeight);
            foreach(var shape in geometry)key.AppendFormat(CultureInfo.InvariantCulture,"|{0}:{1:R},{2:R},{3:R},{4:R}",shape.Id,shape.Left,shape.Top,shape.Width,shape.Height);
            var used=new HashSet<string>(connected.Select(relation=>relation.ParameterId).Where(id=>id!=null),StringComparer.Ordinal);
            key.Append('|').Append(new ConstraintDocument(connected,context.Document.Parameters.Where(parameter=>used.Contains(parameter.Id))).ToXml());
            string signature=key.ToString();
            if(freedomKey==signature)return;
            freedomByShape=new Dictionary<string,GeometryFreedom>(StringComparer.Ordinal);
            foreach(string id in ids)
            {
                try {freedomByShape[id]=LayoutDiagnostics.AnalyzeShape(geometry,connected,id);}
                catch(ConstraintConflictException) { }
                catch(ArgumentException) { } // A referenced object is missing or unsupported in the current view.
            }
            freedomKey=signature;
        }

        public void SelectShapes(IReadOnlyList<string> ids)
        {
            if(ids==null) throw new ArgumentNullException("ids");
            dynamic window=application.ActiveWindow; dynamic slide=window.View.Slide;
            if(ids.Count==0){window.Selection.Unselect();return;}
            var requested=new HashSet<string>(ids,StringComparer.Ordinal);
            var indices=new List<int>();
            for(int i=1;i<=(int)slide.Shapes.Count;i++)
            {
                dynamic shape=slide.Shapes[i];string id=PowerPointLayout.ReadId((object)shape);
                if(requested.Contains(UiKey((int)slide.SlideID,(int)shape.Id)) || (id!=null && requested.Contains(id))) indices.Add(i);
            }
            if(indices.Count==0) throw new InvalidOperationException("参与对象已不在当前页。");
            slide.Shapes.Range(indices.Cast<object>().ToArray()).Select();
        }

        private static string UiKey(int slide,int shape)
        { return "shape:"+slide.ToString(CultureInfo.InvariantCulture)+"/"+shape.ToString(CultureInfo.InvariantCulture); }
        private static List<object> SelectedShapes(object windowObject)
        {
            dynamic selection=((dynamic)windowObject).Selection;
            var result=new List<object>();
            if((int)selection.Type!=2 && (int)selection.Type!=3) return result;
            if((bool)selection.HasChildShapeRange) return result;
            for(int i=1;i<=(int)selection.ShapeRange.Count;i++) result.Add((object)selection.ShapeRange[i]);
            return result;
        }
        private static bool IsSupported(object shape)
        { return ShapeCapabilities.IsSupported(shape); }
        private static ConstraintKind ParseKind(string value)
        {
            ConstraintKind kind;
            if(!Enum.TryParse(value,false,out kind) || !Enum.IsDefined(typeof(ConstraintKind),kind)) throw new ArgumentException("未知约束类型。");
            return kind;
        }
        internal sealed class NewIdentity {public object Shape;public string Id,Before,OwnerBefore;}
        internal sealed class EditPlan
        {
            public LayoutContext Context;public ConstraintDocument Document;public IReadOnlyList<ShapeGeometry> Solved;
            public List<NewIdentity> NewIdentities=new List<NewIdentity>();public string RelationId;
            public LayoutConstraint Attempted;
        }

        public sealed class ConflictRelation
        {
            public string Id {get;}
            public string Text {get;}
            public bool CanDisable {get;}
            internal ConflictRelation(string id,string text,bool canDisable){Id=id;Text=text;CanDisable=canDisable;}
            public override string ToString(){return Text;}
        }

        public sealed class DraftConflict : ConstraintConflictException
        {
            internal EditPlan Plan {get;}
            public string Attempted {get;}
            public IReadOnlyList<ConflictRelation> Choices {get;}
            internal DraftConflict(EditPlan plan,string attempted,List<ConflictRelation> choices,
                IReadOnlyList<LayoutConstraint> relations,Exception cause)
                : base("这些约束不能同时满足。"+Environment.NewLine+String.Join(Environment.NewLine,choices.Select(choice=>choice.Text)),relations,cause)
            {Plan=plan;Attempted=attempted;Choices=choices.AsReadOnly();}
        }
    }
}
