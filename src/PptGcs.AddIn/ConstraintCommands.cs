using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Ribbon, task pane and radial menu all use these atomic document commands.
    public sealed class ConstraintCommands
    {
        private readonly dynamic application;
        private readonly Action<object> normalizeIdentities;
        private readonly Action<string,string> log;
        public ConstraintCommands(object application, Action<object> normalizeIdentities, Action<string,string> log)
        { this.application=application; this.normalizeIdentities=normalizeIdentities; this.log=log; }

        public string Add(string kind, double? value, double? second)
        {
            NormalizeCurrent();
            EditPlan plan=PlanAdd(ParseKind(kind),value,second);
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

        public void Update(string id, double? value, double? second)
        { Change(id, r=>new LayoutConstraint(r.Id,r.SlideId,r.Kind,r.ShapeIds,value,second,r.Enabled)); }
        public void SetEnabled(string id,bool enabled)
        { Change(id,r=>new LayoutConstraint(r.Id,r.SlideId,r.Kind,r.ShapeIds,r.Value,r.SecondaryValue,enabled)); }
        public void Delete(string id) { Change(id,r=>null); }

        public void DeleteForSelection()
        {
            NormalizeCurrent();
            LayoutContext context=ReadCurrent();
            var ids=new HashSet<string>(SelectedShapes(context.Window).Select(PowerPointLayout.ReadId).Where(id=>id!=null));
            if(ids.Count==0) throw new InvalidOperationException("请选择有约束的对象。");
            var next=new ConstraintDocument(context.Document.Constraints.Where(r=>r.SlideId!=context.SlideId || !r.ShapeIds.Any(ids.Contains)));
            if(next.ToXml()==context.Document.ToXml()) return;
            Commit(PlanChange(context,next));
        }

        private void Change(string id,Func<LayoutConstraint,LayoutConstraint> change)
        {
            NormalizeCurrent();
            LayoutContext context=ReadCurrent();
            string normalized=ShapeIdentity.Normalize(id);
            LayoutConstraint current=context.Document.Constraints.SingleOrDefault(r=>r.Id==normalized && r.SlideId==context.SlideId);
            if(current==null) throw new InvalidOperationException("该约束已不存在或不属于当前页。");
            LayoutConstraint replacement=change(current);
            var next=new ConstraintDocument(context.Document.Constraints.Select(r=>r.Id==current.Id?replacement:r).Where(r=>r!=null));
            if(next.ToXml()==context.Document.ToXml()) return;
            Commit(PlanChange(context,next));
        }

        private EditPlan PlanAdd(ConstraintKind kind,double? value,double? second)
        {
            LayoutContext context=ReadCurrent();
            if((bool)((dynamic)context.Window).Selection.HasChildShapeRange)
                throw new InvalidOperationException("请先退出组合内部编辑，并选中组合整体建立约束。");
            var selected=SelectedShapes(context.Window);
            int minimum=LayoutConstraint.GetMinimumShapeCount(kind), maximum=LayoutConstraint.GetMaximumShapeCount(kind);
            if(selected.Count<minimum || (maximum>0 && selected.Count>maximum))
                throw new InvalidOperationException(maximum==minimum ? "此约束需要选中 "+minimum+" 个对象。" : "请选中至少 "+minimum+" 个对象。");
            foreach(object shape in selected) PowerPointLayout.ValidateShape(shape);
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
            bool ordered=LayoutConstraint.PreservesOrder(kind);
            if(context.Document.Constraints.Any(r=>r.SlideId==context.SlideId && r.Kind==kind &&
                (ordered?r.ShapeIds.SequenceEqual(ids):r.ShapeIds.Count==ids.Count && new HashSet<string>(r.ShapeIds).SetEquals(ids))))
                throw new InvalidOperationException("这些对象已经存在“"+ConstraintCatalog.Label(kind)+"”关系，请在窗格中修改或启用它。");
            var relation=new LayoutConstraint(Guid.NewGuid().ToString("D"),context.SlideId,kind,ids,value,second);
            var next=new ConstraintDocument(context.Document.Constraints.Concat(new[]{relation}));
            EditPlan plan=PlanChange(context,next);
            plan.NewIdentities=tags;
            plan.RelationId=relation.Id;
            return plan;
        }

        private static EditPlan PlanChange(LayoutContext context,ConstraintDocument next)
        {
            var active=next.Constraints.Where(r=>r.SlideId==context.SlideId && r.Enabled).ToList();
            var missing=active.SelectMany(r=>r.ShapeIds).Where(id=>!context.Shapes.ContainsKey(id)).Distinct().ToList();
            if(missing.Count>0) throw new InvalidOperationException("存在引用缺失或不支持对象的约束，请先在约束窗格禁用或删除该关系。");
            try {return new EditPlan { Context=context,Document=next,Solved=LayoutSolver.Normalize(context.Geometry,active) };}
            catch(ConstraintConflictException)
            {throw new ConstraintConflictException("这些约束不能同时满足，请检查固定尺寸、固定位置或已有关系。");}
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
                if(LayoutSolver.MaxResidual(actual,plan.Document.Constraints.Where(r=>r.SlideId==plan.Context.SlideId && r.Enabled))>LayoutSolver.PositionTolerance)
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

        private static void RestoreTag(object target,string key,string value)
        { if(String.IsNullOrEmpty(value)) ((dynamic)target).Tags.Delete(key); else ((dynamic)target).Tags.Add(key,value); }

        private void NormalizeCurrent()
        {
            if((int)application.Presentations.Count==0)throw new InvalidOperationException("请先打开演示文稿。");
            normalizeIdentities((object)application.ActivePresentation);
        }

        private static void RefreshSelectionBounds(LayoutContext context)
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
            var context=new LayoutContext {Window=window,Presentation=window.Presentation,Slide=slide,SlideId=(int)slide.SlideID,
                Document=PowerPointLayout.ReadDocument((object)window.Presentation)};
            context.Constraints=context.Document.Constraints.Where(r=>r.SlideId==context.SlideId && r.Enabled).ToList();
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

        public ConstraintPaneModel BuildModel()
        {
            LayoutContext context=ReadCurrent();
            dynamic slide=context.Slide;
            var selected=SelectedShapes(context.Window);
            var selectedNative=new HashSet<int>(selected.Select(s=>(int)((dynamic)s).Id));
            var defaults=selected.Where(IsSupported).Select(s=>PowerPointLayout.Geometry(s,PowerPointLayout.ReadId(s)??Guid.NewGuid().ToString("D"))).ToList();
            var model=new ConstraintPaneModel {PageLabel="第 "+(int)slide.SlideIndex+" 页 · "+(string)slide.Name,Options=ConstraintCatalog.Options(defaults)};
            var names=new Dictionary<string,string>(StringComparer.Ordinal);
            var uiKeys=new Dictionary<string,string>(StringComparer.Ordinal);
            var geometryById=context.Geometry.ToDictionary(s=>s.Id,StringComparer.Ordinal);
            var relations=context.Document.Constraints.Where(r=>r.SlideId==context.SlideId).ToList();
            for(int i=1;i<=(int)slide.Shapes.Count;i++)
            {
                dynamic shape=slide.Shapes[i]; string id=PowerPointLayout.ReadId((object)shape);
                string uiKey=UiKey(context.SlideId,(int)shape.Id);
                ShapeGeometry geometry;
                model.Objects.Add(new ConstraintPaneObject {Id=uiKey,Name=(string)shape.Name,Selected=selectedNative.Contains((int)shape.Id),
                    Geometry=id!=null && geometryById.TryGetValue(id,out geometry)?geometry:null,
                    UnsupportedReason=ShapeCapabilities.UnsupportedReason((object)shape),
                    RelationCount=id==null?0:relations.Count(r=>r.ShapeIds.Contains(id))});
                if(id!=null) {names[id]=(string)shape.Name;uiKeys[id]=uiKey;}
            }
            foreach(var relation in relations)
            {
                string status;
                if(relation.ShapeIds.Any(id=>!context.Shapes.ContainsKey(id))) status="对象缺失或不支持";
                else if(!relation.Enabled) status="已禁用";
                else status=LayoutSolver.MaxResidual(context.Geometry,new[]{relation})<=LayoutSolver.PositionTolerance?"已满足":"未满足";
                model.Relations.Add(new ConstraintPaneRelation {Id=relation.Id,Kind=relation.Kind.ToString(),Label=ConstraintCatalog.Label(relation.Kind),
                    MemberIds=relation.ShapeIds.Select(id=>uiKeys.ContainsKey(id)?uiKeys[id]:id).ToList(),
                    MemberNames=relation.ShapeIds.Select(id=>names.ContainsKey(id)?names[id]:"（对象缺失）").ToList(),
                    Value=relation.Value,Second=relation.SecondaryValue,Enabled=relation.Enabled,Status=status});
            }
            model.Message="已选 "+selected.Count+" 个对象 · 当前页 "+relations.Count+" 条关系";
            return model;
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
        private sealed class NewIdentity {public object Shape;public string Id,Before,OwnerBefore;}
        private sealed class EditPlan
        {
            public LayoutContext Context;public ConstraintDocument Document;public IReadOnlyList<ShapeGeometry> Solved;
            public List<NewIdentity> NewIdentities=new List<NewIdentity>();public string RelationId;
        }
    }
}
