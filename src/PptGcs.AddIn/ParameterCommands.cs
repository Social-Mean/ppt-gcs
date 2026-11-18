using System;
using System.Collections.Generic;
using System.Linq;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    internal sealed class ParameterCommands
    {
        private readonly dynamic application;

        public ParameterCommands(object application)
        { this.application = application; }

        public static LayoutParameter Find(ConstraintDocument document, string id)
        {
            if (id == null) return null;
            string normalized = ShapeIdentity.Normalize(id);
            return document.Parameters.SingleOrDefault(item => item.Id == normalized) ??
                throw new InvalidOperationException("该命名参数已不存在，请重新选择。");
        }

        public string Add(string name, double value)
        {
            object presentation = ActivePresentation();
            var before = PowerPointLayout.ReadDocument(presentation);
            var parameter = new LayoutParameter(Guid.NewGuid().ToString("D"), name, value);
            Commit(presentation, new ConstraintDocument(before.Constraints, before.Parameters.Concat(new[] { parameter })),
                new List<SlideEdit>());
            return parameter.Id;
        }

        public void Update(string id, string name, double value)
        {
            object presentation = ActivePresentation();
            var before = PowerPointLayout.ReadDocument(presentation);
            var parameter = Find(before, id) ?? throw new ArgumentNullException(nameof(id));
            var replacement = new LayoutParameter(parameter.Id, name, value);
            if (parameter.Name == replacement.Name && parameter.Value == replacement.Value) return;
            var next = new ConstraintDocument(before.Constraints,
                before.Parameters.Select(item => item.Id == parameter.Id ? replacement : item));
            var edits = parameter.Value == replacement.Value ? new List<SlideEdit>() : PlanGeometry(presentation, next, parameter.Id);
            Commit(presentation, next, edits);
        }

        public void Delete(string id)
        {
            object presentation = ActivePresentation();
            var before = PowerPointLayout.ReadDocument(presentation);
            var parameter = Find(before, id) ?? throw new ArgumentNullException(nameof(id));
            int references = before.Constraints.Count(item => item.ParameterId == parameter.Id);
            if (references != 0) throw new InvalidOperationException("仍有 " + references + " 条约束引用“" + parameter.Name + "”。请先解除这些绑定。");
            Commit(presentation, new ConstraintDocument(before.Constraints, before.Parameters.Where(item => item.Id != parameter.Id)),
                new List<SlideEdit>());
        }

        private object ActivePresentation()
        {
            if ((int)application.Presentations.Count == 0) throw new InvalidOperationException("请先打开演示文稿。");
            return (object)application.ActivePresentation;
        }

        private List<SlideEdit> PlanGeometry(object presentationObject, ConstraintDocument document, string parameterId)
        {
            dynamic presentation = presentationObject;
            var affected = document.Constraints.Where(item => item.Enabled && item.ParameterId == parameterId)
                .GroupBy(item => item.SlideId).ToDictionary(group => group.Key, group => group.SelectMany(item => item.ShapeIds).ToList());
            var edits = new List<SlideEdit>();
            if (affected.Count == 0) return edits;
            dynamic window = application.ActiveWindow;
            int activeSlideId = (int)window.ViewType == 9 ? (int)window.View.Slide.SlideID : -1;
            for (int index = 1; index <= (int)presentation.Slides.Count; index++)
            {
                object slide = (object)presentation.Slides[index];
                int slideId = (int)((dynamic)slide).SlideID;
                List<string> drivers;
                if (!affected.TryGetValue(slideId, out drivers)) continue;
                var scope = LayoutSolver.ConnectedConstraints(document.Constraints.Where(item => item.SlideId == slideId), drivers);
                var context = PowerPointLayout.ReadSlide(presentationObject, slide, slideId == activeSlideId ? (object)window : null,
                    document.WithConstraints(scope));
                try
                {
                    edits.Add(new SlideEdit(context, LayoutSolver.Normalize(context.Geometry, context.Constraints)));
                }
                catch (ConstraintConflictException error)
                {
                    var witness = LayoutDiagnostics.ExplainConflict(context.Geometry, context.Constraints);
                    string relations = String.Join(Environment.NewLine, witness.Select(item => ConstraintCommands.DescribeRelation(context, item)));
                    throw new ConstraintConflictException("第 " + index + " 页的约束与参数新值冲突。" + Environment.NewLine + relations, witness, error);
                }
            }
            return edits;
        }

        private void Commit(object presentation, ConstraintDocument document, List<SlideEdit> edits)
        {
            NativeConstraintStore.EncodeXml(document.ToXml());
            string oldTag = (string)((dynamic)presentation).Tags[NativeConstraintStore.TagName];
            var written = new List<SlideEdit>();
            application.StartNewUndoEntry();
            try
            {
                foreach (var edit in edits)
                {
                    written.Add(edit);
                    PowerPointLayout.Apply(edit.Context, edit.Target);
                    if (LayoutSolver.MaxResidual(PowerPointLayout.Snapshot(edit.Context), edit.Context.Constraints) > LayoutSolver.PositionTolerance)
                        throw new InvalidOperationException("PowerPoint 几何读回不满足命名参数约束。");
                }
                PowerPointLayout.WriteDocument(presentation, document);
                foreach (var edit in written.Where(item => item.Context.Window != null))
                    ConstraintCommands.RefreshSelectionBounds(edit.Context);
            }
            catch (Exception failure)
            {
                var failures = new List<Exception> { failure };
                foreach (var edit in written)
                    try { PowerPointLayout.Apply(edit.Context, edit.Context.Geometry); } catch (Exception error) { failures.Add(error); }
                try { ConstraintCommands.RestoreTag(presentation, NativeConstraintStore.TagName, oldTag); } catch (Exception error) { failures.Add(error); }
                if (failures.Count > 1) throw new AggregateException("参数更新失败，恢复文稿时也发生错误。", failures);
                throw;
            }
            finally { application.StartNewUndoEntry(); }
        }

        private sealed class SlideEdit
        {
            public LayoutContext Context { get; }
            public IReadOnlyList<ShapeGeometry> Target { get; }
            public SlideEdit(LayoutContext context, IReadOnlyList<ShapeGeometry> target)
            { Context = context; Target = target; }
        }
    }
}
