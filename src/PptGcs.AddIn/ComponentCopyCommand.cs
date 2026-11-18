using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    internal static class ComponentCopyCommand
    {
        public static IReadOnlyList<string> Apply(object applicationObject, int targetSlideId, bool shareParameters)
        {
            dynamic application = applicationObject;
            if ((int)application.Presentations.Count == 0) throw new InvalidOperationException("请先打开演示文稿。");
            dynamic window = application.ActiveWindow;
            if ((int)window.ViewType != 9) throw new InvalidOperationException("请切换到普通幻灯片编辑视图。");
            dynamic selection = window.Selection;
            if ((int)selection.Type != 2 || (bool)selection.HasChildShapeRange)
                throw new InvalidOperationException("请选中要复制的顶层对象。复制组合时请选择组合整体。");
            dynamic sourceRange = selection.ShapeRange;
            int count = (int)sourceRange.Count;
            if (count == 0) throw new InvalidOperationException("请先选择要复制的对象。");

            dynamic presentation = window.Presentation;
            dynamic sourceSlide = window.View.Slide;
            object targetObject = FindSlide((object)presentation, targetSlideId);
            dynamic target = targetObject;
            int sourceSlideId = (int)sourceSlide.SlideID;
            var document = PowerPointLayout.ReadDocument((object)presentation);
            var mapping = new Dictionary<string, string>(StringComparer.Ordinal);
            var originals = new Dictionary<string, object>(StringComparer.Ordinal);
            for (int index = 1; index <= count; index++)
            {
                object shape = (object)sourceRange[index];
                string id = PowerPointLayout.ReadId(shape);
                if (id == null) continue;
                if (mapping.ContainsKey(id)) throw new InvalidOperationException("所选对象含有重复编号，请重新选择原件和副本。");
                mapping.Add(id, NewId());
                originals.Add(id, shape);
            }
            var next = ComponentCopy.Plan(document, sourceSlideId, targetSlideId, mapping, shareParameters);
            if (next.Constraints.Count != document.Constraints.Count) NativeConstraintStore.EncodeXml(next.ToXml());
            var copyRules = next.Constraints.Skip(document.Constraints.Count).Where(item => item.Enabled).ToList();
            var scope = next.WithConstraints(copyRules);
            var context = PowerPointLayout.NewContext((object)presentation, targetObject, (object)window, scope);
            var required = new HashSet<string>(copyRules.SelectMany(item => item.ShapeIds), StringComparer.Ordinal);
            var predicted = new List<ShapeGeometry>();
            foreach (var entry in mapping.Where(entry => required.Contains(entry.Value)))
            {
                PowerPointLayout.ValidateShape(originals[entry.Key]);
                predicted.Add(PowerPointLayout.Geometry(originals[entry.Key], entry.Value));
            }
            LayoutSolver.Normalize(predicted, context.Constraints);

            var existing = new HashSet<int>();
            for (int index = 1; index <= (int)target.Shapes.Count; index++) existing.Add((int)target.Shapes[index].Id);
            string previousTag = (string)presentation.Tags[NativeConstraintStore.TagName];
            string previousToken = (string)presentation.Tags[ShapeIdentityManager.DocumentTagName];
            string token = String.IsNullOrEmpty(previousToken) ? NewId() : ShapeIdentity.Normalize(previousToken);
            bool sameSlide = sourceSlideId == targetSlideId;
            uint clipboardSequence = 0;
            if (!sameSlide)
            {
                sourceRange.Copy();
                clipboardSequence = GetClipboardSequenceNumber();
            }

            dynamic pasted = null;
            var copiedIds = new List<string>();
            application.StartNewUndoEntry();
            try
            {
                if (!sameSlide && clipboardSequence != GetClipboardSequenceNumber())
                    throw new InvalidOperationException("剪贴板内容已改变，请重新复制约束组件。");
                pasted = sameSlide ? sourceRange.Duplicate() : target.Shapes.Paste();
                if ((int)pasted.Count != count) throw new InvalidOperationException("PowerPoint 未完整复制所选对象。");
                var matches = new HashSet<string>(StringComparer.Ordinal);
                var newShapes = new HashSet<int>();
                if (token != previousToken) presentation.Tags.Add(ShapeIdentityManager.DocumentTagName, token);
                for (int index = 1; index <= count; index++)
                {
                    dynamic shape = pasted[index];
                    int nativeId = (int)shape.Id;
                    if (existing.Contains(nativeId) || !newShapes.Add(nativeId))
                        throw new InvalidOperationException("无法确认 PowerPoint 返回的是独立副本。");
                    string originalId = PowerPointLayout.ReadId((object)shape);
                    string id;
                    if (originalId == null) id = NewId();
                    else if (!mapping.TryGetValue(originalId, out id) || !matches.Add(originalId))
                        throw new InvalidOperationException("无法匹配副本和原有约束成员。");
                    shape.Tags.Add(PowerPointLayout.TagName, id);
                    shape.Tags.Add(ShapeIdentityManager.OwnerTagName, token + "/" + targetSlideId.ToString(CultureInfo.InvariantCulture) +
                        "/" + nativeId.ToString(CultureInfo.InvariantCulture));
                    if (PowerPointLayout.ReadId((object)shape) != id)
                        throw new InvalidOperationException("PowerPoint 未保留副本编号。");
                    if (required.Contains(id))
                    {
                        PowerPointLayout.ValidateShape((object)shape);
                        context.Shapes.Add(id, (object)shape);
                        context.Geometry.Add(PowerPointLayout.Geometry((object)shape, id));
                    }
                    copiedIds.Add(id);
                }
                if (matches.Count != mapping.Count) throw new InvalidOperationException("复制时丢失了部分约束成员编号。");
                PowerPointLayout.Apply(context, LayoutSolver.Normalize(context.Geometry, context.Constraints));
                if (LayoutSolver.MaxResidual(PowerPointLayout.Snapshot(context), context.Constraints) > LayoutSolver.PositionTolerance)
                    throw new InvalidOperationException("复制后的几何未满足内部约束。");
                if (next.Constraints.Count != document.Constraints.Count) PowerPointLayout.WriteDocument((object)presentation, next);
            }
            catch (Exception failure)
            {
                var failures = new List<Exception> { failure };
                for (int index = (int)target.Shapes.Count; index >= 1; index--)
                    try { if (!existing.Contains((int)target.Shapes[index].Id)) target.Shapes[index].Delete(); }
                    catch (Exception error) { failures.Add(error); }
                foreach (var tag in new[] { Tuple.Create(NativeConstraintStore.TagName, previousTag), Tuple.Create(ShapeIdentityManager.DocumentTagName, previousToken) })
                    try
                    {
                        if ((string)presentation.Tags[tag.Item1] != tag.Item2)
                            ConstraintCommands.RestoreTag((object)presentation, tag.Item1, tag.Item2);
                    }
                    catch (Exception error) { failures.Add(error); }
                if (failures.Count > 1) throw new AggregateException("复制失败，恢复文稿时也发生错误。", failures);
                throw;
            }
            finally { application.StartNewUndoEntry(); }

            if (!sameSlide) window.View.GotoSlide((int)target.SlideIndex);
            pasted.Select();
            return copiedIds;
        }

        private static object FindSlide(object presentationObject, int id)
        {
            dynamic presentation = presentationObject;
            for (int index = 1; index <= (int)presentation.Slides.Count; index++)
            {
                object slide = (object)presentation.Slides[index];
                if ((int)((dynamic)slide).SlideID == id) return slide;
            }
            throw new InvalidOperationException("目标幻灯片已不存在，请重新选择。");
        }

        private static string NewId() { return Guid.NewGuid().ToString("D").ToUpperInvariant(); }

        [DllImport("user32.dll")]
        private static extern uint GetClipboardSequenceNumber();
    }
}
