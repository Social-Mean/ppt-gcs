using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace PptGcs.AddIn
{
    internal sealed class SlideLayers
    {
        private const string StackTag = "PPTGCS_LAYERS";
        private const string MemberTag = "PPTGCS_LAYER";
        internal const string BaseId = "BASE";
        internal static readonly string[] ArrangeCommands =
            { "ObjectBringToFront", "ObjectBringForward", "ObjectSendBackward", "ObjectSendToBack" };

        internal sealed class Layer
        {
            public readonly string Id, Name;
            public Layer(string id, string name) { Id = id; Name = name; }
        }

        internal sealed class Member
        {
            public int Id, Index, Position;
            public object Shape;
            public string Owner, LayerId, StoredTag;
        }

        internal readonly object Window, Slide;
        internal readonly Layer[] Stack;
        internal readonly int SlideId;
        internal readonly string Encoded;
        private readonly string scope;
        internal bool Configured { get { return Encoded.Length != 0; } }

        private SlideLayers(object window, object slide)
        {
            Window = window;
            Slide = slide;
            SlideId = (int)((dynamic)slide).SlideID;
            Encoded = (string)((dynamic)slide).Tags[StackTag];
            if (!Configured)
            {
                Stack = new[] { new Layer(BaseId, "默认图层") };
                return;
            }
            var root = XElement.Parse(NativeConstraintStore.DecodeXml(Encoded));
            scope = (string)root.Attribute("scope");
            Stack = root.Elements("layer").Select(element => new Layer((string)element.Attribute("id"), (string)element.Attribute("name"))).ToArray();
            Guid token;
            if (root.Name != "layers" || !Guid.TryParseExact(scope, "N", out token) || token == Guid.Empty ||
                Stack.Length == 0 || Stack.Any(layer => String.IsNullOrWhiteSpace(layer.Id) || String.IsNullOrWhiteSpace(layer.Name)) ||
                Stack.Select(layer => layer.Id).Distinct().Count() != Stack.Length)
                throw new InvalidOperationException("图层数据有误。");
        }

        internal static SlideLayers Read(object windowObject)
        {
            if (windowObject == null) return null;
            dynamic window = windowObject;
            if ((int)window.ViewType != 9) return null;
            // In Normal view the active pane may be a thumbnail or notes pane.
            object slide = window.View.Slide;
            return slide == null ? null : new SlideLayers(windowObject, slide);
        }

        internal static void EnsurePresentationOrder(object presentation)
        {
            dynamic slides = ((dynamic)presentation).Slides;
            for (int i = 1; i <= (int)slides.Count; i++)
                new SlideLayers(null, (object)slides[i]).EnsureOrder();
        }

        internal Member[] Members()
        {
            dynamic shapes = ((dynamic)Slide).Shapes;
            int count = (int)shapes.Count;
            var members = new List<Member>(count);
            var ranks = Stack.Select((layer, index) => new { layer.Id, Index = index })
                .ToDictionary(layer => layer.Id, layer => layer.Index);
            for (int index = 1; index <= count; index++)
            {
                dynamic shape = shapes[index];
                int id = (int)shape.Id;
                string owner = ShapeSuffix(id);
                string stored = (string)shape.Tags[MemberTag];
                members.Add(new Member {
                    Id = id, Index = index, Position = (int)shape.ZOrderPosition, Shape = (object)shape,
                    Owner = owner, LayerId = Stack[LayerRank((object)shape, stored, owner, ranks)].Id, StoredTag = stored
                });
            }
            return members.OrderByDescending(member => member.Position).ToArray();
        }

        private int LayerRank(object shapeObject, string stored, string owner, Dictionary<string, int> ranks)
        {
            if (Stack.Length == 1) return 0;
            if (stored.StartsWith(scope + ":", StringComparison.Ordinal) && stored.EndsWith(owner, StringComparison.Ordinal) &&
                stored.Length > scope.Length + 1 + owner.Length)
            {
                string id = stored.Substring(scope.Length + 1, stored.Length - scope.Length - 1 - owner.Length);
                int rank;
                if (ranks.TryGetValue(id, out rank)) return rank;
            }
            dynamic shape = shapeObject;
            if ((int)shape.Type != 6) return 0;
            // Newly made groups occupy the native stacking position of their
            // frontmost member. An explicit group assignment takes precedence.
            dynamic children = shape.GroupItems;
            int frontmost = Stack.Length - 1;
            for (int i = 1; i <= (int)children.Count; i++)
            {
                dynamic child = children[i];
                frontmost = Math.Min(frontmost, LayerRank((object)child, (string)child.Tags[MemberTag], ShapeSuffix((int)child.Id), ranks));
                if (frontmost == 0) break;
            }
            return frontmost;
        }

        private static string ShapeSuffix(int id) { return ":" + id.ToString(CultureInfo.InvariantCulture); }

        internal HashSet<int> Selected(bool includeGroupChildren = false)
        {
            dynamic selection = ((dynamic)Window).Selection;
            if ((int)selection.Type != 2) return new HashSet<int>();
            bool childSelection = (bool)selection.HasChildShapeRange;
            if (childSelection && !includeGroupChildren) return new HashSet<int>();
            dynamic range = childSelection ? selection.ChildShapeRange : selection.ShapeRange;
            var ids = new HashSet<int>();
            for (int i = 1; i <= (int)range.Count; i++) ids.Add((int)range[i].Id);
            return ids;
        }

        internal void Select(IEnumerable<Member> members)
        {
            object[] indexes = members.Select(member => (object)member.Index).ToArray();
            if (indexes.Length == 0)
                ((dynamic)Window).Selection.Unselect();
            else
                ((dynamic)Slide).Shapes.Range(indexes).Select();
        }

        internal LayerPaneRow[] PaneRows(Member[] members)
        {
            var byLayer = members.ToLookup(member => member.LayerId);
            return Stack.Select(layer =>
            {
                var objects = byLayer[layer.Id].Select(member => ObjectRow(layer.Id, member.Shape)).ToArray();
                return new LayerPaneRow(SlideId, layer.Id, null, layer.Name,
                    Summary(objects.Select(row => row.Visible)),
                    objects.Any(row => row.Locked == LayerRowState.Empty) ? LayerRowState.Empty : Summary(objects.Select(row => row.Locked)), objects);
            }).ToArray();
        }

        private LayerPaneRow ObjectRow(string layerId, object shapeObject)
        {
            dynamic shape = shapeObject;
            var children = Children(shapeObject).Select(child => ObjectRow(layerId, child)).ToArray();
            int? locked = ShapeCapabilities.ReadNativeLock(shapeObject);
            var lockState = !locked.HasValue || children.Any(child => child.Locked == LayerRowState.Empty)
                ? LayerRowState.Empty : locked.Value == 0 ? LayerRowState.Off : LayerRowState.On;
            return new LayerPaneRow(SlideId, layerId, (int)shape.Id, (string)shape.Name,
                (int)shape.Visible == 0 ? LayerRowState.Off : LayerRowState.On, lockState, children);
        }

        private static LayerRowState Summary(IEnumerable<LayerRowState> values)
        {
            var states = values.Distinct().ToArray();
            return states.Length == 0 ? LayerRowState.Empty : states.Length == 1 ? states[0] : LayerRowState.Mixed;
        }

        internal void SelectObject(string layerId, int shapeId)
        {
            object shape = FindObject(layerId, shapeId);
            if (shape == null) return;
            dynamic current = shape;
            while (true)
            {
                if ((int)current.Visible == 0) return;
                if ((int)current.Child == 0) break;
                current = current.ParentGroup;
            }
            ((dynamic)shape).Select();
        }

        internal void ToggleObjectState(object app, string layerId, int? shapeId, bool visibility)
        {
            var shapes = shapeId.HasValue
                ? new[] { FindObject(layerId, shapeId.Value) }.Where(shape => shape != null).ToArray()
                : Members().Where(member => member.LayerId == layerId).Select(member => member.Shape).ToArray();
            if (shapes.Length == 0) return;
            var before = shapes.Select(shape => ReadStateTree(shape, visibility).ToArray()).ToArray();
            if (before.Any(tree => tree.Any(state => !state.Value.HasValue))) return;
            bool turnOn = visibility ? before.All(tree => tree[0].Value == 0) : before.Any(tree => tree[0].Value == 0);
            int value = turnOn ? -1 : 0;
            var attempted = new List<int>();
            dynamic application = app;
            application.StartNewUndoEntry();
            try
            {
                for (int index = 0; index < shapes.Length; index++)
                {
                    if ((before[index][0].Value != 0) == turnOn) continue;
                    attempted.Add(index);
                    WriteState(shapes[index], visibility, value);
                }
            }
            catch (Exception error)
            {
                var failures = new List<Exception> { error };
                foreach (int index in attempted)
                    foreach (var state in before[index])
                        Restore(() =>
                        {
                            if (ReadState(state.Shape, visibility) != state.Value)
                                WriteState(state.Shape, visibility, state.Value.Value);
                        }, failures);
                if (failures.Count > 1) throw new AggregateException("图层状态更新失败，且部分对象无法恢复。", failures);
                throw;
            }
            finally { application.StartNewUndoEntry(); }
        }

        private object FindObject(string layerId, int id)
        {
            foreach (var member in Members().Where(member => member.LayerId == layerId))
            {
                var match = FindObject(member.Shape, id);
                if (match != null) return match;
            }
            return null;
        }

        private static object FindObject(object root, int id)
        {
            if ((int)((dynamic)root).Id == id) return root;
            foreach (var child in Children(root))
            {
                var match = FindObject(child, id);
                if (match != null) return match;
            }
            return null;
        }

        private static IEnumerable<object> Children(object shapeObject)
        {
            dynamic shape = shapeObject;
            if ((int)shape.Type != 6) return Array.Empty<object>();
            dynamic children = shape.GroupItems;
            return Enumerable.Range(1, (int)children.Count).Select(index => (object)children[index])
                .OrderByDescending(child => (int)((dynamic)child).ZOrderPosition);
        }

        private static int? ReadState(object shape, bool visibility)
        {
            return visibility ? (int)((dynamic)shape).Visible : ShapeCapabilities.ReadNativeLock(shape);
        }

        private static IEnumerable<(object Shape, int? Value)> ReadStateTree(object shape, bool visibility)
        {
            yield return (shape, ReadState(shape, visibility));
            foreach (var child in Children(shape))
                foreach (var state in ReadStateTree(child, visibility)) yield return state;
        }

        private static void WriteState(object shapeObject, bool visibility, int value)
        {
            dynamic shape = shapeObject;
            if (visibility) shape.Visible = value;
            else shape.Locked = value;
            int? actual = ReadState(shapeObject, visibility);
            if (!actual.HasValue || (actual.Value != 0) != (value != 0))
                throw new InvalidOperationException("PowerPoint 未能应用对象的显示或锁定状态。");
        }

        internal string Add(object app, string name)
        {
            name = CheckedName(name);
            var added = new Layer(Guid.NewGuid().ToString("N").ToUpperInvariant(), name);
            Commit(app, new[] { added }.Concat(Stack).ToArray());
            return added.Id;
        }

        internal void Rename(object app, string id, string name)
        {
            name = CheckedName(name);
            Commit(app, Stack.Select(layer => layer.Id == id ? new Layer(id, name) : layer).ToArray());
        }

        internal void Reorder(object app, string id, int index)
        {
            var stack = Stack.ToList();
            Layer layer = stack.Single(item => item.Id == id);
            stack.Remove(layer);
            stack.Insert(Math.Max(0, Math.Min(index, stack.Count)), layer);
            Commit(app, stack.ToArray());
        }

        internal void Assign(object app, string id)
        {
            if (!Stack.Any(layer => layer.Id == id)) return;
            var selected = Selected();
            if (selected.Count == 0) throw new InvalidOperationException("请先在幻灯片中选中要移动的对象或整个组合。");
            Commit(app, Stack, member => selected.Contains(member.Id) ? id : member.LayerId);
        }

        internal bool MoveObject(object app, string sourceLayer, int shapeId, string destinationLayer, int? anchorId, bool after)
        {
            if (!Stack.Any(layer => layer.Id == destinationLayer)) return false;
            var members = Members();
            var source = members.FirstOrDefault(member => member.Id == shapeId && member.LayerId == sourceLayer);
            var anchor = members.FirstOrDefault(member => member.Id == anchorId && member.LayerId == destinationLayer);
            if (source == null || source == anchor || (anchorId.HasValue && anchor == null)) return false;
            Commit(app, Stack,
                assign: member => member == source ? destinationLayer : member.LayerId,
                arrange: list => list.Contains(source) ? InsertBeside(list, source, anchor, after) : list,
                snapshot: members);
            return true;
        }

        internal void ReorderGroupMember(object app, string layerId, int groupId, int shapeId, int anchorId, bool after)
        {
            object group = FindObject(layerId, groupId);
            if (group == null || shapeId == anchorId) return;
            var children = Children(group).ToArray();
            var source = children.FirstOrDefault(shape => (int)((dynamic)shape).Id == shapeId);
            var anchor = children.FirstOrDefault(shape => (int)((dynamic)shape).Id == anchorId);
            if (source == null || anchor == null) return;
            var desired = InsertBeside(children, source, anchor, after).ToArray();
            if (children.SequenceEqual(desired)) return;
            dynamic application = app;
            application.StartNewUndoEntry();
            try { SetShapeOrder(desired); }
            catch (Exception error)
            {
                var failures = new List<Exception> { error };
                Restore(() => SetShapeOrder(children), failures);
                if (failures.Count > 1) throw new AggregateException("组合成员排序失败，且无法恢复原叠放顺序。", failures);
                throw;
            }
            finally { application.StartNewUndoEntry(); }
        }

        private static IEnumerable<T> InsertBeside<T>(IEnumerable<T> items, T source, T anchor, bool after) where T : class
        {
            var order = items.ToList();
            order.Remove(source);
            order.Insert(anchor == null ? 0 : order.IndexOf(anchor) + (after ? 1 : 0), source);
            return order;
        }

        internal void Remove(object app, string id)
        {
            int index = Array.FindIndex(Stack, layer => layer.Id == id);
            if (Stack.Length <= 1 || index < 0) return;
            var next = Stack.Where(layer => layer.Id != id).ToArray();
            string mergeInto = next[Math.Min(index, next.Length - 1)].Id;
            Commit(app, next, member => member.LayerId == id ? mergeInto : member.LayerId);
        }

        internal bool Arrange(object app, string command)
        {
            var selected = Selected();
            int operation = Array.IndexOf(ArrangeCommands, command);
            if (!Configured || Stack.Length < 2 || selected.Count == 0 || operation < 0) return false;
            Commit(app, Stack, arrange: layer =>
            {
                var items = layer.ToList();
                if (operation == 0) return items.OrderByDescending(item => selected.Contains(item.Id));
                if (operation == 3) return items.OrderBy(item => selected.Contains(item.Id));
                for (int step = 1; step < items.Count; step++)
                {
                    int i = operation == 1 ? step : items.Count - step - 1;
                    int neighbor = operation == 1 ? i - 1 : i + 1;
                    if (!selected.Contains(items[i].Id) || selected.Contains(items[neighbor].Id)) continue;
                    Member move = items[i];
                    items[i] = items[neighbor];
                    items[neighbor] = move;
                }
                return items;
            });
            return true;
        }

        internal bool EnsureOrder()
        {
            if (Stack.Length < 2) return false;
            var members = Members();
            var desired = SortIntoLayers(members);
            if (members.SequenceEqual(desired)) return false;
            try { SetOrder(desired); }
            catch (Exception error)
            {
                var failures = new List<Exception> { error };
                Restore(() => SetOrder(members), failures);
                if (failures.Count > 1) throw new AggregateException("图层顺序修正失败，且无法恢复原叠放顺序。", failures);
                throw;
            }
            return true;
        }

        internal Member[] SortIntoLayers(Member[] members)
        {
            var grouped = members.ToLookup(member => member.LayerId);
            return Stack.SelectMany(layer => grouped[layer.Id]).ToArray();
        }

        internal void ReapplyNativeOrder(object app, int[] rawOrder, int[] correctedOrder)
        {
            var baseline = Members();
            var byId = baseline.ToDictionary(member => member.Id);
            var corrected = correctedOrder.Select(id => byId[id]).ToArray();
            dynamic application = app;
            application.StartNewUndoEntry();
            try
            {
                // A cross-layer-only gesture can leave the same final order.
                // The replacement must still consume the discarded native
                // redo entry; otherwise Redo would resurrect the raw order.
                if (baseline.SequenceEqual(corrected)) SetOrder(rawOrder.Select(id => byId[id]).ToArray());
                SetOrder(corrected);
            }
            catch (Exception error)
            {
                var failures = new List<Exception> { error };
                Restore(() => SetOrder(baseline), failures);
                if (failures.Count > 1) throw new AggregateException("图层排序失败，且无法恢复原叠放顺序。", failures);
                throw;
            }
            finally { application.StartNewUndoEntry(); }
        }

        private void Commit(object app, Layer[] next, Func<Member, string> assign = null,
            Func<IEnumerable<Member>, IEnumerable<Member>> arrange = null, Member[] snapshot = null)
        {
            var members = snapshot ?? Members();
            var layerByMember = members.ToDictionary(member => member.Id, member => assign == null ? member.LayerId : assign(member));
            var grouped = members.ToLookup(member => layerByMember[member.Id]);
            var desired = next.SelectMany(layer => arrange == null ? grouped[layer.Id] : arrange(grouped[layer.Id])).ToArray();
            string nextScope = scope ?? Guid.NewGuid().ToString("N").ToUpperInvariant();
            string encoded = Stack.Select(layer => Tuple.Create(layer.Id, layer.Name)).SequenceEqual(next.Select(layer => Tuple.Create(layer.Id, layer.Name)))
                ? Encoded : NativeConstraintStore.EncodeXml(new XElement("layers", new XAttribute("scope", nextScope),
                    next.Select(layer => new XElement("layer", new XAttribute("id", layer.Id), new XAttribute("name", layer.Name)))).ToString(SaveOptions.DisableFormatting));
            var changed = new List<MembershipChange>();
            foreach (var member in members)
            {
                string layerId = layerByMember[member.Id];
                string prefix = nextScope + ":" + layerId;
                if (member.LayerId != layerId || (Stack[0].Id != next[0].Id && member.StoredTag != prefix + member.Owner))
                    PlanMembership(changed, member.Shape, member.StoredTag, member.Owner, prefix);
            }
            bool reordered = !members.SequenceEqual(desired);
            if (changed.Count == 0 && encoded == Encoded && !reordered) return;

            dynamic application = app;
            var written = new List<MembershipChange>();
            bool stackWritten = false, orderStarted = false;
            application.StartNewUndoEntry();
            try
            {
                foreach (var change in changed)
                {
                    written.Add(change);
                    WriteTag(change.Shape, MemberTag, change.Value);
                }
                if (encoded != Encoded)
                {
                    stackWritten = true;
                    WriteTag(Slide, StackTag, encoded);
                }
                if (reordered)
                {
                    orderStarted = true;
                    SetOrder(desired);
                }
            }
            catch (Exception error)
            {
                var failures = new List<Exception> { error };
                if (orderStarted) Restore(() => SetOrder(members), failures);
                if (stackWritten) Restore(() => WriteTag(Slide, StackTag, Encoded), failures);
                foreach (var change in written) Restore(() => WriteTag(change.Shape, MemberTag, change.Before), failures);
                if (failures.Count > 1) throw new AggregateException("图层更新失败，且部分对象无法恢复。", failures);
                throw;
            }
            finally { application.StartNewUndoEntry(); }
        }

        private static void PlanMembership(List<MembershipChange> changes, object shapeObject, string before, string owner, string prefix)
        {
            string value = prefix + owner;
            if (before != value) changes.Add(new MembershipChange { Shape = shapeObject, Before = before, Value = value });
            dynamic shape = shapeObject;
            if ((int)shape.Type != 6) return;
            dynamic children = shape.GroupItems;
            for (int i = 1; i <= (int)children.Count; i++)
            {
                dynamic child = children[i];
                PlanMembership(changes, (object)child, (string)child.Tags[MemberTag], ShapeSuffix((int)child.Id), prefix);
            }
        }

        private sealed class MembershipChange
        {
            public object Shape;
            public string Before, Value;
        }

        private static void SetOrder(Member[] frontToBack)
        { SetShapeOrder(frontToBack.Select(member => member.Shape)); }

        private static void SetShapeOrder(IEnumerable<object> frontToBack)
        {
            var backToFront = frontToBack.Reverse().ToArray();
            var positions = backToFront.Select(shape => (int)((dynamic)shape).ZOrderPosition).ToArray();
            int firstMove = 1;
            while (firstMove < positions.Length && positions[firstMove] > positions[firstMove - 1]) firstMove++;
            for (int i = firstMove; i < backToFront.Length; i++) ((dynamic)backToFront[i]).ZOrder(0);
            int lastPosition = 0;
            foreach (var shape in backToFront)
            {
                int position = (int)((dynamic)shape).ZOrderPosition;
                if (position <= lastPosition)
                    throw new InvalidOperationException("PowerPoint 未能应用图层叠放顺序。");
                lastPosition = position;
            }
        }

        private static void WriteTag(object owner, string tag, string value)
        {
            dynamic tags = ((dynamic)owner).Tags;
            if (value.Length == 0) tags.Delete(tag);
            else tags.Add(tag, value);
            if ((string)tags[tag] != value) throw new InvalidOperationException("PowerPoint 未能完整保存图层数据。");
        }

        private static void Restore(Action action, List<Exception> failures)
        {
            try { action(); }
            catch (Exception error) { failures.Add(error); }
        }

        private static string CheckedName(string name)
        {
            name = (name ?? "").Trim();
            if (name.Length == 0 || name.Length > 40 || name.Any(Char.IsControl))
                throw new InvalidOperationException("请输入不超过 40 个字符的图层名称。");
            return name;
        }
    }
}
