using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace PptGcs.AddIn
{
    internal sealed class ConstraintConflictControl : GroupBox
    {
        private readonly Label prompt = new Label();
        private readonly CheckedListBox choices = new CheckedListBox();
        private readonly Button retry = new Button();

        public ConstraintCommands.DraftConflict Draft { get; private set; }
        public event Action<IReadOnlyList<string>> Retry;
        public event Action Cancel;

        public ConstraintConflictControl()
        {
            Text="约束冲突";
            AccessibleName=Text;
            AutoSize=true;
            AutoSizeMode=AutoSizeMode.GrowAndShrink;
            Dock=DockStyle.Fill;
            Padding=new Padding(8,14,8,8);
            Visible=false;

            var content=new TableLayoutPanel {ColumnCount=1,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,Dock=DockStyle.Top};
            prompt.AutoSize=true;
            prompt.Dock=DockStyle.Fill;
            content.Controls.Add(prompt);
            choices.AccessibleName="可以停用的冲突关系";
            choices.CheckOnClick=true;
            choices.HorizontalScrollbar=true;
            choices.IntegralHeight=false;
            choices.Dock=DockStyle.Fill;
            choices.ItemCheck+=(sender,args)=>retry.Enabled=(choices.CheckedItems.Count+
                (args.NewValue==CheckState.Checked?1:0)-(args.CurrentValue==CheckState.Checked?1:0))>0;
            choices.KeyDown+=(sender,args)=>{
                if(args.KeyCode!=Keys.Enter || !retry.Enabled)return;
                args.SuppressKeyPress=true;
                Submit();
            };
            content.Controls.Add(choices);
            var buttons=new FlowLayoutPanel {AutoSize=true,Dock=DockStyle.Fill,Margin=new Padding(0,3,0,0)};
            retry.Text="停用所选并重试";
            retry.AutoSize=true;
            retry.Click+=(sender,args)=>Submit();
            buttons.Controls.Add(retry);
            var cancel=new Button {Text="取消本次提交",AutoSize=true};
            cancel.Click+=(sender,args)=>Cancel?.Invoke();
            buttons.Controls.Add(cancel);
            content.Controls.Add(buttons);
            Controls.Add(content);
            SizeChanged+=(sender,args)=>prompt.MaximumSize=new Size(Math.Max(150,ClientSize.Width-Padding.Horizontal-8),0);
            content.SizeChanged+=(sender,args)=>{
                int height=content.Bottom+Padding.Bottom;
                if(Height!=height)Height=height;
            };
        }

        public void ShowDraft(ConstraintCommands.DraftConflict conflict)
        {
            Draft=conflict;
            string title=conflict.Attempted==null?"现有约束不能同时满足。":"无法提交 "+conflict.Attempted;
            int staged=conflict.Plan.Document.Constraints.Count(item=>!item.Enabled &&
                conflict.Plan.Context.Document.Constraints.Any(old=>old.Id==item.Id && old.Enabled));
            prompt.Text=title+"\r\n选择要停用的关系，再重试本次编辑。"+(staged>0?"已选停用 "+staged+" 条。":String.Empty);
            choices.Items.Clear();
            foreach(var choice in conflict.Choices.Where(item=>item.CanDisable))choices.Items.Add(choice);
            choices.Height=Math.Max(48,Math.Min(120,choices.ItemHeight*(choices.Items.Count+1)));
            retry.Enabled=false;
            Visible=true;
            if(choices.Items.Count>0)choices.SelectedIndex=0;
            if(IsHandleCreated)BeginInvoke((Action)(()=>{if(Draft==conflict && Visible)choices.Focus();}));
        }

        public void ClearDraft()
        { Draft=null;Visible=false;choices.Items.Clear(); }

        private void Submit()
        {
            var selected=choices.CheckedItems.Cast<ConstraintCommands.ConflictRelation>().Select(item=>item.Id).ToList();
            if(Draft!=null && selected.Count>0)Retry?.Invoke(selected.AsReadOnly());
        }
    }
}
