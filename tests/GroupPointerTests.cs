using System;
using System.Drawing;
using PptGcs.AddIn;

internal static class GroupPointerTests
{
    private static int passed;
    private static readonly RectangleF Bounds=new RectangleF(100,100,240,120);
    private static void Main()
    {
        // Called only after the actual RangeFromPoint top-level hit has been
        // checked against the constrained shape. A real group hit may lie
        // beyond the COM-reported box and must still own a whole-group drag.
        Check(GroupPointerPolicy.AcceptNativeHit(6,true),"a real group body hit starts a whole-group gesture");
        var outside=new PointF(180,233);
        Check(!Border(6,true,Bounds,outside,6) && GroupPointerPolicy.AcceptNativeHit(6,false),
            "a real group hit 13 pixels below its reported box is not dropped by the outline tolerance");
        Check(GroupPointerPolicy.AcceptNativeHit(6,false),"an unselected group also accepts its exact native hit beyond the reported box");
        foreach(int type in new[]{1,13,17})
        {
            Check(GroupPointerPolicy.AcceptNativeHit(type,true),"an existing body or border hit stays eligible for type "+type);
            Check(!GroupPointerPolicy.AcceptNativeHit(type,false),"group-specific native-hit tolerance does not extend type "+type);
            Check(!Border(type,true,Bounds,new PointF(98,140),6),"non-groups do not acquire a group selection outline for type "+type);
        }

        // With no exact native hit, only the selected group's narrow outline
        // can fill a blank hit. The controller separately rejects foreign hits.
        foreach(var point in new[]{new PointF(97,140),new PointF(343,140),new PointF(180,97),new PointF(180,223)})
            Check(Border(6,true,Bounds,point,6),"all four sides of a selected group can begin a whole-group move at "+point);
        Check(Border(6,true,Bounds,new PointF(94,140),6),"the exact outer outline tolerance is included");
        Check(!Border(6,true,Bounds,new PointF(93,140),6),"blank canvas beyond the outline is excluded");
        Check(!Border(6,false,Bounds,new PointF(98,140),6),"an unselected group does not guess an external outline");
        Check(Border(6,true,Bounds,new PointF(106,140),6),"the exact inner boundary still belongs to the outline");
        Check(!Border(6,true,Bounds,new PointF(106.1f,140),6),"the group interior is not treated as a guessed outline");
        foreach(float scale in new[]{1.25f,1.5f,2f})
        {
            var box=new RectangleF(Bounds.X*scale,Bounds.Y*scale,Bounds.Width*scale,Bounds.Height*scale);
            Check(Border(6,true,box,new PointF(97*scale,140*scale),6*scale) &&
                !Border(6,true,box,new PointF(93*scale,140*scale),6*scale),
                "the limited blank-hit outline follows display DPI at "+scale);
        }
        Check(!Border(6,true,new RectangleF(100,100,8,8),new PointF(104,104),6) &&
            GroupPointerPolicy.AcceptNativeHit(6,true),"a small group's native body hit stays draggable without expanding its guessed outline");

        // Selection is committed only after a click has completed. Passing the
        // threshold, even after returning to the start or a rejected solve,
        // must never turn a drag into child selection.
        Check(SelectChild(true,true,true,false,false,true),"a completed click on an already selected group's exact child enters that child");
        Check(!SelectChild(false,true,true,false,false,true),"the first click on an unselected group selects its parent");
        Check(!SelectChild(true,false,true,false,false,true),"a group gap or parent-only native hit never guesses a child");
        Check(!SelectChild(true,true,false,false,false,true),"resize input cannot select a child");
        Check(!SelectChild(true,true,true,true,false,true),"a whole-group body drag does not finish by selecting a child");
        Check(!SelectChild(true,true,true,true,true,true),"a cancelled drag never selects its original hit child");
        Check(!SelectChild(true,true,true,false,true,true),"Escape before the drag threshold does not commit child selection");
        Check(!SelectChild(true,true,true,false,false,false),"focus or capture loss does not commit child selection");
        Check(!SelectChild(true,true,true,false,true,false),"a cancelled click with no normal release does not change the selection");

        // A release can be delivered without an intervening motion callback.
        // It must use the same distance test before taking the click path.
        Check(!GroupPointerPolicy.HasDragDistance(0,0,4),"a stationary release remains a click");
        Check(!GroupPointerPolicy.HasDragDistance(3.99,-3.99,4),"sub-threshold jitter on both axes remains a click");
        foreach(double dx in new[]{4.0,-4.0})
            Check(GroupPointerPolicy.HasDragDistance(dx,0,4),"the exact horizontal drag threshold is included at "+dx);
        foreach(double dy in new[]{4.0,-4.0})
            Check(GroupPointerPolicy.HasDragDistance(0,dy,4),"the exact vertical drag threshold is included at "+dy);
        Check(GroupPointerPolicy.HasDragDistance(4.25,0,4),"a release point crossing the horizontal threshold cannot commit a click");
        Check(GroupPointerPolicy.HasDragDistance(0,-4.25,4),"a release point crossing the negative vertical threshold cannot commit a click");
        foreach(double scale in new[]{1.25,1.5,2.0})
        {
            double threshold=4*scale;
            Check(!GroupPointerPolicy.HasDragDistance(threshold-0.01,-threshold+0.01,threshold),
                "jitter below the DPI-scaled threshold remains a click at "+scale);
            Check(GroupPointerPolicy.HasDragDistance(-threshold,0,threshold),
                "the exact DPI-scaled threshold starts a drag at "+scale);
        }
        Console.WriteLine("PASS: "+passed+" group pointer tests");
    }
    private static bool Border(int type,bool selected,RectangleF bounds,PointF point,float border)
    {return GroupPointerPolicy.HitMoveBorder(type,selected,bounds,point,border);}
    private static bool SelectChild(bool selected,bool exact,bool move,bool dragged,bool cancelled,bool released)
    {return GroupPointerPolicy.ShouldSelectChild(selected,exact,move,dragged,cancelled,released);}
    private static void Check(bool value,string message)
    {if(!value)throw new InvalidOperationException(message);passed++;Console.WriteLine("PASS "+message);}
}
