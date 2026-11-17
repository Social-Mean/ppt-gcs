[CmdletBinding()]
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if(-not $OutputDirectory){$OutputDirectory=Join-Path $root 'artifacts\ribbon-ui'}
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$temporary=Join-Path ([IO.Path]::GetTempPath()) ('PptGcsGridRender-'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporary | Out-Null
foreach($name in @('PptGcs.AddIn.dll','PptGcs.Core.dll','MathNet.Numerics.dll')){
    Copy-Item -LiteralPath (Join-Path $root ('bin\'+$name)) -Destination $temporary
}
$source=@'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class ConstraintGridRenderer
{
    [STAThread] private static void Main(string[] arguments)
    {
        // No Form, Office COM object, or input injection is created. Images are
        // real Office exports already captured by the separate icon probe.
        Application.EnableVisualStyles();
        var images=new Dictionary<string,Image>();
        try
        {
            // Bind historical captures to their actual Office resource IDs.
            // A constraint may now use a different glyph than when captured.
            var samples=new Dictionary<string,string> {
                {"ObjectsAlignLeftSmart","SameX"}, {"TableColumnsDistribute","SameWidth"},
                {"AlignDistributeHorizontally","EqualHorizontalGap"}
            };
            foreach(var sample in samples)
                images.Add(sample.Key,Image.FromFile(Path.Combine(arguments[0],
                    "artifacts/ui-refinement/icons-before/"+sample.Value+"-16-straight.png")));
            images.Add("DiagramTargetInsertClassic",Image.FromFile(Path.Combine(arguments[0],
                "artifacts/center-icon-correction/diagram-target-16.png")));
            foreach(int width in new[]{380,250}) Render(width,arguments[1],images);
        }
        finally {foreach(var image in images.Values)image.Dispose();}
    }

    private static void Render(int width,string directory,Dictionary<string,Image> images)
    {
        using(var pane=new ConstraintPaneControl())
        {
            pane.Size=new Size(width,1100);
            pane.CreateControl();
            pane.Bind(new ConstraintPaneActions {Add=(kind,value,second)=>{}});
            pane.SetIcons((kind,size)=>{
                Image image;
                string id=OfficeIconProvider.ImageMso(kind);
                return id!=null && images.TryGetValue(id,out image) && image.Width==size?image:null;
            });
            var geometry=new List<ShapeGeometry> {
                new ShapeGeometry("AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA",100,150,120,80),
                new ShapeGeometry("BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB",400,150,120,80)
            };
            pane.Render(new ConstraintPaneModel {
                PageLabel="组件布局检查",
                Objects=new List<ConstraintPaneObject> {
                    new ConstraintPaneObject {Id="A",Name="三角形 A",Selected=true,RelationCount=1},
                    new ConstraintPaneObject {Id="B",Name="文本框 B",Selected=true,RelationCount=1}
                },
                Relations=new List<ConstraintPaneRelation> {
                    new ConstraintPaneRelation {Id="R",Kind="SameWidth",Label="同宽",MemberIds=new[]{"A","B"},
                        MemberNames=new[]{"三角形 A","文本框 B"},Status="已满足"}
                },
                Options=ConstraintCatalog.Options(geometry)
            });
            pane.ShowEditor("FixedHorizontalGap");
            pane.AutoScrollPosition=Point.Empty;
            pane.PerformLayout();
            Application.DoEvents();
            var grid=Descendants(pane).OfType<TableLayoutPanel>().Single(c=>c.AccessibleName=="约束类型按钮");
            var viewport=Descendants(pane).OfType<Panel>().Single(c=>c.AccessibleName=="约束类型");
            var editor=Descendants(pane).OfType<GroupBox>().Single();
            var fields=editor.Controls.OfType<TableLayoutPanel>().Single();
            var active=grid.Controls.OfType<Button>().Single(c=>(string)c.Tag=="FixedHorizontalGap");
            if(grid.Controls.Count!=20 || grid.RowCount!=5 || grid.ColumnCount!=4)
                throw new InvalidOperationException("Expected four columns and five rows.");
            var choices=ConstraintCatalog.Options(geometry);
            for(int index=0;index<choices.Count;index++)
            {
                var button=grid.Controls.OfType<Button>().Single(c=>(string)c.Tag==choices[index].Kind);
                var cell=grid.GetPositionFromControl(button);
                if(cell.Column!=index%4 || cell.Row!=index/4)
                    throw new InvalidOperationException("Choices must fill four columns from left to right.");
            }
            if(fields.Bottom+editor.Padding.Bottom>editor.Height)
                throw new InvalidOperationException("Editor clips its parameter or apply controls.");
            if(active.Image==null || active.FlatAppearance.BorderSize!=2)
                throw new InvalidOperationException("The selected choice must show a real Office icon and selection border.");
            Capture(pane,Path.Combine(directory,"grid-final-"+width+".png"));
            if(width==250)
            {
                if(!viewport.HorizontalScroll.Visible)throw new InvalidOperationException("Narrow grid must scroll horizontally.");
                viewport.AutoScrollPosition=new Point(grid.Width,0);
                Application.DoEvents();
                Capture(pane,Path.Combine(directory,"grid-final-250-right.png"));
            }
            Console.WriteLine("PASS offscreen grid "+width+": 20 choices, four columns, full editor, real selected Office glyph; horizontal="+viewport.HorizontalScroll.Visible);
        }
    }

    private static void Capture(Control control,string path)
    {
        using(var bitmap=new Bitmap(control.Width,control.Height))
        {control.DrawToBitmap(bitmap,control.ClientRectangle);bitmap.Save(path,ImageFormat.Png);}
    }
    private static IEnumerable<Control> Descendants(Control parent)
    {
        foreach(Control child in parent.Controls)
        {yield return child;foreach(var item in Descendants(child))yield return item;}
    }
}
'@
$sourcePath=Join-Path $temporary 'ConstraintGridRenderer.cs'
[IO.File]::WriteAllText($sourcePath,$source,[Text.UTF8Encoding]::new($true))
$vswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$msbuild=& $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
if(-not $msbuild){throw 'Visual Studio MSBuild/Roslyn is required.'}
$compiler=Join-Path (Split-Path -Parent $msbuild) 'Roslyn\csc.exe'
$program=Join-Path $temporary 'ConstraintGridRenderer.exe'
& $compiler /nologo /target:exe ("/out:$program") /r:System.dll /r:System.Core.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll ("/r:$temporary\PptGcs.AddIn.dll") ("/r:$temporary\PptGcs.Core.dll") $sourcePath
if($LASTEXITCODE -ne 0){throw 'Constraint grid renderer compilation failed.'}
& $program $root $OutputDirectory
if($LASTEXITCODE -ne 0){throw 'Constraint grid offscreen verification failed.'}
