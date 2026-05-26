# VBA Shape Patterns

Use these patterns as building blocks for editable Office reconstruction. Adapt `sld.Shapes` to `ws.Shapes` for Excel or `doc.Shapes` / `Selection.InlineShapes` patterns for Word as needed.

## Basic Setup

```vb
Option Explicit

Sub ReconstructImage()
    Dim pres As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim i As Long
    Const PREFIX As String = "AITVBA_"

    Set pres = ActivePresentation
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)

    For i = sld.Shapes.Count To 1 Step -1
        If Left$(sld.Shapes(i).Name, Len(PREFIX)) = PREFIX Then
            sld.Shapes(i).Delete
        End If
    Next i
End Sub
```

Use a new blank slide by default. If the user explicitly wants to update an existing slide, delete only generated shapes with the agreed prefix; do not delete unrelated user content.

## Add Rectangle

```vb
Set shp = sld.Shapes.AddShape(msoShapeRectangle, 40, 50, 220, 90)
shp.Name = "AITVBA_Panel_Rectangle"
shp.Fill.ForeColor.RGB = RGB(245, 247, 250)
shp.Line.Visible = msoFalse
```

## Add Rounded Rectangle

```vb
Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, 60, 70, 180, 56)
shp.Name = "AITVBA_Primary_Button"
shp.Adjustments.Item(1) = 0.18
shp.Fill.ForeColor.RGB = RGB(35, 101, 235)
shp.Line.Visible = msoFalse
```

## Add Ellipse

```vb
Set shp = sld.Shapes.AddShape(msoShapeOval, 100, 120, 72, 72)
shp.Name = "AITVBA_Avatar_Circle"
shp.Fill.ForeColor.RGB = RGB(255, 210, 90)
shp.Line.ForeColor.RGB = RGB(230, 170, 45)
shp.Line.Weight = 1.25
```

## Add Line

```vb
Set shp = sld.Shapes.AddLine(80, 220, 340, 220)
shp.Name = "AITVBA_Divider_Line"
shp.Line.ForeColor.RGB = RGB(210, 216, 224)
shp.Line.Weight = 1
shp.Line.DashStyle = msoLineSolid
```

## Add TextBox

```vb
Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 80, 250, 320, 36)
shp.Name = "AITVBA_Title_Text"
With shp.TextFrame2
    .TextRange.Text = "Editable title"
    .MarginLeft = 0
    .MarginRight = 0
    .MarginTop = 0
    .MarginBottom = 0
    .WordWrap = msoFalse
    With .TextRange.Font
        .Name = "Aptos Display"
        .Size = 24
        .Bold = msoTrue
        .Fill.ForeColor.RGB = RGB(24, 30, 42)
    End With
End With
shp.Line.Visible = msoFalse
shp.Fill.Visible = msoFalse
```

## Set Font

```vb
With shp.TextFrame2.TextRange.Font
    .Name = "Aptos"
    .Size = 12
    .Bold = msoFalse
    .Italic = msoFalse
    .Fill.ForeColor.RGB = RGB(72, 80, 96)
End With
```

## Set Fill Color

```vb
shp.Fill.Visible = msoTrue
shp.Fill.Solid
shp.Fill.ForeColor.RGB = RGB(248, 250, 252)
```

## Set Transparency

```vb
shp.Fill.Transparency = 0.35  ' 0 opaque, 1 fully transparent
```

## Set Border

```vb
With shp.Line
    .Visible = msoTrue
    .ForeColor.RGB = RGB(180, 190, 205)
    .Weight = 1.25
    .Transparency = 0.1
End With
```

## Set Shadow

```vb
With shp.Shadow
    .Visible = msoTrue
    .Style = msoShadowStyleOuterShadow
    .Blur = 12
    .OffsetX = 2
    .OffsetY = 3
    .Transparency = 0.65
    .ForeColor.RGB = RGB(0, 0, 0)
End With
```

## Set Gradient

```vb
With shp.Fill
    .TwoColorGradient msoGradientHorizontal, 1
    .ForeColor.RGB = RGB(44, 123, 229)
    .BackColor.RGB = RGB(117, 83, 255)
End With
```

For more precise gradients in hosts that support `GradientStops`:

```vb
With shp.Fill
    .ForeColor.RGB = RGB(44, 123, 229)
    .GradientStops.Insert RGB(117, 83, 255), 1
End With
```

## Set ZOrder

Creation order usually controls layers. Use explicit ordering when necessary:

```vb
shp.ZOrder msoSendToBack
shp.ZOrder msoBringToFront
```

## Insert Preserved Raster Crop

Use this only for hybrid preserved-element reconstruction. Keep the crop local and limited to the complex artwork; rebuild text, arrows, boxes, and callouts as editable shapes around it.

```vb
Dim assetPath As String
assetPath = ThisPresentation.Path & Application.PathSeparator & "assets" & Application.PathSeparator & "01_complex_rendering.png"

Set shp = sld.Shapes.AddPicture( _
    FileName:=assetPath, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=240, Top:=90, Width:=260, Height:=180)
shp.Name = "AITVBA_Raster_Complex_Rendering_01"
shp.Line.Visible = msoFalse
```

If the presentation has not been saved yet, use an absolute asset path generated alongside the `.bas` file. After the deck is saved, prefer relative paths next to the deck.

## FreeformBuilder Example

```vb
Dim ff As FreeformBuilder
Set ff = sld.Shapes.BuildFreeform(msoEditingCorner, 420, 140)
ff.AddNodes msoSegmentLine, msoEditingCorner, 470, 115
ff.AddNodes msoSegmentLine, msoEditingCorner, 525, 152
ff.AddNodes msoSegmentLine, msoEditingCorner, 500, 205
ff.AddNodes msoSegmentLine, msoEditingCorner, 438, 194
Set shp = ff.ConvertToShape
shp.Name = "AITVBA_Custom_Polygon"
shp.Fill.ForeColor.RGB = RGB(255, 196, 80)
shp.Line.ForeColor.RGB = RGB(205, 142, 30)
shp.Line.Weight = 1
```

## Layered Raster-Like Approximation

For soft glows, blur-like highlights, or photo-like gradients, stack transparent shapes:

```vb
Set shp = sld.Shapes.AddShape(msoShapeOval, 300, 90, 260, 160)
shp.Name = "AITVBA_Soft_Highlight_Layer_1"
shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
shp.Fill.Transparency = 0.72
shp.Line.Visible = msoFalse
shp.SoftEdge.Radius = 20
```
