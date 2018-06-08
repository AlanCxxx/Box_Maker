Option Explicit

' Maximize the window, close any existing drawing without saving and start a new drawing.
dcSetDrawingWindowMode dcMaximizeWin
dcCloseWithoutSaving
dcNew ""

' Drawing units and scale
dcSetDrawingScale 1
dcSetDrawingUnits dcMillimeters

Call Main
End


Sub Main

' Units are mm
Dim BoxSheetThickness As Double
Dim BoxFlangeWidth As Double
Dim FlangeHoleDiameter As Double
Dim OverCutWidth As Double
Dim EndMillDiameter As Double

' Inside dimensions
Dim BoxDepth As Double
Dim BoxWidth As Double
Dim BoxHeight As Double

BoxDepth=25.0
BoxWidth=25.0
BoxHeight=24.0
FlangeHoleDiameter=3.0
BoxSheetThickness=3.0
BoxFlangeWidth=2*BoxSheetThickness+FlangeHoleDiameter
OverCutWidth=0.0
EndMillDiameter=BoxSheetThickness/2

Begin Dialog DesignInput 50,50,160,110, "Input data for the Box"

  Text 10,10,80,10, "Inside box depth"
  Text 10,20,80,10, "Inside box width"
  Text 10,30,80,10, "Inside box height"
  Text 10,40,80,10, "Box sheet thickness"
  Text 10,50,80,10, "Box flange width"
  Text 10,60,80,10, "Flange hole diameter"
  Text 10,70,80,10, "Over (under) cut width"
  Text 10,80,80,10, "End Mill Diameter"
  TextBox 90,10,45,10, .bd
  TextBox 90,20,45,10, .bw
  TextBox 90,30,45,10, .bh
  TextBox 90,40,45,10, .st
  TextBox 90,50,45,10, .fw
  TextBox 90,60,45,10, .hd
  TextBox 90,70,45,10, .cw
  TextBox 90,80,45,10, .ed
  Text 140,10,15,10, "mm"
  Text 140,20,15,10, "mm"
  Text 140,30,15,10, "mm"
  Text 140,40,15,10, "mm"
  Text 140,50,15,10, "mm"
  Text 140,60,15,10, "mm"
  Text 140,70,15,10, "mm"
  Text 140,80,15,10, "mm"

  OKButton 10,95,30,10
  CancelButton 90,95,30,10

End Dialog

'Initialize 
Dim DesignData As DesignInput
Dim button As Long

DesignData.bd=BoxDepth
DesignData.bw=BoxWidth
DesignData.bh=BoxHeight
DesignData.st=BoxSheetThickness
DesignData.fw=BoxFlangeWidth
DesignData.hd=FlangeHoleDiameter
DesignData.cw=OverCutWidth
DesignData.ed=EndMillDiameter

button=Dialog(DesignData)

If (button=True) Then 

  BoxDepth=CDbl(DesignData.bd)
  BoxWidth=CDbl(DesignData.bw)
  BoxHeight=CDbl(DesignData.bh)
  BoxSheetThickness = CDbl(DesignData.st)
  BoxFlangeWidth=CDbl(DesignData.fw)
  FlangeHoleDiameter=CDbl(DesignData.hd)
  OverCutWidth=CDbl(DesignData.cw)
  EndMillDiameter=CDbl(DesignData.ed)
  
  Dim i As Long
  Dim XOffset As Double
  Dim YOffset As Double

  dcSetLineParms dcBLACK, dcSOLID, dcTHIN
  dcSetCircleParms dcBLACK, dcSOLID, dcTHIN

  XOffset=5.0  
  YOffset=5.0  

  ' Top and bottom
  For i=1 to 2

    dcSetCircleParms dcBLUE, dcSOLID, dcTHIN

    ' Outside (outline)
    dcCreateLine _
      XOffset-OverCutWidth/2, _
      YOffset-OverCutWidth/2+BoxSheetThickness/2, _
      XOffset-OverCutWidth/2+BoxSheetThickness/2, _
      YOffset-OverCutWidth/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness/2, _
      YOffset-OverCutWidth/2, _
      XOffset+OverCutWidth/2+BoxWidth+2*BoxFlangeWidth-BoxSheetThickness/2, _
      YOffset-OverCutWidth/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxWidth+2*BoxFlangeWidth-BoxSheetThickness/2, _
      YOffset-OverCutWidth/2, _
      XOffset+OverCutWidth/2+BoxWidth+2*BoxFlangeWidth, _
      YOffset-OverCutWidth/2+BoxSheetThickness/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxWidth+2*BoxFlangeWidth, _
      YOffset-OverCutWidth/2+BoxSheetThickness/2, _
      XOffset+OverCutWidth/2+BoxWidth+2*BoxFlangeWidth, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth-BoxSheetThickness/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxWidth+2*BoxFlangeWidth, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth-BoxSheetThickness/2, _
      XOffset+OverCutWidth/2+BoxWidth+2*BoxFlangeWidth-BoxSheetThickness/2, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxWidth+2*BoxFlangeWidth-BoxSheetThickness/2, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth, _
      XOffset-OverCutWidth/2+BoxSheetThickness/2, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness/2, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth, _
      XOffset-OverCutWidth/2, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth-BoxSheetThickness/2

    dcCreateLine _
      XOffset-OverCutWidth/2, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth-BoxSheetThickness/2, _
      XOffset-OverCutWidth/2, _
      YOffset-OverCutWidth/2+BoxSheetThickness/2

    ' West slot
    dcCreateLine _
      XOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxFlangeWidth+BoxDepth/4, _
      XOffset-OverCutWidth/2+BoxFlangeWidth, _
      YOffset+OverCutWidth/2+BoxFlangeWidth+BoxDepth/4

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth+BoxDepth/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxFlangeWidth, _
      YOffset+OverCutWidth/2+BoxFlangeWidth+BoxDepth/4, _
      XOffset-OverCutWidth/2+BoxFlangeWidth, _
      YOffset-OverCutWidth/2+BoxFlangeWidth+BoxDepth*3/4

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxFlangeWidth, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth+BoxDepth/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxFlangeWidth, _
      YOffset-OverCutWidth/2+BoxFlangeWidth+BoxDepth*3/4, _
      XOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness, _
      YOffset-OverCutWidth/2+BoxFlangeWidth+BoxDepth*3/4

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxFlangeWidth, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth+BoxDepth*3/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness, _
      YOffset-OverCutWidth/2+BoxFlangeWidth+BoxDepth*3/4, _
      XOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxFlangeWidth+BoxDepth/4

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth+BoxDepth*3/4, _
      EndMillDiameter/2

    ' East slot
    dcCreateLine _
      XOffset-OverCutWidth/2+BoxWidth+BoxFlangeWidth+BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxFlangeWidth+BoxDepth/4, _
      XOffset+OverCutWidth/2+BoxWidth+BoxFlangeWidth, _
      YOffset+OverCutWidth/2+BoxFlangeWidth+BoxDepth/4

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxWidth+BoxFlangeWidth+BoxSheetThickness, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth+BoxDepth/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxWidth+BoxFlangeWidth, _
      YOffset+OverCutWidth/2+BoxFlangeWidth+BoxDepth/4, _
      XOffset+OverCutWidth/2+BoxWidth+BoxFlangeWidth, _
      YOffset-OverCutWidth/2+BoxFlangeWidth+BoxDepth*3/4

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxWidth+BoxFlangeWidth, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth+BoxDepth/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxWidth+BoxFlangeWidth, _
      YOffset-OverCutWidth/2+BoxFlangeWidth+BoxDepth*3/4, _
      XOffset-OverCutWidth/2+BoxWidth+BoxFlangeWidth+BoxSheetThickness, _
      YOffset-OverCutWidth/2+BoxFlangeWidth+BoxDepth*3/4

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxWidth+BoxFlangeWidth, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth+BoxDepth*3/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxWidth+BoxFlangeWidth+BoxSheetThickness, _
      YOffset-OverCutWidth/2+BoxFlangeWidth+BoxDepth*3/4, _
      XOffset-OverCutWidth/2+BoxWidth+BoxFlangeWidth+BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxFlangeWidth+BoxDepth/4

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxWidth+BoxFlangeWidth+BoxSheetThickness, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth+BoxDepth*3/4, _
      EndMillDiameter/2

    ' South slot
    dcCreateLine _
      XOffset+OverCutWidth/2+BoxFlangeWidth+BoxWidth/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness, _
      XOffset+OverCutWidth/2+BoxFlangeWidth+BoxWidth/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxFlangeWidth+BoxWidth/4, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxFlangeWidth+BoxWidth/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth, _
      XOffset-OverCutWidth/2+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxFlangeWidth+BoxWidth/4, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth, _
      XOffset-OverCutWidth/2+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness, _
      XOffset+OverCutWidth/2+BoxFlangeWidth+BoxWidth/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness, _
      EndMillDiameter/2

    ' North slot
    dcCreateLine _
      XOffset+OverCutWidth/2+BoxFlangeWidth+BoxWidth/4, _
      YOffset-OverCutWidth/2+BoxDepth+BoxFlangeWidth+BoxSheetThickness, _
      XOffset+OverCutWidth/2+BoxFlangeWidth+BoxWidth/4, _
      YOffset+OverCutWidth/2+BoxDepth+BoxFlangeWidth

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxFlangeWidth+BoxWidth/4, _
      YOffset-0.3536*EndMillDiameter+BoxDepth+BoxFlangeWidth+BoxSheetThickness, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxFlangeWidth+BoxWidth/4, _
      YOffset+OverCutWidth/2+BoxDepth+BoxFlangeWidth, _
      XOffset-OverCutWidth/2+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset+OverCutWidth/2+BoxDepth+BoxFlangeWidth

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxFlangeWidth+BoxWidth/4, _
      YOffset+0.3536*EndMillDiameter+BoxDepth+BoxFlangeWidth, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset+OverCutWidth/2+BoxDepth+BoxFlangeWidth, _
      XOffset-OverCutWidth/2+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset-OverCutWidth/2+BoxDepth+BoxFlangeWidth+BoxSheetThickness

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset+0.3536*EndMillDiameter+BoxDepth+BoxFlangeWidth, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset-OverCutWidth/2+BoxDepth+BoxFlangeWidth+BoxSheetThickness, _
      XOffset+OverCutWidth/2+BoxFlangeWidth+BoxWidth/4, _
      YOffset-OverCutWidth/2+BoxDepth+BoxFlangeWidth+BoxSheetThickness

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxFlangeWidth+BoxWidth*3/4, _
      YOffset-0.3536*EndMillDiameter+BoxDepth+BoxFlangeWidth+BoxSheetThickness, _
      EndMillDiameter/2


    dcSetCircleParms dcBLACK, dcSOLID, dcTHIN

    ' Screw holes
    dcCreateCircle _
      XOffset+BoxFlangeWidth-BoxSheetThickness/2, _
      YOffset+BoxFlangeWidth-BoxSheetThickness/2, _
      EndMillDiameter/2
    dcCreateCircle _
      XOffset+BoxFlangeWidth-BoxSheetThickness/2, _
      YOffset+BoxFlangeWidth-BoxSheetThickness/2+BoxSheetThickness+BoxDepth, _
      EndMillDiameter/2
    dcCreateCircle _
      XOffset+BoxFlangeWidth-BoxSheetThickness/2+BoxSheetThickness+BoxWidth, _
      YOffset+BoxFlangeWidth-BoxSheetThickness/2+BoxSheetThickness+BoxDepth, _
      EndMillDiameter/2
    dcCreateCircle _
      XOffset+BoxFlangeWidth-BoxSheetThickness/2+BoxSheetThickness+BoxWidth, _
      YOffset+BoxFlangeWidth-BoxSheetThickness/2, _
      EndMillDiameter/2


    ' Flange holes
    dcCreateCircle _
      XOffset+BoxFlangeWidth-BoxSheetThickness-FlangeHoleDiameter/2, _
      YOffset+BoxFlangeWidth-BoxSheetThickness-FlangeHoleDiameter/2, _
      FlangeHoleDiameter/2
    dcCreateCircle _
      XOffset+BoxFlangeWidth-BoxSheetThickness-FlangeHoleDiameter/2, _
      YOffset+BoxDepth+BoxFlangeWidth+BoxSheetThickness+FlangeHoleDiameter/2, _
      FlangeHoleDiameter/2-OverCutWidth/2
    dcCreateCircle _
      XOffset+BoxWidth+BoxFlangeWidth+BoxSheetThickness+FlangeHoleDiameter/2, _
      YOffset+BoxDepth+BoxFlangeWidth+BoxSheetThickness+FlangeHoleDiameter/2, _
      FlangeHoleDiameter/2-OverCutWidth/2
    dcCreateCircle _
      XOffset+BoxWidth+BoxFlangeWidth+BoxSheetThickness+FlangeHoleDiameter/2, _
      YOffset+BoxFlangeWidth-BoxSheetThickness-FlangeHoleDiameter/2, _
      FlangeHoleDiameter/2-OverCutWidth/2





    XOffset=XOffset+BoxWidth+2*BoxFlangeWidth+5.0

    dcSetCircleParms dcBLUE, dcSOLID, dcTHIN

    ' Side (start bottom left corner and clockwise) 
    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset-OverCutWidth/2, _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth/4

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth/4, _
      XOffset-OverCutWidth/2, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth/4

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxSheetThickness, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth/4, _
      XOffset-OverCutWidth/2, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth*3/4

    dcCreateLine _
      XOffset-OverCutWidth/2, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth*3/4, _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth*3/4

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth*3/4, _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxSheetThickness, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth*3/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxDepth+2*BoxFlangeWidth, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth*3/4

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth*3/4, _
      XOffset+OverCutWidth/2+2*BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth*3/4

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth*3/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+2*BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth*3/4, _
      XOffset+OverCutWidth/2+2*BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth/4

    dcCreateLine _
      XOffset+OverCutWidth/2+2*BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth/4, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth/4

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth/4, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2, _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset-OverCutWidth/2


    ' Side North Slot 
    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+2*BoxSheetThickness+BoxDepth, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight/4, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+2*BoxSheetThickness+BoxDepth, _
      EndMillDiameter/2
     
    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth, _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth
     
    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight/4, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth, _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+2*BoxSheetThickness+BoxDepth
 
    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxDepth, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+2*BoxSheetThickness+BoxDepth, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+2*BoxSheetThickness+BoxDepth
 
    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+2*BoxSheetThickness+BoxDepth, _
      EndMillDiameter/2

    ' Side South Slot
    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness
 
    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness
 
    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight/4, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness, _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness
 
    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight/4, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness, _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness
 
    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness, _
      EndMillDiameter/2



    YOffset=YOffset+BoxDepth+2*BoxFlangeWidth+5.0

    ' End 
    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness, _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth/4

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth/4, _
      XOffset-OverCutWidth/2, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth/4

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxSheetThickness, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth/4, _
      XOffset-OverCutWidth/2, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth*3/4

    dcCreateLine _
      XOffset-OverCutWidth/2, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth*3/4, _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth*3/4

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth*3/4, _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxSheetThickness, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth*3/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth, _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth
    
    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth, _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+2*BoxSheetThickness+BoxWidth

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight/4, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+2*BoxSheetThickness+BoxWidth, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+2*BoxSheetThickness+BoxWidth

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+2*BoxSheetThickness+BoxWidth, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight*3/4, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth*3/4

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth*3/4, _
      XOffset+OverCutWidth/2+2*BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth*3/4

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight, _
      YOffset+0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth*3/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+2*BoxSheetThickness+BoxHeight, _
      YOffset+OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth*3/4, _
      XOffset+OverCutWidth/2+2*BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth/4

    dcCreateLine _
      XOffset+OverCutWidth/2+2*BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth/4, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth/4

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth/4, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness+BoxWidth/4, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness, _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness

    dcCreateCircle _
      XOffset+0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-0.3536*EndMillDiameter+BoxSheetThickness+BoxFlangeWidth-BoxSheetThickness, _
      EndMillDiameter/2

    dcCreateLine _
      XOffset+OverCutWidth/2+BoxSheetThickness+BoxHeight*3/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness, _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness, _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness

    dcCreateLine _
      XOffset-OverCutWidth/2+BoxSheetThickness+BoxHeight/4, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness, _
      XOffset-OverCutWidth/2+BoxSheetThickness, _
      YOffset-OverCutWidth/2+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness

    dcCreateCircle _
      XOffset-0.3536*EndMillDiameter+BoxSheetThickness+BoxHeight/4, _
      YOffset-0.3536*EndMillDiameter+BoxFlangeWidth-BoxSheetThickness+BoxSheetThickness, _
      EndMillDiameter/2


    YOffset=YOffset-BoxDepth-2*BoxFlangeWidth-5.0
    XOffset=XOffset+BoxHeight+2*BoxSheetThickness+5.0
  
  Next i
End If

End Sub

