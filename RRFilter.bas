Attribute VB_Name = "Module2"
' Module2 RRFilter.bas

Option Base 1
DefLng A-W
DefSng X-Z

Public Sub ApplyFilter()

Form1.Shap.FillColor = QBColor(12)
Form1.Shap.Refresh


If ASM Then

' MCode Structure
'Public Type MCodeStruc
'   PICW As Long
'   PICH As Long
'   zmf1 As Single
'   zmf2 As Single
'   zdf As Single
'   zf As Single
'   ckQB As Long
'   ckABS As Long
'   QBLongColor As Long
'   PtrMat As Long
'   PtrPalBGR As Long
'End Type
'Public MC As MCodeStruc
MC.PICW = PICW
MC.PICH = PICH
MC.zmf1 = zmf1
MC.zmf2 = zmf2
MC.zdf = zdf
MC.zf = zf

If Form1.chkQB.Value = Unchecked Then
   MC.ckQB = 0
Else
   MC.ckQB = 1
End If

If Form1.chkAbs.Value = Unchecked Then
   MC.ckABS = 0
Else
   MC.ckABS = 1
End If

MC.QBLongColor = QBLongColor
MC.PtrMat = VarPtr(Mat(1, 1))
MC.PtrPalBGR = VarPtr(PalBGR(1, 1, 1, 1))

ptMC = VarPtr(IFilterMC(1))     ' Pointer to mcode
ptrStruc = VarPtr(MC.PICW)  ' Pointer to structure


res = CallWindowProc(ptMC, ptrStruc, 0&, 0&, 0&)


Else


' Determine PIC limits
PicXLo = 2: PicXHi = PICW - 1
PicYLo = 2: PicYHi = PICH - 1
limx = 1
limy = 1

For ix = 1 To 5
   If Mat(ix, 1) <> 0 Or Mat(ix, 5) <> 0 Then
      PicXLo = 3: PicXHi = PICW - 2
      limx = 2
      Exit For
   End If
Next ix

For iy = 1 To 5
   If Mat(1, iy) <> 0 Or Mat(5, iy) <> 0 Then
      PicYLo = 3: PicYHi = PICH - 2
      limy = 2
      Exit For
   End If
Next iy


For iy = PicYLo To PicYHi 'Step 1 + 3 * Rnd
For ix = PicXLo To PicXHi 'Step 1 + 3 * Rnd
         
   culB = 0
   culG = 0
   culR = 0
   
   For iyy = -limy To limy
   miy = iyy + 3
   For ixx = -limx To limx
      
      If ixx <> 0 Or iyy <> 0 Then
         mix = ixx + 3
         culB = culB + Mat(mix, miy) * PalBGR(1, ix + ixx, iy + iyy, 3)
         culG = culG + Mat(mix, miy) * PalBGR(2, ix + ixx, iy + iyy, 3)
         culR = culR + Mat(mix, miy) * PalBGR(3, ix + ixx, iy + iyy, 3)
      End If
      
   Next ixx
   Next iyy
      
   culB = (zmf1 * culB + zmf2 * Mat(3, 3) * PalBGR(1, ix, iy, 3)) / zdf
   culG = (zmf1 * culG + zmf2 * Mat(3, 3) * PalBGR(2, ix, iy, 3)) / zdf
   culR = (zmf1 * culR + zmf2 * Mat(3, 3) * PalBGR(3, ix, iy, 3)) / zdf
      
   If Form1.chkQB.Value = Checked Then
      culB = culB + QBBlue
      culG = culG + QBGreen
      culR = culR + QBRed
   Else
      culB = culB + zf
      culG = culG + zf
      culR = culR + zf
   End If
      
   If Form1.chkAbs.Value = Checked Then
      culB = Abs(culB)
      culG = Abs(culG)
      culR = Abs(culR)
   End If
      
   If culB > 255 Then culB = 255
   If culG > 255 Then culG = 255
   If culR > 255 Then culR = 255
      
   If culB < 0 Then culB = 0
   If culG < 0 Then culG = 0
   If culR < 0 Then culR = 0
      
   PalBGR(1, ix, iy, 2) = culB
   PalBGR(2, ix, iy, 2) = culG
   PalBGR(3, ix, iy, 2) = culR
      
Next ix
Next iy

End If
   
   '---------------------
   ShowPalBGR 2
   '---------------------


Form1.Shap.FillColor = QBColor(10)
Form1.Shap.Refresh

End Sub

Public Sub UR()

ReDim URName$(8)
ReDim zURFactors(6, 8), URMatrix(5, 5, 8)

URName$(1) = "None"
URName$(2) = "Average"
URName$(3) = "Edge Vertical"
URName$(4) = "Edge Horizontal"
URName$(5) = "Sobel Vertical"
URName$(6) = "Sobel Horizontal"
URName$(7) = "Laplacian Edge"
URName$(8) = "Laplacian Sharpen"

Form1.LisUnLimReal.Clear
For i = 1 To 8
   Form1.LisUnLimReal.AddItem URName$(i)
Next i

FillFactors zURFactors(), 1, "0,1,1,9,0,0"
FillFactors zURFactors(), 2, "0,1,1,9,0,0"
FillFactors zURFactors(), 3, "0,1,1,2,0,0"
FillFactors zURFactors(), 4, "0,1,1,2,0,0"
FillFactors zURFactors(), 5, "0,1,1,1,0,0"
FillFactors zURFactors(), 6, "0,1,1,1,0,0"
FillFactors zURFactors(), 7, "0,1,1,1,0,0"
FillFactors zURFactors(), 8, "0,1,1,1,0,0"

Item = 1 ' None
FillMatrix URMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 2, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 3, Item, "0,0,9,0,0"
FillMatrix URMatrix(), 4, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 5, Item, "0,0,0,0,0"

Item = 2 ' Average
FillMatrix URMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 2, Item, "0,1,1,1,0"
FillMatrix URMatrix(), 3, Item, "0,1,1,1,0"
FillMatrix URMatrix(), 4, Item, "0,1,1,1,0"
FillMatrix URMatrix(), 5, Item, "0,0,0,0,0"

Item = 3 ' Edge vertical
FillMatrix URMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 2, Item, "0,-1,0,1,0"
FillMatrix URMatrix(), 3, Item, "0,-1,0,1,0"
FillMatrix URMatrix(), 4, Item, "0,-1,0,1,0"
FillMatrix URMatrix(), 5, Item, "0,0,0,0,0"

Item = 4 ' Edge Horizontal
FillMatrix URMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 2, Item, "0,-1,-1,-1,0"
FillMatrix URMatrix(), 3, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 4, Item, "0,1,1,1,0"
FillMatrix URMatrix(), 5, Item, "0,0,0,0,0"

Item = 5 ' Sobel vertical
FillMatrix URMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 2, Item, "0,-1,0,1,0"
FillMatrix URMatrix(), 3, Item, "0,-2,0,2,0"
FillMatrix URMatrix(), 4, Item, "0,-1,0,1,0"
FillMatrix URMatrix(), 5, Item, "0,0,0,0,0"

Item = 6 ' Sobel horizontal
FillMatrix URMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 2, Item, "0,-1,-2,-1,0"
FillMatrix URMatrix(), 3, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 4, Item, "0,1,2,1,0"
FillMatrix URMatrix(), 5, Item, "0,0,0,0,0"

Item = 7 ' Laplacian edge
FillMatrix URMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 2, Item, "0,0,-1,0,0"
FillMatrix URMatrix(), 3, Item, "0,-1,4,-1,0"
FillMatrix URMatrix(), 4, Item, "0,0,-1,0,0"
FillMatrix URMatrix(), 5, Item, "0,0,0,0,0"

Item = 8 ' Laplacian sharpen
FillMatrix URMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix URMatrix(), 2, Item, "0,0,-1,0,0"
FillMatrix URMatrix(), 3, Item, "0,-1,5,-1,0"
FillMatrix URMatrix(), 4, Item, "0,0,-1,0,0"
FillMatrix URMatrix(), 5, Item, "0,0,0,0,0"

End Sub

Public Sub MS()

ReDim MSName$(8)
ReDim zMSFactors(6, 8), MSMatrix(5, 5, 8)

MSName$(1) = "Sharpen"
MSName$(2) = "Engrave "
MSName$(3) = "Engrave More"
MSName$(4) = "Emboss"
MSName$(5) = "Emboss More"
MSName$(6) = "Edge Enhance"
MSName$(7) = "Contour"
MSName$(8) = "Relief"

Form1.LisManSantos.Clear
For i = 1 To 8
   Form1.LisManSantos.AddItem MSName$(i)
Next i

'Abs, zmf1, zmf2, zdf, zf|QB ' factors
FillFactors zMSFactors(), 1, "0,-2,26,10,0,0"
FillFactors zMSFactors(), 2, "1,1,1,1,0,1"
FillFactors zMSFactors(), 3, "0,1,0,1,0,1"
FillFactors zMSFactors(), 4, "1,1,1,1,0,1"
FillFactors zMSFactors(), 5, "0,1,0,1,0,1"
FillFactors zMSFactors(), 6, "0,-1,10,2,0,0"
FillFactors zMSFactors(), 7, "0,-1,8,1,0,1"
FillFactors zMSFactors(), 8, "0,1,1,2,50,0"

Item = 1 ' Sharpen
FillMatrix MSMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix MSMatrix(), 2, Item, "0,1,1,1,0"
FillMatrix MSMatrix(), 3, Item, "0,1,1,1,0"
FillMatrix MSMatrix(), 4, Item, "0,1,1,1,0"
FillMatrix MSMatrix(), 5, Item, "0,0,0,0,0"

Item = 2 ' Engrave
FillMatrix MSMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix MSMatrix(), 2, Item, "0,0,0,1,0"
FillMatrix MSMatrix(), 3, Item, "0,0,-1,0,0"
FillMatrix MSMatrix(), 4, Item, "0,0,0,0,0"
FillMatrix MSMatrix(), 5, Item, "0,0,0,0,0"

Item = 3 ' Engrave More
FillMatrix MSMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix MSMatrix(), 2, Item, "0,-1,0,1,0"
FillMatrix MSMatrix(), 3, Item, "0,-1,1,1,0"
FillMatrix MSMatrix(), 4, Item, "0,-1,0,1,0"
FillMatrix MSMatrix(), 5, Item, "0,0,0,0,0"

Item = 4 ' Emboss
FillMatrix MSMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix MSMatrix(), 2, Item, "0,0,0,-1,0"
FillMatrix MSMatrix(), 3, Item, "0,0,1,0,0"
FillMatrix MSMatrix(), 4, Item, "0,0,0,0,0"
FillMatrix MSMatrix(), 5, Item, "0,0,0,0,0"

Item = 5 ' Emboss More
FillMatrix MSMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix MSMatrix(), 2, Item, "0,1,0,-1,0"
FillMatrix MSMatrix(), 3, Item, "0,1,1,-1,0"
FillMatrix MSMatrix(), 4, Item, "0,1,0,-1,0"
FillMatrix MSMatrix(), 5, Item, "0,0,0,0,0"

Item = 6 ' Edge Enhance
FillMatrix MSMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix MSMatrix(), 2, Item, "0,1,1,1,0"
FillMatrix MSMatrix(), 3, Item, "0,1,1,1,0"
FillMatrix MSMatrix(), 4, Item, "0,1,1,1,0"
FillMatrix MSMatrix(), 5, Item, "0,0,0,0,0"

Item = 7 ' Contour
FillMatrix MSMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix MSMatrix(), 2, Item, "0,1,1,1,0"
FillMatrix MSMatrix(), 3, Item, "0,1,1,1,0"
FillMatrix MSMatrix(), 4, Item, "0,1,1,1,0"
FillMatrix MSMatrix(), 5, Item, "0,0,0,0,0"

Item = 8 ' Relief
FillMatrix MSMatrix(), 1, Item, "0,0,0,0,0"
FillMatrix MSMatrix(), 2, Item, "0,0,-1,-2,0"
FillMatrix MSMatrix(), 3, Item, "0,1,1,-1,0"
FillMatrix MSMatrix(), 4, Item, "0,2,1,0,0"
FillMatrix MSMatrix(), 5, Item, "0,0,0,0,0"

End Sub

Public Sub FillFactors(zArr(), Item, Fac6$)

' zArr(6,Item)
' Fac6$ = "1.1,2,3,4,5,6"

i = 1
p1 = 1
Do
   p2 = InStr(p1, Fac6$, ",")
   If p2 <> 0 Then
      zArr(i, Item) = Val(Mid$(Fac6$, p1, p2 - p1))
      p1 = p2 + 1
   Else
      zArr(i, Item) = Val(Mid$(Fac6$, p1, Len(Fac6$) - p1 + 1))
      Exit Do
   End If
   i = i + 1
Loop

End Sub

Public Sub FillMatrix(Matrix(), iy, Item, Fil$)

' Matrix(x=1-5,iy,item)
' Fil$ ="0,0,0,0,0"  five

i = 1
p1 = 1
Do
   p2 = InStr(p1, Fil$, ",")
   If p2 <> 0 Then
      Matrix(i, iy, Item) = Val(Mid$(Fil$, p1, p2 - p1))
      p1 = p2 + 1
   Else
      Matrix(i, iy, Item) = Val(Mid$(Fil$, p1, Len(Fil$) - p1 + 1))
      Exit Do
   End If
   i = i + 1
Loop

End Sub


Public Sub ShowPalBGR(N)

' Blit PalBGR(N) to PIC

' N= 1,2 or 3

Form1.PIC.Picture = LoadPicture()
Form1.PIC.Visible = True

PalBGRPtr = VarPtr(PalBGR(1, 1, 1, N))

bm.bmiH.biwidth = PICW
bm.bmiH.biheight = PICH

   If StretchDIBits(Form1.PIC.HDC, _
      0, 0, _
      PICW, PICH, _
      0, 0, _
      PICW, PICH, _
      ByVal PalBGRPtr, bm, _
      1, vbSrcCopy) = 0 Then
         
         Erase PalBGR
         MsgBox ("Blit Error")
         End
   
   End If

Form1.PIC.Refresh

End Sub

