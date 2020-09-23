Attribute VB_Name = "ModuleLS"
' ModuleLS: LoadSave.bas  by Robert Rayment

' NB Refers back to Form1


DefInt A-W
DefSng X-Z

' Input:
Public PathSpec$                    ' Application path
Public AppName$                     ' Name for INI file
Public LoadDir$, SaveDir$           ' Load & Save Folder (from INI file)
Public Title$, Choice$, InitDir$    ' CommonDialog Parameters
Public LoadSave                     ' Boolean: True for Load, False for Save

' Output:
Public LoadFileSpec$, SaveFileSpec$

Public Sub GetInifile()

'Get app path
PathSpec$ = App.Path
If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"

If AppName$ = "" Then AppName$ = "NoName"

IniSpec$ = PathSpec$ & AppName$ & ".ini"

On Error GoTo NoIni
Open IniSpec$ For Input As #1
Line Input #1, LoadDir$
Line Input #1, SaveDir$
Close #1
Exit Sub
'========
NoIni:
LoadDir$ = PathSpec$
SaveDir$ = PathSpec$
Close #1
Open IniSpec$ For Output As #1
Print #1, LoadDir$
Print #1, SaveDir$
Close #1

On Error GoTo 0

Exit Sub
End Sub

Public Sub LSDialog(Title$, Choice$, InitDir$, LoadSave, Optional Extension$)

' LoadSave = True  = Load
' LoadSave = False = Save

Form1.CommonDialog1.DialogTitle = Title$
'&H2 checks if file exists
'&H8 forces save to be same directory as open
Form1.CommonDialog1.Flags = &H2
Form1.CommonDialog1.CancelError = True
On Error GoTo CancelLS
'Leaving the following two lines out allows
'the file to be saved as any extension
Form1.CommonDialog1.Filter = Choice$
Form1.CommonDialog1.InitDir = InitDir$
If LoadSave Then
   Form1.CommonDialog1.FileName = ""
   Form1.CommonDialog1.ShowOpen
   LoadFileSpec$ = Form1.CommonDialog1.FileName
   LoadDir$ = ExtractPath$(LoadFileSpec$)
Else
   Form1.CommonDialog1.FileName = ""
   Form1.CommonDialog1.ShowSave
   SaveFileSpec$ = Form1.CommonDialog1.FileName
   FixFileExtension Extension$           ' SPECIAL
   SaveDir$ = ExtractPath$(SaveFileSpec$)
End If
Exit Sub
'============
CancelLS:
Close

If LoadSave Then
   LoadFileSpec$ = ""
Else
   SaveFileSpec$ = ""
End If
   
On Error GoTo 0

Exit Sub
Resume
End Sub

Public Function ExtractFileName$(FSpec$)
'In:  FSpec$ = Full FileSpec
'Out: Name = ExtractFileName$

ExtractFileName$ = ""

If FSpec$ = "" Then Exit Function

'Find pbs on last backslash \
p = 0: pbs = 0
Do: p = InStr(p + 1, FSpec$, "\")
    If p <> 0 Then pbs = p Else Exit Do
Loop
If pbs > 0 Then
   ExtractFileName$ = UCase$(Mid$(FSpec$, pbs + 1))
End If

End Function

Public Sub FixFileExtension(Ext$)
'Public LoadFileSpec$, SaveFileSpec$
'eg Ext$="pal" etc
E$ = "." + Ext$
pdot = InStr(1, SaveFileSpec$, ".")
If pdot = 0 Then
   SaveFileSpec$ = SaveFileSpec$ + E$
Else
   Ext$ = LCase$(Mid$(SaveFileSpec$, pdot))
   If Ext$ <> E$ Then
      SaveFileSpec$ = Left$(SaveFileSpec$, pdot - 1) + E$
   End If
End If

End Sub

Public Function ExtractPath$(FSpec$)
'In:  FSpec$ = Full FileSpec
'Out: Path = ExtractPath$

ExtractPath$ = ""

If FSpec$ = "" Then Exit Function

'Find pbs on last backslash \
p = 0: pbs = 0
Do
   p = InStr(p + 1, FSpec$, "\")
   If p <> 0 Then pbs = p Else Exit Do
Loop

If pbs = 0 Then
   ExtractPath$ = FSpec$ & "\"   'ie include a last \
Else
   ExtractPath$ = Left$(FSpec$, pbs)
End If

End Function

