VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Alejandra Ramos de Vega
'Inventory of fire bulkheads
'this Excel file works with two different codes:
'Code "FotosL�schenI" deletes all images in row I, to avoid repeated photos
'"AnpassenBilder"  adds all images in row I
'Before selecting the option to run, you should select the desired code from the drop-down list in the upper right corner of this window.

Sub FotosL�schenI()
    

    Const cSpalte1 = 9  ' Column I
    
    'Variables
    Dim sh As Shape
    'Loop
    For Each sh In ActiveSheet.Shapes
        If sh.TopLeftCell.Column = cSpalte1 Then sh.Delete
    Next sh
End Sub
Sub FotosL�schenJ()
    
    Const cSpalte2 = 10  ' Column J
     
    'Variables
    Dim sh As Shape
    'Loop
    For Each sh In ActiveSheet.Shapes
        If sh.TopLeftCell.Column = cSpalte2 Then sh.Delete
    Next sh
End Sub

Sub AnpassenBilder()

'Variables
Dim Adresse As String
Dim BildRange As Range
Dim shp As Object
Dim GanzeAdresse


'Folder address
Adresse = ThisWorkbook.Path

'Column in Excel for the images
ActiveSheet.Range("I7").Select

'Execute until ID column (column C) is empty
Do While ActiveCell.Offset(0, -6).Value <> Empty

    Set BildRange = ActiveCell.Offset(0, -6)
    
 'Save images to Excel file
 
    GanzeAdresse = Adresse & "\Photos\" & BildRange.Value & ".jpg"
    With ActiveSheet.Shapes.AddPicture(Filename:=GanzeAdresse, linktofile:=msoFalse, _
        SaveWithDocument:=msoCTrue, Left:=0, Top:=0, Width:=-1, Height:=-1)
        .LockAspectRatio = 0
        .Top = ActiveCell.Top
        .Left = ActiveCell.Left
        .Width = ActiveCell.Width
        .Height = ActiveCell.Height
    End With
   
    
    ActiveCell.Offset(1, 0).Select
    
Loop

Range("C7").Select
On Error GoTo 0
End Sub


Sub AnpassenBilderJ()

'Variables
Dim AdresseJ As String
Dim BildRangeJ As Range
Dim shp As Object
Dim GanzeAdresseJ

On Error Resume Next

'Folder address
AdresseJ = ThisWorkbook.Path

'Column in Excel for the images
ActiveSheet.Range("J7").Select

'Execute until ID column (column C) is empty
Do While ActiveCell.Offset(0, -7).Value <> Empty

'Save images to Excel file
    Set BildRangeJ = ActiveCell.Offset(0, -7)
    
    '''''''''''''''''''''''''''''''''''''''''''' old Code
    'ActiveSheet.Pictures.Insert(Adresse & "\Bilder\" & BildRange.Value & ".jpg").Select
    
    
    ''''''''''''''''''''''''''''''''''' Speichern Bilder in die Excel
    GanzeAdresseJ = AdresseJ & "\New Photos\" & BildRangeJ.Value & ".jpg"
    With ActiveSheet.Shapes.AddPicture(Filename:=GanzeAdresseJ, linktofile:=msoFalse, _
        SaveWithDocument:=msoCTrue, Left:=0, Top:=0, Width:=-1, Height:=-1)
        .LockAspectRatio = 0
        .Top = ActiveCell.Top
        .Left = ActiveCell.Left
        .Width = ActiveCell.Width
        .Height = ActiveCell.Height
    End With
    
    ActiveCell.Offset(1, 0).Select
    
Loop

Range("C7").Select
On Error GoTo 0
End Sub

