Attribute VB_Name = "Module1"
Option Explicit
Global r, g, b As Byte

Dim blank_fnd_build_chkbx, paint_cells_chkbx, fnd_build_chkbx, status_chkbx, resolution_chkbx, module_chkbx, terminal_chkbx, issue_type_chkbx As OLEObject
Dim change_color As Boolean
Dim triage_range As Long

Sub RDD1()
Dim x As Long
Dim y As Long

triage_range = Range("C2").End(xlDown).Row

Set_checkboxes
rst
Next_color

'---REDO Defects V0.4.3-----

    For x = 2 To triage_range
        change_color = False
        
        For y = 2 To triage_range
        
            If ReDo_conditions(x, y) = True Then
            Cells(y, 6) = "Redo"
            Cells(x, 6) = "Parent"
            Cells(x, 4) = Cells(x, 3)
            End If
            
        Next y
        
        If change_color = True Then Next_color
        
    Next x
    
    Fix_mtx
    
    Show_RedoID

End Sub

Function ReDo_conditions(i As Long, j As Long) As Boolean
Dim linked_triage, triage, aux As String
Dim linked_triage_mtx() As String
Dim link_triage As Variant
Dim redo, exist_redo As Boolean
Dim c As Byte

'first cell of mtx
c = 18

exist_redo = False
aux = ""
triage = Cells(i, 3)
linked_triage_mtx = Split(Cells(j, 5), ",")

'Remove merge triages

If Cells(j, 1) <> Cells(i, 1) Then

    For Each link_triage In linked_triage_mtx
        
        redo = False
    
        If link_triage = triage Then
            
        'Resolution date
        triage = Cells(i, 17)
        'Creation date
        linked_triage = Cells(j, 16)
        
            'check if triage was resolved before the linked triage was created
            If CDate(triage) <= CDate(linked_triage) Then
            
                redo = True
                
                '--------------Redo conditions--------------
                
                'Check that the version number is not the same
                If fnd_build_chkbx.Object.Value = True Then
                    
                    triage = Cells(i, 7)
                    linked_triage = Cells(j, 7)
                    
                    If linked_triage = triage Then redo = False
                    
                    'Consider defects with blank version number as posible redos
                    If blank_fnd_build_chkbx.Object.Value = True Then
                        If linked_triage = "" And triage = "" Then redo = True
                    End If
                End If
                
                'Check resolution = Done
                If resolution_chkbx.Object.Value = True Then
                    triage = Cells(i, 10)
                    linked_triage = Cells(j, 10)
                    If triage <> "Done" Or triage <> "Unresolved" Then redo = False
                End If
                'Check for same terminal
                If terminal_chkbx.Object.Value = True Then
                triage = Cells(i, 13)
                linked_triage = Cells(j, 13)
                If linked_triage <> triage Then redo = False
                End If
                 
                'New rules can be entered here:
                
                '--------------Redo conditions--------------
                
            End If
            
        End If
        
        If redo = True Then
            change_color = 1
            'Paint Cells
            If paint_cells_chkbx.Object.Value = True Then
                Cells(i, 3).Interior.color = RGB(r, g, b)
                Cells(j, 4).Interior.color = RGB(r, g, b)
                'Cells(j, 5).Interior.color = RGB(r, g, b)
            End If
            
            'Enter the REDo ID
            Cells(j, c) = Cells(i, 3)
            'Enter the TRIAGE ID
            Cells(j, 2) = Cells(i, 1)
            
            exist_redo = True
            
        End If
            
    c = c + 1
    
    Next link_triage
    
 End If
 
   ReDo_conditions = exist_redo
    
End Function
Sub Fix_mtx()
    Range(Cells(2, 18), Cells(triage_range, 28)).SpecialCells(xlCellTypeBlanks).Select
    Selection.Delete Shift:=xlToLeft
End Sub
Sub Show_RedoID()
Dim d, e As Byte
Dim triage_str As String

    For d = 2 To triage_range
        
        If Cells(d, 18).Value <> "" Then
            
            For e = 18 To 28
                If (triage_str = "" And Cells(d, e).Value <> "") Then
                triage_str = Cells(d, e)
                ElseIf Cells(d, e).Value <> "" Then triage_str = triage_str & "," & Cells(d, e)
                End If
                'erase mtx
                Cells(d, e).Value = ""
            Next
            
            Cells(d, 4) = triage_str
            triage_str = ""
            
        End If
    Next
    
    Range(Cells(d, 18), Cells(d, 28)).ClearContents
    
End Sub

Sub rst()
    Dim rng1, rng2 As Long
    
    rng2 = Range("A2").End(xlDown).Row
    
    For rng1 = 2 To rng2
    
        Cells(rng1, 3).Interior.color = xlNone
        Cells(rng1, 4).Interior.color = xlNone
        
        Cells(rng1, 2).ClearContents
        Cells(rng1, 4).ClearContents
        Cells(rng1, 6).ClearContents
        Range(Cells(rng1, 18), Cells(rng1, 28)).ClearContents
        
    Next rng1
    
End Sub

Sub Next_color()
r = Application.WorksheetFunction.RandBetween(100, 240)
g = Application.WorksheetFunction.RandBetween(100, 240)
b = Application.WorksheetFunction.RandBetween(100, 240)
End Sub

Sub Set_checkboxes()
Set blank_fnd_build_chkbx = ActiveSheet.OLEObjects("CheckBox1")
Set fnd_build_chkbx = ActiveSheet.OLEObjects("CheckBox2")
Set status_chkbx = ActiveSheet.OLEObjects("CheckBox3")
Set resolution_chkbx = ActiveSheet.OLEObjects("CheckBox4")
Set module_chkbx = ActiveSheet.OLEObjects("CheckBox5")
Set terminal_chkbx = ActiveSheet.OLEObjects("CheckBox6")
Set issue_type_chkbx = ActiveSheet.OLEObjects("CheckBox7")
Set paint_cells_chkbx = ActiveSheet.OLEObjects("CheckBox_PntC")
End Sub

