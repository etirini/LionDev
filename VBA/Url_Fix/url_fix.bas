Attribute VB_Name = "Module2"

Sub AnchoredURLfix()
'here dims for range finder
Dim rFind, w_range As Range

'here dims of the URL handling code
Dim original_url, part_1, part_2, part_3, new_url, correct_sheet As String
Dim locate_q, locate_p, result, destination_col, count_q As Integer
Dim destination_row, count_p As Long

'Dims for error handling
Dim col_confirm, count_p_warning, count_q_warning, p_q_warning As Boolean
count_p_warning = False
count_q_warning = False


'correct_sheet = ActiveSheet.Name
    If Not ActiveSheet.Name = "Traffic Workbook" Then
        MsgBox ("Select the Traffic Workbook tab before running")
        Exit Sub
    End If
        
    col_confirm = False
    While col_confirm = False
        tagged_url = InputBox("What is the NAME of your Tagged URL Column? [Look for Tagged URL, Clickthrough URL, or any column that gets a url formed by a formula :) ]")
        
        'If user leaves box blank, exit sub
        If tagged_url = "" Then
            MsgBox ("No value was entered. The macro did not run")
            Exit Sub
        Else
        End If
        
        'Define range where to look for the TAGGED URL value
        With Range("A1:BB20")
            Set rFind = .Find(What:=tagged_url, LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)
            If Not rFind Is Nothing Then
                col_confirm = True
                
                'Take the cell value of the tagged url you got earlier and now create a range that can cover all possible cells within the document
                Set w_range = .Range(.Cells(rFind.Row, rFind.column), .Cells(rFind.Row + 25, rFind.column))
                On Error Resume Next
            
                'For each line in the document do this
                For Each c In w_range
                    original_url = c.value
                    count_p = 0
                    count_q = 0
                    p_q_warning = False
                
                    'Strip each url, analize if character is # or ? and count them. If more than one of each issue some warning
                For rep = 1 To Len(original_url)
                    If Mid(original_url, rep, 1) = "#" Then count_p = count_p + 1
                        If count_p >= 2 Then
                            count_p_warning = True
                            p_q_warning = True
                        End If
                    If Mid(original_url, rep, 1) = "?" Then count_q = count_q + 1
                        If count_q >= 2 Then
                            count_q_warning = True
                            p_q_warning = True
                        End If
                Next rep
                
                'Locate the positions of the symbols, checks for Excel error in the formular, reorganizes the url (in case there is only one of each!)
                locate_q = InStr(1, original_url, "?")
                locate_p = InStr(1, original_url, "#")
                locate_e = IsError(c)
                
                If p_q_warning = False And (locate_q > locate_p And locate_p > 0) And locate_e = 0 Then
                    part_1 = Left(original_url, locate_p - 1)
                    part_2 = Mid(original_url, locate_p, (locate_q - locate_p))
                    part_3 = Mid(original_url, locate_q, 1000)
                    new_url = part_1 + part_3 + part_2
                    destination_row = c.Row
                    destination_col = c.column + 1
                    Cells(destination_row, destination_col) = new_url
                End If
                Next
            End If
        End With
        
        'The warnings in caase of duplicate or misplaced values and the information if everything went a-ok
        If count_p_warning Then MsgBox "There are incorrect line items containing Tracking URLs with misplaced anchors (#) in the " & tagged_url & " column." & vbNewLine & vbNewLine & "Please check your Tagged URLs to fix these errors, and then run the macro again.", vbCritical
        If count_q_warning Then MsgBox "There are incorrect line items containing multiple query strings (?) in the " & tagged_url & " column." & vbNewLine & vbNewLine & "Please check your Tagged URLs to fix these errors, and then run the macro again.", vbCritical
        If count_q_warning = False And count_p_warning = False Then MsgBox "The macro ran succesfully", vbInformation
    Wend
End Sub

