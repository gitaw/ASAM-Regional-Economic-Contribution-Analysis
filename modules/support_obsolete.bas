Attribute VB_Name = "support_obsolete"
'==============================================
'not used anymore but stilll may come in handy
'==============================================

Function z_ObjectExists(strObjectType, strObjectName, myaccess) As Boolean
    Dim db As DAO.Database
    Set db = currentdb
     Dim tbl As TableDef
     Dim qry As QueryDef
     Dim i As Integer
     ObjectExists = False
     If strObjectType = "Table" Then
          For Each tbl In db.TableDefs
               If tbl.Name = strObjectName Then
                    ObjectExists = True
                    Exit Function
               End If
          Next tbl
     ElseIf strObjectType = "Query" Then
          For Each qry In db.QueryDefs
               If qry.Name = strObjectName Then
                    ObjectExists = True
                    Exit Function
               End If
          Next qry
     ElseIf strObjectType = "Form" Or strObjectType = "Report" Or strObjectType = "Module" Then
          For i = 0 To db.Containers(strObjectType & "s").Documents.Count - 1
               If db.Containers(strObjectType & "s").Documents(i).Name = strObjectName Then
                    ObjectExists = True
                    Exit Function
               End If
          Next i
     ElseIf strObjectType = "Macro" Then
          For i = 0 To db.Containers("Scripts").Documents.Count - 1
               If db.Containers("Scripts").Documents(i).Name = strObjectName Then
                    ObjectExists = True
                    Exit Function
               End If
          Next i
     Else
          MsgBox "Invalid Object Type passed, must be Table, Query, Form, Report, Macro, or Module"
     End If
End Function


