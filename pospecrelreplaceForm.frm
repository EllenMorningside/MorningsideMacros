VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pospecrelreplaceForm 
   Caption         =   "pos(spec+rel) Replace Form"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7635
   OleObjectBlob   =   "pospecrelreplaceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "pospecrelreplaceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub cancelButton_Click()

    Unload Me

End Sub


Public Sub okButton_Click()
   

    Application.ScreenUpdating = False

    Dim r As Range
    Set r = ActiveDocument.Content

    ''Replace FieldOne Text
    With ActiveDocument.Content.Find
        
        If fieldOneInput.Value <> "" Then
            .Text = "fieldone"
            .Replacement.ClearFormatting
            .Replacement.Text = fieldOneInput.Value
            .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
        
    End With
    
    
    ''Replace FieldTwo Text
    
    With ActiveDocument.Content.Find

        If fieldTwoInput.Value <> "" Then
        .Text = "FIELDTWO"
        .Replacement.Text = "xxxxxxxxxx"
        .MatchCase = True
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With

    With ActiveDocument.Content.Find
       
        If fieldTwoInput.Value <> "" Then
        .Text = "xxxxxxxxxx"
        .Replacement.ClearFormatting
        .Replacement.Text = fieldTwoInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With

    ''Replace First Name Text
    With ActiveDocument.Content.Find

        If fNameInput.Value <> "" Then
            .Text = "firstname"
            .Replacement.ClearFormatting
            .Replacement.Text = fNameInput.Value
            .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With

    'Replace Last Name Text
    With ActiveDocument.Content.Find

        If lNameInput.Value <> "" Then
            .Text = "lastname"
            .Replacement.ClearFormatting
            .Replacement.Text = lNameInput.Value
            .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    'Replace Position Text
    With ActiveDocument.Content.Find

        If positionInput.Value <> "" Then
            .Text = "positionone"
            .Replacement.ClearFormatting
            .Replacement.Text = positionInput.Value
            .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    'Replace Company Text
    With ActiveDocument.Content.Find

        If companyInput.Value <> "" Then
            .Text = "companyone"
            .Replacement.ClearFormatting
            .Replacement.Text = companyInput.Value
            .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    'Replace DegreeIn Text
    With ActiveDocument.Content.Find

        If degreeInput.Value <> "" Then
            .Text = "degreein"
            .Replacement.ClearFormatting
            .Replacement.Text = degreeInput.Value
            .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
        
    End With

    'Replace Case Number Text
    With ActiveDocument.Content.Find

        If caseNumberInput.Value <> "" Then
            .Text = "casenumber"
            .Replacement.ClearFormatting
            .Replacement.Text = caseNumberInput.Value
            .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With

    Application.ScreenUpdating = True

    MsgBox ("Find and Replace completed successfully!")
    
    Unload Me

End Sub


Private Sub UserForm_Click()

End Sub
