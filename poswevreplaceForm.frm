VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} poswevreplaceForm 
   Caption         =   "pos(wev) Replace Form"
   ClientHeight    =   13035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14145
   OleObjectBlob   =   "poswevreplaceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "poswevreplaceForm"
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
    
    ''Replace degreereceivedone Text
    With ActiveDocument.Content.Find

        If degreeReceivedOneInput.Value <> "" Then
        .Text = "degreereceivedone"
        .Replacement.ClearFormatting
        .Replacement.Text = degreeReceivedOneInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        
         .Text = "degreerecievedone"
        .Replacement.ClearFormatting
        .Replacement.Text = degreeReceivedOneInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace in foreign degree Text
    With ActiveDocument.Content.Find
    
        If inForeignDegreeInput.Value <> "" Then
        .Text = "inforeigndegree"
        .Replacement.ClearFormatting
        .Replacement.Text = inForeignDegreeInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
      End If
    
    End With
    
    ''Replace title of foreign degree Text
    With ActiveDocument.Content.Find
    
        If titleForeignDegreeInput.Value <> "" Then
        .Text = "titleforeigndegree"
        .Replacement.ClearFormatting
        .Replacement.Text = titleForeignDegreeInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace major three Text
    With ActiveDocument.Content.Find
    
        If majorThreeInput.Value <> "" Then
        .Text = "majorthree"
        .Replacement.ClearFormatting
        .Replacement.Text = majorThreeInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace major two Text
    With ActiveDocument.Content.Find

        If majorTwoInput.Value <> "" Then
        .Text = "majortwo"
        .Replacement.ClearFormatting
        .Replacement.Text = majorTwoInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace year three Text
    With ActiveDocument.Content.Find

        If yearThreeInput.Value <> "" Then
        .Text = "yearthree"
        .Replacement.ClearFormatting
        .Replacement.Text = yearThreeInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With

    ''Replace year two three Text
    With ActiveDocument.Content.Find

        If yearTwoInput.Value <> "" Then
        .Text = "yeartwo"
        .Replacement.ClearFormatting
        .Replacement.Text = yearTwoInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace degree received three Text
    With ActiveDocument.Content.Find

        If degreeReceivedThreeInput.Value <> "" Then
        .Text = "degreereceivedthree"
        .Replacement.ClearFormatting
        .Replacement.Text = degreeReceivedThreeInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace degree received two Text
    With ActiveDocument.Content.Find

        If degreeReceivedTwoInput.Value <> "" Then
        .Text = "degreereceivedtwo"
        .Replacement.ClearFormatting
        .Replacement.Text = degreeReceivedTwoInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace the country three Text
    With ActiveDocument.Content.Find

        If theCountryThreeInput.Value <> "" Then
        .Text = "thecountrythree"
        .Replacement.ClearFormatting
        .Replacement.Text = theCountryThreeInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace the country two Text
    With ActiveDocument.Content.Find

        If theCountryTwoInput.Value <> "" Then
        .Text = "thecountrytwo"
        .Replacement.ClearFormatting
        .Replacement.Text = theCountryTwoInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace school name three Text
    If prefixTheCheckBox3 = True Then
    
        With ActiveDocument.Content.Find
        .Text = "SchoolNameThree"
        .Replacement.ClearFormatting
        .Replacement.Text = "the SchoolNameThree"
        .MatchCase = True
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End With
        
    End If
    
    With ActiveDocument.Content.Find

        If schoolNameThreeInput.Value <> "" Then
        .Text = "schoolnamethree"
        .Replacement.ClearFormatting
        .Replacement.Text = schoolNameThreeInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace school name two Text
    If prefixTheCheckBox2 = True Then
    
        With ActiveDocument.Content.Find
        .Text = "SchoolNameTwo"
        .Replacement.ClearFormatting
        .Replacement.Text = "the SchoolNameTwo"
        .MatchCase = True
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End With
        
    End If
    
    With ActiveDocument.Content.Find

        If schoolNameTwoInput.Value <> "" Then
        .Text = "schoolnametwo"
        .Replacement.ClearFormatting
        .Replacement.Text = schoolNameTwoInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace years of coursework three Text
    With ActiveDocument.Content.Find

        If yearsOfCourseworkThreeInput.Value <> "" Then
        .Text = "yearsofcourseworkthree"
        .Replacement.ClearFormatting
        .Replacement.Text = yearsOfCourseworkThreeInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace years of coursework two Text
    With ActiveDocument.Content.Find

        If yearsOfCourseworkTwoInput.Value <> "" Then
        .Text = "yearsofcourseworktwo"
        .Replacement.ClearFormatting
        .Replacement.Text = yearsOfCourseworkTwoInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace years of coursework one Text
    With ActiveDocument.Content.Find
    
        If yearsOfCourseworkOneInput.Value <> "" Then
        .Text = "yearsofcourseworkone"
        .Replacement.ClearFormatting
        .Replacement.Text = yearsOfCourseworkOneInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    

    ''Replace major one Text
    With ActiveDocument.Content.Find

        If majorOneInput.Value <> "" Then
        .Text = "majorone"
        .Replacement.ClearFormatting
        .Replacement.Text = majorOneInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace year one Text
    With ActiveDocument.Content.Find

        If yearOneInput.Value <> "" Then
        .Text = "yearone"
        .Replacement.ClearFormatting
        .Replacement.Text = yearOneInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace the country one Text
    With ActiveDocument.Content.Find

        If theCountryOneInput.Value <> "" Then
        .Text = "thecountryone"
        .Replacement.ClearFormatting
        .Replacement.Text = theCountryOneInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace school name one Text
    If prefixTheCheckBox1 = True Then
    
        With ActiveDocument.Content.Find
        .Text = "SchoolNameOne"
        .Replacement.ClearFormatting
        .Replacement.Text = "the SchoolNameOne"
        .MatchCase = True
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End With
        
    End If
    
    With ActiveDocument.Content.Find

        If schoolNameOneInput.Value <> "" Then
        .Text = "schoolnameone"
        .Replacement.ClearFormatting
        .Replacement.Text = schoolNameOneInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With
    
    ''Replace number of years Text
    
    With ActiveDocument.Content.Find
        If numberOfYearsInput.Value <> "" Then
        .Text = "numberofyears"
        .Replacement.ClearFormatting
        .Replacement.Text = numberOfYearsInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
    
    End With

    ''Replace FieldOne Text
    With ActiveDocument.Content.Find
    
        If fieldOneInput.Value <> "" Then
        .Text = "fieldone"
        .Replacement.ClearFormatting
        .Replacement.Text = fieldOneInput.Value
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

    'Replace Degreereceivedone Text
    With ActiveDocument.Content.Find

        If degreeReceivedOneInput.Value <> "" Then
        .Text = "degreereceivedone"
        .Replacement.ClearFormatting
        .Replacement.Text = degreeReceivedOneInput.Value
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

    'Replace degreetitle Text
    With ActiveDocument.Content.Find

        If degreeTitleInput.Value <> "" Then
        .Text = "degreetitle"
        .Replacement.ClearFormatting
        .Replacement.Text = degreeTitleInput.Value
        .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindContinue
        End If
            
    End With

    Application.ScreenUpdating = True

    MsgBox ("Find and Replace completed successfully!")
    
    Unload Me

End Sub

