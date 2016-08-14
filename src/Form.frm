VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Enter Match Information"
   ClientHeight    =   9405.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9750.001
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const allowedChars As String = "1234567890"
Const allowedCharsNeg As String = "-1234567890"
Const allowedCharsK As String = ".kK1234567890"

Dim Roles(1 To 5) As String
Dim Champions(1 To 132) As String
Dim Ranks(1 To 27) As String

Public Sub chk_Date_Click()
    If chk_Date.Value = True Then
        txt_Date.Enabled = True
        txt_Date.BackColor = txt_Length_M.BackColor
        If chk_Tab_Date.Value = False Then
            txt_Date.TabStop = True
        End If
    Else
        txt_Date.Enabled = False
        txt_Date.BackColor = Me.BackColor
        txt_Date.TabStop = False
    End If
End Sub

Public Sub chk_Dodge_Click()
    If chk_Dodge.Value = True Then
        txt_LP_Base.Enabled = True
        txt_LP_Base.BackColor = txt_Length_M.BackColor
        If chk_Tab_Dodge.Value = False Then
            txt_LP_Base.TabStop = True
        End If
    Else
        txt_LP_Base.Enabled = False
        txt_LP_Base.BackColor = Me.BackColor
        txt_LP_Base.TabStop = False
    End If
End Sub

Public Sub chk_Tab_Cancel_Click()
    If chk_Tab_Clear.Value = True Then
            btn_Cancel.TabStop = False
    Else
            btn_Cancel.TabStop = True
    End If
End Sub

Public Sub chk_Tab_Clear_Click()
    If chk_Tab_Clear.Value = True Then
            btn_Clear.TabStop = False
            btn_Cancel.TabStop = False
    Else
            btn_Clear.TabStop = True
    End If
End Sub

Public Sub chk_Tab_Date_Click()
    If chk_Tab_Date.Value = True Then
            chk_Date.TabStop = False
            txt_Date.TabStop = False
    Else
            chk_Date.TabStop = True
            txt_Date.TabStop = True
    End If
End Sub

Public Sub chk_Tab_Dodge_Click()
    If chk_Tab_Dodge.Value = True Then
            chk_Dodge.TabStop = False
    Else
            chk_Dodge.TabStop = True
    End If
End Sub

Public Sub chk_Tab_Screenshot_Click()
    If chk_Tab_Screenshot.Value = True Then
            chk_Screenshot.TabStop = False
    Else
            chk_Screenshot.TabStop = True
    End If
End Sub

Public Sub Chk_Tab_Settings_Click()
    If Chk_Tab_Settings.Value = True Then
            chk_Clear_Rank.TabStop = False
            chk_Submit_Clear.TabStop = False
            chk_Submit_Close.TabStop = False
            chk_Save.TabStop = False
    Else
            chk_Clear_Rank.TabStop = True
            chk_Submit_Clear.TabStop = True
            chk_Submit_Close.TabStop = True
            chk_Save.TabStop = True
    End If
End Sub

Public Sub chk_Tab_Submit_Click()
    If chk_Tab_Submit.Value = True Then
            btn_Submit.TabStop = False
    Else
            btn_Submit.TabStop = True
    End If
End Sub

Public Sub txt_Length_M_Change()
' Only allows numerical values

        Dim inp As String
        Dim subStr As String
        Dim index As Integer
        inp = txt_Length_M.Text
        index = txt_Length_M.SelStart

        If index <> 0 Then
            subStr = Mid(txt_Length_M.Text, index, 1)
            If InStr(1, allowedChars, subStr, vbTextCompare) = 0 Then
                inp = Replace(inp, subStr, "")
                txt_Length_M.Text = inp
                If index <> 0 Then
                    txt_Length_M.SelStart = (index - 1)
                End If
            End If
        End If
End Sub

Public Sub txt_Length_S_Change()
' Only allows numerical values

        Dim inp As String
        Dim subStr As String
        Dim index As Integer
        inp = txt_Length_S.Text
        index = txt_Length_S.SelStart

        If index <> 0 Then
            subStr = Mid(txt_Length_S.Text, index, 1)
            If InStr(1, allowedChars, subStr, vbTextCompare) = 0 Then
                inp = Replace(inp, subStr, "")
                txt_Length_S.Text = inp
                If index <> 0 Then
                    txt_Length_S.SelStart = (index - 1)
                End If
            End If
        End If
End Sub

Public Sub LoadSettings()
    Dim path As String
    Dim DataLine As String
    
    path = ThisWorkbook.path & "\settings.ini"
    Open path For Input As #2
     
    Dim LineNum As Long
    Dim VBComp As Object
    Dim VBCodeMod As Object
     
    'Set VBComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
    'VBComp.Name = "NewModule"
    Set VBCodeMod = ThisWorkbook.VBProject.VBComponents("SettingsModule").CodeModule
    
    With VBCodeMod
        LineNum = .CountOfLines + 1
        .InsertLines LineNum, "Public Sub LoadSettings()"
    End With
    Do Until EOF(2)
        Line Input #2, DataLine
        With VBCodeMod
            LineNum = .CountOfLines + 1
            .InsertLines LineNum, DataLine
        End With
    Loop
    Close #2
    With VBCodeMod
        LineNum = .CountOfLines + 1
        .InsertLines LineNum, "End Sub"
    End With
     
     'run the new module
    Application.Run "SettingsModule"
     'remove the new module
    With VBCodeMod
        .DeleteLines 1, 24
    End With
     'ThisWorkbook.VBProject.VBComponents.Remove VBComp
End Sub

Private Sub UserForm_Activate()
    'Add champions to array
    'How to start array with these?
    Champions(1) = "Aatrox"
    Champions(2) = "Ahri"
    Champions(3) = "Akali"
    Champions(4) = "Alistar"
    Champions(5) = "Amumu"
    Champions(6) = "Anivia"
    Champions(7) = "Annie"
    Champions(8) = "Ashe"
    Champions(9) = "Aurelion Sol"
    Champions(10) = "Azir"
    Champions(11) = "Bard"
    Champions(12) = "Blitzcrank"
    Champions(13) = "Brand"
    Champions(14) = "Braum"
    Champions(15) = "Caitlyn"
    Champions(16) = "Cassiopeia"
    Champions(17) = "Cho'Gath"
    Champions(18) = "Corki"
    Champions(19) = "Darius"
    Champions(20) = "Diana"
    Champions(21) = "Dr. Mundo"
    Champions(22) = "Draven"
    Champions(23) = "Ekko"
    Champions(24) = "Elise"
    Champions(25) = "Evelynn"
    Champions(26) = "Ezreal"
    Champions(27) = "Fiddlesticks"
    Champions(28) = "Fiora"
    Champions(29) = "Fizz"
    Champions(30) = "Galio"
    Champions(31) = "Gankplank"
    Champions(32) = "Garen"
    Champions(33) = "Gnar"
    Champions(34) = "Gragas"
    Champions(35) = "Graves"
    Champions(36) = "Hecarim"
    Champions(37) = "Heimerdinger"
    Champions(38) = "Illaoi"
    Champions(39) = "Irelia"
    Champions(40) = "Janna"
    Champions(41) = "Jarvan IV"
    Champions(42) = "Jax"
    Champions(43) = "Jayce"
    Champions(44) = "Jhin"
    Champions(45) = "Jinx"
    Champions(46) = "Kalista"
    Champions(47) = "Karma"
    Champions(48) = "Karthus"
    Champions(49) = "Kassadin"
    Champions(50) = "Katarina"
    Champions(51) = "Kayle"
    Champions(52) = "Kennen"
    Champions(53) = "Kha'Zix"
    Champions(54) = "Kindred"
    Champions(55) = "Kled"
    Champions(56) = "Kog'Maw"
    Champions(57) = "LeBlanc"
    Champions(58) = "Lee Sin"
    Champions(59) = "Leona"
    Champions(60) = "Lissandra"
    Champions(61) = "Lucian"
    Champions(62) = "Lulu"
    Champions(63) = "Lux"
    Champions(64) = "Malphite"
    Champions(65) = "Malzahar"
    Champions(66) = "Maokai"
    Champions(67) = "Master Yi"
    Champions(68) = "Miss Fortune"
    Champions(69) = "Mordekaiser"
    Champions(70) = "Morgana"
    Champions(71) = "Nami"
    Champions(72) = "Nasus"
    Champions(73) = "Nautilus"
    Champions(74) = "Nidalee"
    Champions(75) = "Nocturne"
    Champions(76) = "Nunu"
    Champions(77) = "Olaf"
    Champions(78) = "Orianna"
    Champions(79) = "Pantheon"
    Champions(80) = "Poppy"
    Champions(81) = "Quinn"
    Champions(82) = "Rammus"
    Champions(83) = "Rek'Sai"
    Champions(84) = "Renekton"
    Champions(85) = "Rengar"
    Champions(86) = "Riven"
    Champions(87) = "Rumble"
    Champions(88) = "Ryze"
    Champions(89) = "Sejuani"
    Champions(90) = "Shaco"
    Champions(91) = "Shen"
    Champions(92) = "Shyvana"
    Champions(93) = "Singed"
    Champions(94) = "Sion"
    Champions(95) = "Sivir"
    Champions(96) = "Skarner"
    Champions(97) = "Sona"
    Champions(98) = "Soraka"
    Champions(99) = "Swain"
    Champions(100) = "Syndra"
    Champions(101) = "Tahm Kench"
    Champions(102) = "Taliyah"
    Champions(103) = "Talon"
    Champions(104) = "Taric"
    Champions(105) = "Teemo"
    Champions(106) = "Thresh"
    Champions(107) = "Tristana"
    Champions(108) = "Trundle"
    Champions(109) = "Tryndamere"
    Champions(110) = "Twisted Fate"
    Champions(111) = "TWitch"
    Champions(112) = "Udyr"
    Champions(113) = "Urgot"
    Champions(114) = "Varus"
    Champions(115) = "vayne"
    Champions(116) = "Veigar"
    Champions(117) = "Vel'Koz"
    Champions(118) = "Vi"
    Champions(119) = "Viktor"
    Champions(120) = "Vladimir"
    Champions(121) = "Volibear"
    Champions(122) = "Warwick"
    Champions(123) = "Wukong"
    Champions(124) = "Xerath"
    Champions(125) = "Xin Zhao"
    Champions(126) = "Yasuo"
    Champions(127) = "Yorick"
    Champions(128) = "Zac"
    Champions(129) = "Zed"
    Champions(130) = "Ziggs"
    Champions(131) = "Zilean"
    Champions(132) = "Zyra"
    
    'Add to roles array
    Roles(1) = "Top"
    Roles(2) = "Jungle"
    Roles(3) = "Middle"
    Roles(4) = "ADC"
    Roles(5) = "Support"
    
    'Add to ranks array
    Ranks(1) = "Challenger"
    Ranks(2) = "Master"
    Ranks(3) = "Diamond I"
    Ranks(4) = "Diamond II"
    Ranks(5) = "Diamond III"
    Ranks(6) = "Diamond IV"
    Ranks(7) = "Diamond V"
    Ranks(8) = "Platinum I"
    Ranks(9) = "Platinum II"
    Ranks(10) = "Platinum III"
    Ranks(11) = "Platinum IV"
    Ranks(12) = "Platinum V"
    Ranks(13) = "Gold I"
    Ranks(14) = "Gold II"
    Ranks(15) = "Gold III"
    Ranks(16) = "Gold IV"
    Ranks(17) = "Gold V"
    Ranks(18) = "Silver I"
    Ranks(19) = "Silver II"
    Ranks(20) = "Silver III"
    Ranks(21) = "Silver IV"
    Ranks(22) = "Silver V"
    Ranks(23) = "Bronze I"
    Ranks(24) = "Bronze II"
    Ranks(25) = "Bronze III"
    Ranks(26) = "Bronze IV"
    Ranks(27) = "Bronze V"

    'Tab Indexing
    'General
    chk_Screenshot.TabIndex = 0
    txt_Screenshot.TabIndex = 1
    txt_Length_M.TabIndex = 2
    txt_Length_S.TabIndex = 3
    
    'Statistics
    txt_Kills.TabIndex = 4
    txt_Deaths.TabIndex = 5
    txt_Assists.TabIndex = 6
    txt_CS.TabIndex = 7
    txt_Gold.TabIndex = 8
    
    'Rank
    cmb_Rank.TabIndex = 9
    txt_LP.TabIndex = 10
    chk_Dodge.TabIndex = 11
    txt_LP_Base.TabIndex = 12
    txt_Grade.TabIndex = 13
    
    'Champion & Lane
    cmb_Role.TabIndex = 14
    cmb_Champ.TabIndex = 15
    cmb_Opp.TabIndex = 16
    
    MultiPage1.TabIndex = 17
    
    'Settings
    chk_Date.TabIndex = 18
    txt_Date.TabIndex = 19
    chk_Clear_Rank.TabIndex = 20
    chk_Submit_Clear.TabIndex = 21
    chk_Submit_Close.TabIndex = 22
    chk_Save.TabIndex = 23
    
    'Buttons
    btn_Submit.TabIndex = 24
    btn_Clear.TabIndex = 25
    btn_Cancel.TabIndex = 26
    
    'Groups
    frm_Screenshot.TabIndex = 0
    frm_Stats.TabIndex = 1
    frm_Rank.TabIndex = 2
    frm_Champ.TabIndex = 3
    MultiPage1.TabIndex = 4
    frm_Other.TabIndex = 5
    
    'Populate combo box lists
    For i = 1 To 27
        cmb_Rank.AddItem Ranks(i)
    Next
    For i = 1 To 5
        cmb_Role.AddItem Roles(i)
    Next
    For i = 1 To 132
        cmb_Champ.AddItem Champions(i)
        cmb_Opp.AddItem Champions(i)
    Next
    
    'Set Defaults
    txt_Date.BackColor = Me.BackColor
    txt_LP_Base.BackColor = Me.BackColor
    chk_Date.TabStop = False
    txt_Date.TabStop = False
    chk_Clear_Rank.TabStop = False
    chk_Submit_Clear.TabStop = False
    chk_Submit_Close.TabStop = False
    chk_Save.TabStop = False
    btn_Clear.TabStop = False
    btn_Cancel.TabStop = False
    
    'LoadSettings
End Sub

Public Sub saveSettings()
    Dim path As String
    Dim strWr As String
    
    path = ThisWorkbook.path & "\settings.ini"
    Open path For Output As #1
    
    For Each cntrl In Me.Controls
        If TypeName(cntrl) = "CheckBox" Then
            strWr = cntrl.Name & ".Value = " & cntrl.Value
            Print #1, strWr
        End If
    Next
    If chk_Clear_Rank.Value = False Then
        strWr = cmb_Rank.Name & ".Value = " & Chr(34) & cmb_Rank.Value & Chr(34)
        Print #1, strWr
    End If
    Close #1
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Public Sub resetColours()
    'Resets colour
    For Each cntrl In Me.Controls
        If TypeName(cntrl) = "ComboBox" Then
            cntrl.BackColor = txt_Prob.BackColor
            cntrl.BorderColor = txt_Prob.BorderColor
            cntrl.ForeColor = txt_Prob.ForeColor
        End If
        If TypeName(cntrl) = "TextBox" Then
            Select Case cntrl.Name
                'Prevents screenshot field from changing if disabled
                Case txt_Screenshot.Name
                    If txt_Screenshot.Enabled = False Then
                        GoTo NextIterationColour
                    Else
                        GoTo SetColor
                    End If
                Case txt_Date.Name
                    If txt_Date.Enabled = False Then
                        GoTo NextIterationColour
                    Else
                        GoTo SetColor
                    End If
                Case txt_LP_Base.Name
                    If txt_LP_Base.Enabled = False Then
                        GoTo NextIterationColour
                    Else
                        GoTo SetColor
                    End If
                Case Else
SetColor:
                    cntrl.BackColor = txt_Prob.BackColor
                    cntrl.BorderColor = txt_Prob.BorderColor
                    cntrl.ForeColor = txt_Prob.ForeColor
            End Select
        End If
NextIterationColour:
    Next
End Sub

Public Sub btn_Cancel_Click()
    Unload Me
End Sub

Public Sub btn_Clear_Click()
    'General
    If chk_Clear_Screenshot.Value = True Then
        chk_Screenshot.Value = False
    End If
    txt_Screenshot.Text = ""
    txt_Length_M.Text = ""
    txt_Length_S.Text = ""
    
    'Statistics
    txt_Kills.Text = ""
    txt_Deaths.Text = ""
    txt_Assists.Text = ""
    txt_CS.Text = ""
    txt_Grade.Text = ""
    
    'Rank
    cmb_Rank.Text = ""
    txt_LP.Text = ""
    If chk_Clear_Dodge.Value = True Then
        chk_Dodge.Value = False
    End If
    txt_LP_Base.Text = ""
    
    'Champion & Lane
    cmb_Role.Text = ""
    cmb_Champ.Text = ""
    cmb_Opp.Text = ""
    
    'Date Settings
    If chk_Clear_Date.Value = True Then
        txt_Date.Text = ""
    End If
    If chk_Clear_Date_Chk.Value = True Then
        chk_Date.Value = False
    End If
    'Rank Settings
    If chk_Clear_Settings.Value = True Then
        chk_Clear_Rank.Value = False
    End If
        
    resetColours
End Sub

'Public Sub cFormat(ByVal r, ByVal f, ByVal colour)
'    With Worksheets(1).Range(r).FormatCondition(1)
'        .Modify
'    End With
'End Sub

Public Sub btn_Submit_Click()
    
    'cFormat "$H$2:$H$1048576", 1, "blue"
    
    Dim iRow As Long
    Dim ws As Worksheet
    Set ws = Worksheets("Sheet1")
    Dim cntrl As Control
    Dim bErr As Boolean
    bErr = False
    
    'Saves Settings
    'If chk_Save.Value = True Then
    '    saveSettings
    'End If
    
    'Finds the first empty row
    iRow = ws.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1
    
    resetColours
    
    'Checks for empty or invalid fields
    If chk_Error_All.Value = False Then
        'Gold K abbreviation
        Dim strDecK As String
        Dim strGold As String
        strGold = txt_Gold.Text
        If InStr(Len(strGold), strGold, "k", vbTextCompare) = 0 Then
            If InStr(1, strGold, ".", vbTextCompare) <> 0 Then
                txt_Gold.BackColor = RGB(255, 199, 206)
                txt_Gold.BorderColor = RGB(156, 0, 6)
                txt_Gold.ForeColor = RGB(156, 0, 6)
                bErr = True
            End If
        Else
            strGold = Mid(strGold, 1, Len(strGold) - 1)
            txt_Gold.Value = CDbl(strGold) * 1000
            If txt_Gold.Value < 1 Then
                bErr = True
            End If
        End If
        
        For Each cntrl In Me.Controls
            If TypeName(cntrl) = "ComboBox" Then
                Select Case cntrl.Name
                    Case cmb_Rank.Name
                        If (IsInArray(cntrl.Text, Ranks) = False) Or (cntrl.Text = "") Then
                            cntrl.BackColor = RGB(255, 199, 206)
                            cntrl.BorderColor = RGB(156, 0, 6)
                            cntrl.ForeColor = RGB(156, 0, 6)
                            bErr = True
                        End If
                    Case cmb_Role.Name
                        If (IsInArray(cntrl.Text, Roles) = False) Or (cntrl.Text = "") Then
                            cntrl.BackColor = RGB(255, 199, 206)
                            cntrl.BorderColor = RGB(156, 0, 6)
                            cntrl.ForeColor = RGB(156, 0, 6)
                            bErr = True
                        End If
                    Case Else
                        If chk_Error_Champ.Value = False Then
                            If (IsInArray(cntrl.Text, Champions) = False) Or (cntrl.Text = "") Then
                                cntrl.BackColor = RGB(255, 199, 206)
                                cntrl.BorderColor = RGB(156, 0, 6)
                                cntrl.ForeColor = RGB(156, 0, 6)
                                bErr = True
                            End If
                        End If
                End Select
            End If
            If TypeName(cntrl) = "TextBox" Then
                Select Case cntrl.Name
                    'Excludes optional comment fields
                    Case txt_Prob.Name
                        GoTo NextIteration
                    Case txt_Other.Name
                        GoTo NextIteration
                    Case txt_Screenshot.Name
                        If txt_Screenshot.Enabled = False Then
                            GoTo NextIteration
                        Else
                            GoTo CheckEmpty
                        End If
                    Case txt_Date.Name
                        If txt_Date.Enabled = False Then
                            GoTo NextIteration
                        Else
                            GoTo CheckEmpty
                        End If
                    Case txt_LP_Base.Name
                        If txt_LP_Base.Enabled = False Then
                            GoTo NextIteration
                        Else
                            GoTo CheckEmpty
                        End If
                    Case Else
CheckEmpty:
                        If cntrl.Text = "" Then
                            cntrl.BackColor = RGB(255, 199, 206)
                            cntrl.BorderColor = RGB(156, 0, 6)
                            cntrl.ForeColor = RGB(156, 0, 6)
                            bErr = True
                        End If
                End Select
            End If
NextIteration:
        Next
    End If
    
    If bErr = True Then
        'MsgBox "The were errors in your inputs. The erreneous fields have been marked in red.", vbOKOnly, "Incomplete Form"
        Exit Sub
    End If
    
    Dim currentNum As Integer
    currentNum = iRow - 1
    
    'Copies the data to the spreadsheet
    'Column A - Number
    ws.Cells(iRow, 1).Value = currentNum

    'Column B - Date
    Dim currentDate As String
    If chk_Date.Value = True Then
        currentDate = txt_Date.Value
    Else
        currentDate = Date
    End If
      
    ws.Cells(iRow, 2).Value = currentDate

    'Column C - Length
    ws.Cells(iRow, 3).Value = txt_Length_M.Value & ":" & txt_Length_S.Value

    'Column D - Tier
    Dim RanksArray() As String
    RanksArray() = Split(cmb_Rank.Value, , , vbTextCompare)
    ws.Cells(iRow, 4).Value = RanksArray(0)

    'Column E - Divison
    Dim divison As String
    
    If UBound(RanksArray, 1) < 1 Then
        ws.Cells(iRow, 5).Value = ""
    Else
        Select Case RanksArray(1)
            Case "I"
                division = "1"
            Case "II"
                division = "2"
            Case "III"
                division = "3"
            Case "IV"
                division = "4"
            Case "V"
                division = "5"
        End Select
        ws.Cells(iRow, 5).Value = division
    End If

    'Column F - LP Change
    ws.Cells(iRow, 6).Value = txt_LP.Value

    'Column G - Final LP
    Dim baseLP As Variant
    Dim finalLP As Integer
    
    If currentNum = 1 Then
        If chk_Dodge.Value = False Then
            baseLP = Application.InputBox(prompt:="Since this is the first match being entered, enter the amount of LP you currently have after this match.", Title:="Error Calculating LP", Type:=1)
            If baseLP = False Then
                Exit Sub
            Else
                finalLP = baseLP
            End If
        Else
            baseLP = CInt(txt_LP_Base.Text)
            finalLP = baseLP
        End If
    Else
        If chk_Dodge.Value = True Then
            baseLP = CInt(txt_LP_Base.Text)
            finalLP = baseLP
        Else
            finalLP = ws.Cells(iRow - 1, 7).Value + txt_LP.Value
        End If
    End If
    
    ws.Cells(iRow, 7).Value = finalLP

    'Column H - Kills
    ws.Cells(iRow, 8).Value = txt_Kills.Value

    'Column I - Deaths
    ws.Cells(iRow, 9).Value = txt_Deaths.Value

    'Column J - Assists
    ws.Cells(iRow, 10).Value = txt_Assists.Value

    'Column K - Ratio
    Dim ratio As Double
    ratio = (CDbl(txt_Kills.Value) + CDbl(txt_Assists.Value)) / CDbl(txt_Deaths.Value)
    ws.Cells(iRow, 11).Value = ratio

    'Column L - CS
    ws.Cells(iRow, 12).Value = txt_CS.Value

    'Column M - CS/m
    Dim tim As Double
    tim = CDbl(txt_Length_S.Value) / 60
    tim = CDbl(txt_Length_M.Value) + tim
    
    ws.Cells(iRow, 13).Value = CInt(txt_CS.Value) / tim
    
    'Column N - Gold
    ws.Cells(iRow, 14).Value = txt_Gold.Value

    'Column O - Gold/m
    ws.Cells(iRow, 15).Value = CInt(txt_Gold.Value) / tim
    
    'Column P - Screenshot
    ws.Cells(iRow, 16).Value = txt_Screenshot.Value

    'Column Q - Grade
    ws.Cells(iRow, 17).Value = UCase(txt_Grade.Text)

    'Column R - Role
    ws.Cells(iRow, 18).Value = cmb_Role.Value

    'Column S - Champion
    ws.Cells(iRow, 19).Value = cmb_Champ.Value

    'Column T - Lane Opponent
    ws.Cells(iRow, 20).Value = cmb_Opp.Value

    'Column U - W/L Lane
    ws.Cells(iRow, 21).Value = txt_WL.Value

    'Column V - Problems
    ws.Cells(iRow, 22).Value = txt_Prob.Value

    'Column W - Other Comments
    ws.Cells(iRow, 23).Value = txt_Other.Value
    
    If txt_Screenshot.Enabled = True Then
        txt_Screenshot.SetFocus
    Else
        txt_Length_M.SetFocus
    End If
    
    If chk_Submit_Clear.Value = True Then
        btn_Clear_Click
    End If
    
    If chk_Submit_Close.Value = True Then
        btn_Cancel_Click
    End If
End Sub

Public Sub chk_Screenshot_Click()
    If chk_Screenshot.Value = True Then
        txt_Screenshot.Enabled = True
        txt_Screenshot.BackColor = txt_Length_M.BackColor
        txt_Screenshot.TabStop = True
    Else
        txt_Screenshot.Enabled = False
        txt_Screenshot.BackColor = Me.BackColor
        txt_Screenshot.TabStop = False
    End If
End Sub

Public Sub txt_Assists_Change()
' Only allows numerical values

        Dim inp As String
        Dim subStr As String
        Dim index As Integer
        inp = txt_Assists.Text
        index = txt_Assists.SelStart

        If index <> 0 Then
            subStr = Mid(txt_Assists.Text, index, 1)
            If InStr(1, allowedChars, subStr, vbTextCompare) = 0 Then
                inp = Replace(inp, subStr, "")
                txt_Assists.Text = inp
                If index <> 0 Then
                    txt_Assists.SelStart = (index - 1)
                End If
            End If
        End If
End Sub

Public Sub txt_Gold_Change()
' Only allows numerical values and k

        Dim inp As String
        Dim subStr As String
        Dim index As Integer
        inp = txt_Gold.Text
        index = txt_Gold.SelStart

        If index <> 0 Then
            subStr = Mid(txt_Gold.Text, index, 1)
            If InStr(1, allowedCharsK, subStr, vbTextCompare) = 0 Then
                inp = Replace(inp, subStr, "")
                txt_Gold.Text = inp
                If index <> 0 Then
                    txt_Gold.SelStart = (index - 1)
                End If
            End If
        End If
End Sub

Public Sub txt_CS_Change()
' Only allows numerical values

        Dim inp As String
        Dim subStr As String
        Dim index As Integer
        inp = txt_CS.Text
        index = txt_CS.SelStart

        If index <> 0 Then
            subStr = Mid(txt_CS.Text, index, 1)
            If InStr(1, allowedChars, subStr, vbTextCompare) = 0 Then
                inp = Replace(inp, subStr, "")
                txt_CS.Text = inp
                If index <> 0 Then
                    txt_CS.SelStart = (index - 1)
                End If
            End If
        End If
End Sub

Public Sub txt_Deaths_Change()
' Only allows numerical values

        Dim inp As String
        Dim subStr As String
        Dim index As Integer
        inp = txt_Deaths.Text
        index = txt_Deaths.SelStart

        If index <> 0 Then
            subStr = Mid(txt_Deaths.Text, index, 1)
            If InStr(1, allowedChars, subStr, vbTextCompare) = 0 Then
                inp = Replace(inp, subStr, "")
                txt_Deaths.Text = inp
                If index <> 0 Then
                    txt_Deaths.SelStart = (index - 1)
                End If
            End If
        End If
End Sub

Public Sub txt_Kills_Change()
' Only allows numerical values

        Dim inp As String
        Dim subStr As String
        Dim index As Integer
        inp = txt_Kills.Text
        index = txt_Kills.SelStart

        If index <> 0 Then
            subStr = Mid(txt_Kills.Text, index, 1)
            If InStr(1, allowedChars, subStr, vbTextCompare) = 0 Then
                inp = Replace(inp, subStr, "")
                txt_Kills.Text = inp
                If index <> 0 Then
                    txt_Kills.SelStart = (index - 1)
                End If
            End If
        End If
End Sub

Public Sub txt_LP_Change()
' Only allows numerical values and negative

        Dim inp As String
        Dim subStr As String
        Dim index As Integer
        inp = txt_LP.Text
        index = txt_LP.SelStart

        If index <> 0 Then
            subStr = Mid(txt_LP.Text, index, 1)
            If InStr(1, allowedCharsNeg, subStr, vbTextCompare) = 0 Then
                inp = Replace(inp, subStr, "")
                txt_LP.Text = inp
                If index <> 0 Then
                    txt_LP.SelStart = (index - 1)
                End If
            End If
        End If
End Sub

Public Sub txt_LP_Base_Change()
' Only allows numerical values and negative

        Dim inp As String
        Dim subStr As String
        Dim index As Integer
        inp = txt_LP_Base.Text
        index = txt_LP_Base.SelStart

        If index <> 0 Then
            subStr = Mid(txt_LP_Base.Text, index, 1)
            If InStr(1, allowedCharsNeg, subStr, vbTextCompare) = 0 Then
                inp = Replace(inp, subStr, "")
                txt_LP_Base.Text = inp
                If index <> 0 Then
                    txt_LP_Base.SelStart = (index - 1)
                End If
            End If
        End If
End Sub

Private Sub UserForm_Initialize()
    If txt_Screenshot.Enabled = True Then
        txt_Screenshot.SetFocus
    Else
        txt_Length_M.SetFocus
    End If
End Sub
