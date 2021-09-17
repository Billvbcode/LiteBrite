Public Class Form1
    Dim PLabel(5000) As Label           ' Draw Picture
    Dim PColor(5000) As Label           ' Color Labels
    Dim r(30) As Integer                      ' Red Color Numbers
    Dim g(30) As Integer                      ' Green Color Numbers
    Dim b(30) As Integer                    ' Blue Color Numbers
    Dim bAuto As Boolean                 ' Auto Draw On/Off
    Dim bBlack As Boolean                 ' Black On/Off
    Dim value As String = My.Application.Info.DirectoryPath & "\”
    Dim pName As String                  'PictureName
    Dim ipCnt As Integer                     'Picture Count
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Icnt As Integer
        Dim i As Integer
        Dim j As Integer
        Icnt = 1                                       'Lite Bright Label counter
        For j = 1 To 20                          ' Vertical
            For i = 1 To 20 '                    ‘Horizontal
                PLabel(Icnt) = New System.Windows.Forms.Label    ' Lite Bright  Label
                With PLabel(Icnt)
                    .Name = Icnt                                 'Label Name
                    .Text = "-"
                    .Tag = "0"
                    .Font = New Font("Times New Roman", 10, FontStyle.Bold)
                    .Size = New Size(22, 22)
                    .Location = New System.Drawing.Point(24 + (i - 1) * 22,
                                                         6 + (j - 1) * 22)
                End With
                Me.Controls.Add(PLabel(Icnt)) '  This is the create the Lite Bright Label
                AddHandler PLabel(Icnt).Click, AddressOf PLabel_Click 'Click Event
                AddHandler PLabel(Icnt).MouseMove, AddressOf Plabel_MouseMove 'MouseMove Event
                Icnt = Icnt + 1                                                    ' Increase Label counter
            Next
        Next
        ' Private Sub Form1_Load (Page 2) Add Color Labels
        Icnt = 1                                                                  ' Color Label counter
        For j = 1 To 13    ' Vertical
            PColor(Icnt) = New System.Windows.Forms.Label    'Color Label
            With PColor(Icnt)
                .Name = "C" & Icnt                                 'Color Label Name
                .Text = Chr(64 + j)
                .Tag = Chr(64 + j)
                .Font = New Font("Times New Roman", 10, FontStyle.Bold)

                .Size = New Size(22, 22)
                .Location = New System.Drawing.Point(500, 6 + (j - 1) * 22)
            End With
            Me.Controls.Add(PColor(Icnt))          '  This is the create the Color label
            AddHandler PColor(Icnt).Click, AddressOf PColor_Click 'Color Click Event
            Icnt = Icnt + 1                                          ' Increase Color Label counter
        Next
        'Private Sub Form1_Load (Page 3)
        For j = 1 To 13                                 ' 2nd Row Vertical of Color Labels
            PColor(Icnt) = New System.Windows.Forms.Label    'Color Label
            With PColor(Icnt)
                .Name = "C" & Icnt                                 'Color Label Name
                .Text = Chr(64 + j + 13)
                .Tag = Chr(64 + j + 13)
                .Font = New Font("Times New Roman", 10, FontStyle.Bold)

                .Size = New Size(22, 22)
                .Location = New System.Drawing.Point(532, 6 + (j - 1) * 22)
            End With
            Me.Controls.Add(PColor(Icnt)) '  This is the create the Color label
            AddHandler PColor(Icnt).Click, AddressOf PColor_Click 'Color Click Event
            Icnt = Icnt + 1                                      ' Increase Color Label counter
        Next

        'Private Sub Form1_Load (Page 4)
        ' Assign Colors to Labels  (R=Red   G=Green   B=Blue)
        r = {255, 255, 255, 255, 153, 255, 255, 255, 255, 255, 165, 255, 210, 0,
            70, 138, 128, 50, 128, 152, 192, 128, 0, 0}
        g = {255, 255, 215, 255, 50, 128, 105, 0, 140, 165, 42, 228, 105, 0, 130,
            206, 255, 205, 255, 151, 192, 128, 0, 0}
        b = {255, 224, 0, 128, 204, 255, 180, 0, 0, 0, 42, 196, 30, 205, 180, 250,
            255, 50, 128, 5, 192, 128, 0, 0}
        For i = 1 To 24             'Assign Colors To PColor Labels
            PColor(i).BackColor = Color.FromArgb(r(i - 1), g(i - 1), b(i - 1))
        Next
        PColor(5).ForeColor = Color.White                 'E - White
        PColor(11).ForeColor = Color.White               'K - White
        PColor(14).ForeColor = Color.White                'N - White
        PColor(15).ForeColor = Color.White                'O - White
        PColor(23).ForeColor = Color.White                'W - White
        PColor(24).BackColor = DefaultBackColor      ' X - DefaultColor
        LblColor.BackColor = PColor(1).BackColor       'Copy Color to LblColor
        LblColor.Tag = PColor(1).Tag                               'Copy Tag to LblColor Tag
        bAuto = False                                                       'Auto Draw Off
        PColor(25).Text = "Y (Auto Off)“                        ' Y = Auto Draw On/Off
        PColor(25).BackColor = Color.Tomato
        PColor(25).AutoSize = True
        PColor(26).Text = "Z (Load/Save)"                       'Z = Load/Save
        PColor(26).AutoSize = True
        'Private Sub Form1_Load (Page 5)
        Me.Text = "1-Clear  2-#   3-Black   4-Move   5-Flash    6-Next Picture" 'Instructions
        GetFiles()            'Call Get File Subroutine to Read All Txt Files
        ipCnt = -1
        pName = “”             'Picture Name
        TextBox1.Text = pName
        Panel1.Enabled = False   'Disable Panel
        ListBox1.Enabled = False 'Disable ListBox1
        BtnSave.Enabled = False 'Disable BtnSave
        Panel1.Visible = False     'Hide Panel1
    End Sub
    'PLabel Click (Lite Brite Labels)
    Private Sub PLabel_Click(ByVal sender As Object, ByVal e As EventArgs)
        sender.BackColor = LblColor.BackColor
        sender.Tag = LblColor.Tag
    End Sub

    'PLabel Mouse Move (Lite Brite Labels)
    Private Sub Plabel_MouseMove(sender As Object, e As MouseEventArgs)
        Dim i As Integer
        If sender.text >= "A" Then 'Enter Letter to Change Text to Color
            i = Asc(sender.text) - 64
            sender.BackColor = PColor(i).BackColor
            sender.Tag = PColor(i).Tag
            sender.Text = "-"
            Exit Sub
        End If
        If bAuto Then 'Auto Draw
            sender.BackColor = LblColor.BackColor
            sender.Tag = LblColor.Tag
        End If
    End Sub
    'PColor Click (Change the Paint Color You Use )
    Private Sub PColor_Click(ByVal sender As Object, ByVal e As EventArgs)
        If sender.tag > "X" Then Exit Sub 'Only For Letters A-X
        LblColor.BackColor = sender.BackColor
        LblColor.Tag = sender.Tag
        LblColor.ForeColor = sender.foreColor
    End Sub
    'Get Files with .txt extension 
    Private Sub GetFiles()
        Dim di As IO.DirectoryInfo = New IO.DirectoryInfo(value)
        ListBox1.Items.Clear()  'Clear ListBox
        Dim diar1 As IO.FileInfo() = di.GetFiles("*.txt")
        Dim dra As IO.FileInfo
        'list the names of all files in the specified directory
        For Each dra In diar1
            ListBox1.Items.Add(dra.Name)
        Next
    End Sub

    Private Sub TimMove_Tick(sender As Object, e As EventArgs) Handles TimMove.Tick
        'Move the Picture
        Dim Icnt As Integer
        Dim i As Integer
        Dim j As Integer
        Icnt = 1                      ' Label counter
        For j = 1 To 20           ' Vertical
            'LblSwitch – Save The first Color and Tag
            LblSwitch.BackColor = PLabel(Icnt).BackColor
            LblSwitch.Tag = PLabel(Icnt).Tag
            '1 – 19 Pick up Color and Tag Of Next Label
            For i = 1 To 19           ' Horizontal
                PLabel(Icnt).BackColor = PLabel(Icnt + 1).BackColor
                PLabel(Icnt).Tag = PLabel(Icnt + 1).Tag
                Icnt = Icnt + 1 ' Increase Label counter
            Next
            '20 Pick Up Color and Tag of Saved Label
            PLabel(Icnt).BackColor = LblSwitch.BackColor
            PLabel(Icnt).Tag = LblSwitch.Tag
            Icnt = Icnt + 1           ' Increase Label counter
        Next
    End Sub
    'TimNewColor – Change The Color Of The Picture
    Private Sub TimNewColor_Tick(sender As Object, e As EventArgs) Handles TimNewColor.Tick
        Dim i As Integer
        Dim j As Integer
        For i = 1 To 400
            If PLabel(i).Tag >= "A" Then 'Change Label  Color
                j = Asc(PLabel(i).Tag) - 65 + 3
                If j > 23 Then j -= 23
                PLabel(i).BackColor = PColor(j).BackColor
                PLabel(i).Tag = PColor(j).Tag
            End If
        Next
    End Sub
    'BtnSave – Save Picture file
    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        If TextBox1.Text = "" Then
            TextBox1.Focus()
            Exit Sub
        End If
        Dim value As String = My.Application.Info.DirectoryPath & "\”
        Dim sFileName As String
        sFileName = value & TextBox1.Text & ".txt" 'File Name
        Dim srFileWrite As New System.IO.StreamWriter(sFileName) 'Open File
        Dim sMap As String ' Save Map Information
        Dim i As Integer
        sMap = ""          ' Blank Map
        For i = 1 To 400   ' For all 400 PictureBoxes
            sMap = sMap & PLabel(i).Tag ' Create Map String
        Next
        srFileWrite.WriteLine(sMap) 'Write Map String as a File
        srFileWrite.Close()         'Close File
        GetFiles()
        Panel1.Enabled = False  'Disable Panel1
        ListBox1.Enabled = False 'Disable Listbox1
        BtnSave.Enabled = False  'Disable Button1
        Panel1.Visible = False     'Hide Panel1
    End Sub
    'Read Picture File
    Private Sub ReadColorFile()
        Dim i As Integer
        Dim j As Integer
        Dim sMap As String
        Dim sLet As String
        Dim srFileReader As System.IO.StreamReader
        Dim sFileName As String
        sFileName = value & pName & ".txt"                'File Name
        srFileReader = System.IO.File.OpenText(sFileName) 'Open File
        sMap = srFileReader.ReadLine()    'Read File
        srFileReader.Close()              'Close File     
        For j = 1 To 400
            sLet = UCase(Mid(sMap, j, 1))
            PLabel(j).Text = "-"
            PLabel(j).Tag = "0"
            PLabel(j).BackColor = DefaultBackColor
            If sLet >= "A" And sLet <= "X" Then 'Paint Change Text to Color
                i = Asc(sLet) - 64
                PLabel(j).BackColor = PColor(i).BackColor
                PLabel(j).Tag = PColor(i).Tag
            End If
        Next
        TextBox1.Text = pName 'Set Textbox to Picture Name
        Panel1.Enabled = False  'Disable Panel1
        ListBox1.Enabled = False 'Disable Listbox1
        BtnSave.Enabled = False  'Disable Button1
        Panel1.Visible = False     'Hide Panel1
    End Sub
    'ListBox1 – Code To Select A Picture File
    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        If ListBox1.Text = "" Then Exit Sub 'Exit if No Text is Selected
        pName = Mid(ListBox1.Text, 1, Len(ListBox1.Text) - 4)
        ReadColorFile()
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Dim i As Integer
        '1=Clear Form
        If e.KeyCode = Keys.D1 Then  '1=Clear Form
            For i = 1 To 400
                PLabel(i).BackColor = DefaultBackColor
                PLabel(i).Text = "-"
                PLabel(i).Tag = "0"
            Next
            bAuto = False     'Auto Draw Off
            PColor(25).Text = "Y (Auto Off)"
            PColor(25).BackColor = Color.Tomato
            pName = "" 'Erase Picture Name
            Exit Sub
        End If
        '2=Turn Picture To Letters
        If e.KeyCode = Keys.D2 Then
            For i = 1 To 400
                If PLabel(i).Tag >= "A" Then
                    PLabel(i).Text = PLabel(i).Tag
                    PLabel(i).BackColor = DefaultBackColor
                End If
            Next
            bAuto = False        'Auto Draw Off
            PColor(25).Text = "Y (Auto Off)"
            PColor(25).BackColor = Color.Tomato
            Exit Sub
        End If
        '3=Black or Default Color
        If e.KeyCode = Keys.D3 Then
            If bBlack Then
                bBlack = False  ' Black Off
            Else
                bBlack = True   ' Black On
            End If
            For i = 1 To 400
                If PLabel(i).Tag = "0" Then
                    If bBlack Then
                        PLabel(i).BackColor = Color.Black
                    Else
                        PLabel(i).BackColor = DefaultBackColor
                    End If
                End If
            Next
            Exit Sub
        End If
        '4=Move Picture
        If e.KeyCode = Keys.D4 Then
            If TimMove.Enabled Then
                TimMove.Stop() ' Stop Timer
            Else
                TimMove.Start() ' Start Timer
            End If
            Exit Sub
        End If
        '5=New Label Color
        If e.KeyCode = Keys.D5 Then
            If TimNewColor.Enabled Then
                TimNewColor.Stop() ' Stop Timer
            Else
                TimNewColor.Start() ' Start Timer
            End If
            Exit Sub
        End If
        '6=Read Next Picture
        If e.KeyCode = Keys.D6 Then
            If ListBox1.Items.Count = 0 Then Exit Sub
            ipCnt += 1 'Increase Picture Counter
            If ipCnt >= ListBox1.Items.Count Then ipCnt = 0
            pName = Mid(ListBox1.Items(ipCnt).ToString, 1, Len(ListBox1.Items(ipCnt).ToString) - 4)
            ReadColorFile() 'Call ReadColorFile Subroutine
            Exit Sub
        End If
        '
        '  A - X Change The Color
        If e.KeyCode >= Keys.A And e.KeyCode <= Keys.X Then
            i = e.KeyValue - 64
            LblColor.BackColor = PColor(i).BackColor
            LblColor.Tag = PColor(i).Tag
            LblColor.ForeColor = PColor(i).ForeColor
        End If
        ' Y= Auto Draw On/Off
        If e.KeyCode = Keys.Y Then ' Y= Auto Draw On/Off
            If bAuto Then
                bAuto = False    ' Auto Draw Off
                PColor(25).Text = "Y (Auto Off)"
                PColor(25).BackColor = Color.Tomato
            Else
                bAuto = True      ' Auto Draw On
                PColor(25).Text = "Y (Auto On)"
                PColor(25).BackColor = Color.LightGreen
            End If
        End If
        ' Z= Load/Save
        If e.KeyCode = Keys.Z Then '
            GetFiles()
            TextBox1.Text = pName 'Set Textbox to Picture Name
            Panel1.Enabled = True   'Enable Panel1
            ListBox1.Enabled = True 'Enable Listbox1
            BtnSave.Enabled = True  'Enable BtnSave
            Panel1.Visible = True     'Show Panel1
        End If

    End Sub
End Class
