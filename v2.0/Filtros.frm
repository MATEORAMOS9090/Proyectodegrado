VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8640
   LinkTopic       =   "Form3"
   Picture         =   "Filtros.frx":0000
   ScaleHeight     =   5130
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Buscar por"
      Height          =   2295
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   3495
      Begin VB.OptionButton Option5 
         Caption         =   "Ordenar stock descendente"
         Height          =   555
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Tipo de producto seleccionado"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   2655
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Ordenar stock ascendente"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Menos prendas en stock"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   3495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Más prendas en stock"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   3495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mayor vendido"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   2160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Text            =   "Seleccionar..."
   End
   Begin VB.Label la 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8160
      TabIndex        =   7
      Top             =   4080
      Width           =   45
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub DataCombo1_Change()
    'If Len(DataCombo1.Text) = 0 Or Len(DataCombo1.Text) = 1 Then DataCombo1.Text = "Seleccionar...": Exit Sub
    CTTP
    With TTP
        .Find "Descripción='" & Trim(DataCombo1.BoundText) & "'"
        Label4.Caption = !Id_Tp
    End With
    CTP
    With TP
        If .State = 1 Then .Close
        X = Label4.Caption
        .Open "select * from Producto where [Id_TP_FK]like '" & X & "'", base, adOpenStatic, adLockBatchOptimistic
        Form1.invicible
        For i = 0 To 6
            If .EOF Or .BOF Then Exit Sub
            If i = 0 Then
                .MoveFirst
            Else
                .MoveNext
            End If
            If .EOF Or .BOF Then Exit Sub
            If Trim(!URL) = "" Then
                Form1.Image1(i).Picture = LoadPicture("C:\Proyecto\final\img\nimg.jpg")
            Else
                Y = App.Path
                Form1.Image1(i).Picture = LoadPicture(Y & "\img\" & Trim(!URL))
            End If
            Form1.Label4(i).Caption = !Etiqueta
            Form1.Label6(i).Caption = !Id_Producto
            Form1.Image1(i).Visible = True
            Form1.Label4(i).Visible = True
        Next i
        Form1.Label7.Caption = !Id_Producto
    End With
    CTTP
    Set DataCombo1.RowSource = TTP
    DataCombo1.BoundColumn = "Descripción"
    DataCombo1.ListField = "Descripción"
End Sub

Private Sub DataCombo1_Click(Area As Integer)
    CTTP
    Set DataCombo1.RowSource = TTP
    DataCombo1.BoundColumn = "Descripción"
    DataCombo1.ListField = "Descripción"
    Option4.Value = False
    Option6.Value = False
End Sub

Private Sub Form_Load()
    CTTP
    Set DataCombo1.RowSource = TTP
    DataCombo1.BoundColumn = "Descripción"
    DataCombo1.ListField = "Descripción"
End Sub

Sub bus()
    With TP
        If .State = 1 Then .Close
        X = Label4.Caption
        .Open "select * from Producto where [Id_TP_FK]like '" & X & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
End Sub

Sub bus1()
    With TP
        Form1.invicible
        .MoveFirst
        If .EOF Or .BOF Then Exit Sub
        For i = 0 To 6
            If Trim(!URL) = "" Then
                Form1.Image1(i).Picture = LoadPicture("C:\Proyecto\final\img\nimg.jpg")
            Else
                Y = App.Path
                Form1.Image1(i).Picture = LoadPicture(Y & "\img\" & Trim(!URL))
            End If
            Form1.Label4(i).Caption = !Etiqueta
            Form1.Label6(i).Caption = !Id_Producto
            Form1.Image1(i).Visible = True
            Form1.Label4(i).Visible = True
            .MoveNext
            If .EOF Or .BOF Then Exit Sub
        Next i
        Form1.Label7.Caption = !Id_Producto
    End With
End Sub

Private Sub Option1_Click()
    With DFact
        If .State = 1 Then .Close
        .Open "select * from Detalle_Factura ORDER BY Detalle_Factura.Id_P_FK", base, adOpenStatic, adLockBatchOptimistic
        .MoveFirst
        Y = 0
        X1 = !Id_P_FK
        For i = 1 To .RecordCount
            If X1 = !Id_P_FK Then
                X = X + 1
            Else
                If Y = 0 Then Y1 = !Id_P_FK
                If Y1 = !Id_P_FK Then
                    Y = Y + 1
                Else
                    If X > Y Then
                        Y = 0
                    Else
                        X1 = Y1
                        X = Y
                        Y = 0
                    End If
                End If
            End If
            .MoveNext
        Next i
        If .State = 1 Then .Close
        .Open "select * from Detalle_Factura where [Id_P_FK]like '" & X1 & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    With TP
        If .State = 1 Then .Close
        .Open "select * from Producto where [Id_Producto]like '" & X1 & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    bus1
End Sub

Private Sub Option2_Click()
    With TP
        If .State = 1 Then .Close
        .Open "select * from Producto ORDER BY Producto.[Cantidad] DESC", base, adOpenStatic, adLockBatchOptimistic
        X = !Cantidad
        If .State = 1 Then .Close
        .Open "select * from Producto where [Cantidad]like '" & X & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    bus1
End Sub

Private Sub Option3_Click()
    With TP
        If .State = 1 Then .Close
        .Open "select * from Producto ORDER BY Producto.[Cantidad]", base, adOpenStatic, adLockBatchOptimistic
        X = !Cantidad
        If .State = 1 Then .Close
        .Open "select * from Producto where [Cantidad]like '" & X & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    bus1
End Sub

Private Sub Option4_Click()
    With TP
        If .State = 1 Then .Close
        .Open "select * from Producto ORDER BY Producto.[Cantidad]", base, adOpenStatic, adLockBatchOptimistic
        Set DataReport3.DataSource = TP
        DataReport3.Show
    End With
End Sub



Private Sub Option5_Click()
    With TP
        If .State = 1 Then .Close
        .Open "select * from Producto ORDER BY Producto.[Cantidad] DESC", base, adOpenStatic, adLockBatchOptimistic
        Set DataReport3.DataSource = TP
        DataReport3.Show
    End With
End Sub

Private Sub Option6_Click()
    With TTP
        .Find "Descripción='" & Trim(DataCombo1.Text) & "'"
        If .EOF Then Exit Sub
        la.Caption = !Id_Tp
    End With
    With TP
        If .State = 1 Then .Close
        .Open "select * from Producto Where Producto.[Id_TP_FK]LIKE '" & la.Caption & "'", base, adOpenStatic, adLockBatchOptimistic
        Set DataReport3.DataSource = TP
        DataReport3.Show
    End With
End Sub
