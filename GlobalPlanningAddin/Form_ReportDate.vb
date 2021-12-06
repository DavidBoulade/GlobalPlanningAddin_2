Public Class Form_ReportDate

    Public Property SelectedDate As Date
    Public Property WasCancelled As Boolean = False

    Sub New(DefaultDate As Date)
        InitializeComponent() ' Cet appel est requis par le concepteur.
        CommonNew(DefaultDate)
    End Sub

    Sub New(DefaultDate As Date, BoldedDates As Date())
        InitializeComponent() ' Cet appel est requis par le concepteur.
        CommonNew(DefaultDate)
        If Not (BoldedDates Is Nothing) Then MonthCalendar1.BoldedDates = BoldedDates
    End Sub

    Private Sub CommonNew(DefaultDate As Date)
        _SelectedDate = DefaultDate
        MonthCalendar1.SelectionStart = _SelectedDate
        MonthCalendar1.MaxDate = Today()
    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        _SelectedDate = MonthCalendar1.SelectionStart
        Me.Close()
    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        _WasCancelled = True
        Me.Close()
    End Sub

    Private Sub Form_ReportDate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Globals.CenterForm(Me) 'center the form on the Excel Window
    End Sub
End Class