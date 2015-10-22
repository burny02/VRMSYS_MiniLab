Imports TemplateDB

Module Variables
    Public OverClass As OverClass
    Private Const TablePath As String = "C:\OK\Databases\Backup\Backend2.accdb"
    Private Const PWord As String = "Crypto*Dave02"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Private Const UserTable As String = "[Users]"
    Private Const UserField As String = "Username"
    Private Const LockTable As String = "[Locker]"
    Private Const ActiveUsersTable As String = "[ActiveUsers]"
    Private Contact As String = "Mustafa Dawood"
    Public Const SolutionName As String = "VRMSYS - Mini Lab"
    Public SiteForm As Site
    Public PickCohort As Long
    Public AppID As Long

    Public Function GetTheConnection() As String
        GetTheConnection = Connect2
    End Function


    Public Sub StartUp(WhichForm As Form)

        OverClass = New TemplateDB.OverClass
        OverClass.SetPrivate(UserTable, _
                           UserField, _
                           LockTable, _
                           Contact, _
                           Connect2,
                           ActiveUsersTable)

        OverClass.LockCheck()

        OverClass.LoginCheck()

        OverClass.AddAllDataItem(WhichForm)

        For Each ctl In OverClass.DataItemCollection
            If (TypeOf ctl Is ComboBox) Then
                Dim Com As ComboBox = ctl
                AddHandler Com.SelectionChangeCommitted, AddressOf GenericCombo
            End If
        Next
        For Each ctl In OverClass.DataItemCollection
            If (TypeOf ctl Is Button) Then
                Dim But As Button = ctl
                AddHandler But.Click, AddressOf ButtonSpecifics
            End If
        Next


    End Sub

End Module
