Imports System.Windows.Forms

Public Class ProjectRangeHelper

    Private BS As BindingSource
    Private SBUBS As BindingSource
    Private FamilyBS As BindingSource
    Private SPMBS As BindingSource
    Private PMBS As BindingSource
    Private ProjectBS As BindingSource
    Private RangeBS As BindingSource

    Dim drv As DataRowView

    Public Sub New(BS As BindingSource, SBUBS As BindingSource, FamilyBS As BindingSource, SPMBS As BindingSource, PMBS As BindingSource, ProjectBS As BindingSource, RangeBS As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()
       
        ' Add any initialization after the InitializeComponent() call.
        Me.BS = BS
        drv = BS.Current
        Me.SBUBS = SBUBS
        Me.FamilyBS = FamilyBS
        Me.SPMBS = SPMBS
        Me.PMBS = PMBS
        Me.ProjectBS = ProjectBS
        Me.RangeBS = RangeBS
        initData()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim currDrv As DataRowView = ComboBox1.SelectedItem
        If Not IsNothing(currDrv) Then
            drv.Row.Item("sbuid") = currDrv.Row.Item("sbuid")
            drv.Row.Item("sbuname") = currDrv.Row.Item("sbuname")
        End If
        currDrv = ComboBox2.SelectedItem
        If Not IsNothing(currDrv) Then
            drv.Row.Item("familyid") = currDrv.Row.Item("familyid")
            drv.Row.Item("familyname") = currDrv.Row.Item("familyname")
        End If
        currDrv = ComboBox3.SelectedItem
        If Not IsNothing(currDrv) Then
            drv.Row.Item("ssmid") = currDrv.Row.Item("ssmid")
            drv.Row.Item("ssm") = currDrv.Row.Item("ssmname")
        End If
        currDrv = ComboBox4.SelectedItem        
        If Not IsNothing(currDrv) Then
            drv.Row.Item("spmid") = currDrv.Row.Item("pmid")
            drv.Row.Item("spm") = currDrv.Row.Item("pmname")
        End If
        currDrv = ComboBox5.SelectedItem
        If Not IsNothing(currDrv) Then
            drv.Row.Item("pcprojectid") = currDrv.Row.Item("pcprojectid")
            drv.Row.Item("projectname") = currDrv.Row.Item("projectname")
        End If
        currDrv = ComboBox6.SelectedItem     
        If Not IsNothing(currDrv) Then
            drv.Row.Item("pcrangeid") = currDrv.Row.Item("pcrangeid")
            drv.Row.Item("rangename") = currDrv.Row.Item("rangename")
        End If

        drv.EndEdit()

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub


    Private Sub initData()

        ComboBox1.DataSource = SBUBS
        ComboBox1.DisplayMember = "sbuname"
        ComboBox1.ValueMember = "sbuid"

        ComboBox2.DataSource = FamilyBS
        ComboBox2.DisplayMember = "familyname"
        ComboBox2.ValueMember = "familyid"

        ComboBox3.DataSource = SPMBS
        ComboBox3.DisplayMember = "ssmname"
        ComboBox3.ValueMember = "ssmid"

        ComboBox4.DataSource = PMBS
        ComboBox4.DisplayMember = "pmname"
        ComboBox4.ValueMember = "pmid"

        ComboBox5.DataSource = ProjectBS
        ComboBox5.DisplayMember = "projectname"
        ComboBox5.ValueMember = "pcprojectid"

        ComboBox6.DataSource = RangeBS
        ComboBox6.DisplayMember = "rangename"
        ComboBox6.ValueMember = "pcrangeid"

    End Sub

    Private Sub ProjectRangeHelper_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Locate manually
        Dim pos = SBUBS.Find("sbuid", drv.Row.Item("sbuid"))
        ComboBox1.SelectedIndex = pos

        pos = FamilyBS.Find("familyid", drv.Row.Item("familyid"))
        ComboBox2.SelectedIndex = pos

        pos = SPMBS.Find("ssmid", drv.Row.Item("ssmid"))
        ComboBox3.SelectedIndex = pos

        pos = PMBS.Find("pmid", drv.Row.Item("spmid"))
        ComboBox4.SelectedIndex = pos

        pos = ProjectBS.Find("pcprojectid", drv.Row.Item("pcprojectid"))
        ComboBox5.SelectedIndex = pos

        pos = RangeBS.Find("pcrangeid", drv.Row.Item("pcrangeid"))
        ComboBox6.SelectedIndex = pos


    End Sub

    Private Sub ComboBox5_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ComboBox5.Validating, ComboBox6.Validating
        Dim cb As ComboBox = DirectCast(sender, ComboBox)
        If IsNothing(cb.SelectedItem) Then
            Dim currDrv = BS.Current
            If Not IsNothing(currDrv) Then
                Select Case cb.Name
                    Case "ComboBox5" 'Project
                        Dim ProjectName As String = ComboBox5.Text
                        Dim drv As DataRowView = ProjectBS.AddNew
                        drv.Row.Item("projectname") = ProjectName
                        currDrv.Row.Item("projectname") = drv.Row.Item("projectname")          
                        drv.EndEdit()
                        Dim mypos = ProjectBS.Find("pcprojectid", drv.Row.Item("pcprojectid"))
                        ComboBox5.SelectedIndex = mypos

                    Case "ComboBox6" 'Range
                        Dim RangeName As String = ComboBox6.Text
                        Dim drv As DataRowView = RangeBS.AddNew
                        drv.Row.Item("rangename") = RangeName
                        drv.EndEdit()
                        currDrv.Row.Item("rangename") = drv.Row.Item("rangename")
                        Dim mypos = RangeBS.Find("pcrangeid", drv.Row.Item("pcrangeid"))
                        ComboBox6.SelectedIndex = mypos
                End Select
            End If

        End If
    End Sub

    Private Sub AddToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Dim cb As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Dim mycombo = cb.GetCurrentParent
        Dim abc As ContextMenuStrip = DirectCast(mycombo, ContextMenuStrip)
        Dim bcd As ComboBox = DirectCast(abc.SourceControl, ComboBox)
        If Not IsNothing(bcd.SelectedItem) Then
            Dim currDrv = BS.Current
            Select Case bcd.Name
                Case "ComboBox5" 'Project
                    ProjectBS.RemoveCurrent()
                    currDrv.Row.Item("projectname") = ""
                Case "ComboBox6" 'Range
                    RangeBS.RemoveCurrent()
            End Select           
        End If
    End Sub


   
End Class
