'********************************************************************************************
'* Name:       Aiman Haroon
'* Class:      CIS-1510
'* Assignment: Project 02 Fall 2020
'* File:       frmMainDrugs.vb
'* Purpose:    Pharmacy Drug Inventory Software for Create,Read,and edit items
'********************************************************************************************


Option Strict On
Option Explicit On
Option Infer Off
Option Compare Binary

Imports System.Globalization
Imports System.IO
Imports System.Text

Public Class frmChangeDrugs

    Dim theDrug As Drugs   'used for existing and new drug
    Dim newDrugFlag As Boolean    'used to check if the drug to be added is new or not
    Private Const fileCategory As String = "category.txt"
    Private Const fileSupplier As String = "supplier.txt"


    Public Sub New(pDrug As Drugs, pNewDrugFlag As Boolean)
        theDrug = pDrug
        newDrugFlag = pNewDrugFlag

        InitializeComponent()

    End Sub

    Private Sub frmChangeDrugs_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        dtpDateRecieved.MaxDate = DateTime.Now  'default Max date to todays date
        dtpDateRecieved.Value = DateTime.Now ' 'default date of Today shows on the form


        'Reading from the lookup file1 (Category Listbox)

        Dim inFile As StreamReader
        Dim output As StringBuilder = New StringBuilder("")
        Dim wholeLine As String = ""


        If File.Exists(fileCategory) Then
            inFile = File.OpenText(fileCategory)

            Do Until inFile.Peek = -1

                wholeLine = inFile.ReadLine()
                lstCategory.Items.Add(wholeLine.ToString)
            Loop

            inFile.Close()
        End If

        'Reading from the lookup file2 (supplier dropdownlist)

        If File.Exists(fileSupplier) Then
            inFile = File.OpenText(fileSupplier)

            Do Until inFile.Peek = -1

                wholeLine = inFile.ReadLine()
                ddlSupplier.Items.Add(wholeLine.ToString)

            Loop
            inFile.Close()
        End If

        HideLblErrors()    'hide label errors
        mtbNDC.Select()    'set the focus to first Field 

        'for existing drug
        If Not newDrugFlag Then

            btnCancel.Location = btnClear.Location
            btnClear.Visible = False
            lblDateInfo.Visible = False

            'Converting the string to date type with matching dtp format
            Dim newDate As Date = DateTime.ParseExact(theDrug.dateReceived, "MM/dd/yyyy", CultureInfo.InvariantCulture)


            mtbNDC.Text = theDrug.NDC
            txtDrugName.Text = theDrug.drugName
            txtDosage.Text = theDrug.dosage
            txtDesc.Text = theDrug.desc
            txtUnitPrice.Text = CType(theDrug.unitPrice, String)
            txtPackagePrice.Text = CType(theDrug.packagePrice, String)
            lstCategory.SelectedItem = theDrug.category
            ddlSupplier.SelectedItem = theDrug.supplier
            txtCount.Text = CType(theDrug.count, String)
            dtpDateRecieved.Value = newDate
            txtUnitQuantity.Text = theDrug.unitQuantity


            If theDrug.drugType = radOTC.Text Then
                radOTC.Checked = True
            ElseIf theDrug.drugType = radControlled.Text Then
                radControlled.Checked = True
            ElseIf theDrug.drugType = radAntibiotic.Text Then
                radAntibiotic.Checked = True
            Else
                radAntidepressants.Checked = True

            End If

            If theDrug.inStock = 0 Then
                chkInStock.Checked = False
            Else
                chkInStock.Checked = True
            End If


        Else


            Me.Text = "Add New Drug (Aiman Haroon #73)"
            btnSave.Text = "Add New Drug"
            btnSave.AutoSize = True

        End If


    End Sub

    'Offensive Checking for Controls (Key Press Event)
    Private Sub txtDrugName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDrugName.KeyPress

        'Allow all uppercase & lowercase letters
        If (e.KeyChar >= "a" And e.KeyChar <= "z") Or (e.KeyChar >= "A" And e.KeyChar <= "Z") Then
            Return
        End If

        ' Allow the backspace and Spacebar
        If e.KeyChar = ControlChars.Back OrElse Asc(e.KeyChar) = Keys.Space Then
            Return
        End If

        ' Allow the these characters ('%', '/', ',' , '.', '-', '(', ')' )
        If e.KeyChar = "%" OrElse e.KeyChar = "/" OrElse e.KeyChar = "." OrElse e.KeyChar = "," _
         OrElse e.KeyChar = "(" OrElse e.KeyChar = ")" OrElse e.KeyChar = "-" Then
            Return
        End If

        e.Handled = True
    End Sub
    Private Sub txtDosage_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDosage.KeyPress

        'Allow numbers
        If e.KeyChar >= "0" AndAlso e.KeyChar <= "9" Then
            Return
        End If

        'Allow all uppercase & lowercase letters
        If (e.KeyChar >= "a" And e.KeyChar <= "z") Or (e.KeyChar >= "A" And e.KeyChar <= "Z") Then
            Return
        End If

        ' Allow the backspace and Spacebar
        If e.KeyChar = ControlChars.Back OrElse Asc(e.KeyChar) = Keys.Space Then
            Return
        End If

        ' Allow the %, /, .
        If e.KeyChar = "%" OrElse e.KeyChar = "/" OrElse e.KeyChar = "." Then
            Return
        End If

        e.Handled = True
    End Sub

    Private Sub txtDesc_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDesc.KeyPress

        'Allow all uppercase & lowercase letters
        If (e.KeyChar >= "a" And e.KeyChar <= "z") Or (e.KeyChar >= "A" And e.KeyChar <= "Z") Then
            Return
        End If

        'Allow numbers
        If e.KeyChar >= "0" AndAlso e.KeyChar <= "9" Then
            Return
        End If

        ' Allow the backspace and Spacebar
        If e.KeyChar = ControlChars.Back OrElse Asc(e.KeyChar) = Keys.Space Then
            Return
        End If

        ' Allow the these characters ('%', '/', ',' , '.', '-', '(', ')' )
        If e.KeyChar = "%" OrElse e.KeyChar = "/" OrElse e.KeyChar = "." OrElse e.KeyChar = "," _
         OrElse e.KeyChar = "(" OrElse e.KeyChar = ")" OrElse e.KeyChar = "-" Then
            Return
        End If

        e.Handled = True
    End Sub

    Private Sub txtUnitQuantity_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUnitQuantity.KeyPress

        'Allow all uppercase & lowercase letters
        If (e.KeyChar >= "a" And e.KeyChar <= "z") Or (e.KeyChar >= "A" And e.KeyChar <= "Z") Then
            Return
        End If

        'Allow numbers
        If e.KeyChar >= "0" AndAlso e.KeyChar <= "9" Then
            Return
        End If

        ' Allow the backspace and Spacebar
        If e.KeyChar = ControlChars.Back OrElse Asc(e.KeyChar) = Keys.Space Then
            Return
        End If

        e.Handled = True
    End Sub

    Private Sub txtPrices_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtUnitPrice.KeyPress, txtPackagePrice.KeyPress

        'Allow numbers
        If e.KeyChar >= "0" AndAlso e.KeyChar <= "9" Then
            Return
        End If

        ' Allow the backspace 
        If e.KeyChar = ControlChars.Back Then
            Return
        End If

        ' Allow the .
        If e.KeyChar = "." Then
            Return
        End If

        e.Handled = True

    End Sub

    Private Sub txtCount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCount.KeyPress

        'Allow numbers
        If e.KeyChar >= "0" AndAlso e.KeyChar <= "9" Then
            Return
        End If

        ' Allow the backspace 
        If e.KeyChar = ControlChars.Back Then
            Return
        End If

        e.Handled = True

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        'hide label errors
        HideLblErrors()


        'Defensive Programming 

        'check for Empty NDC 
        If Not mtbNDC.MaskFull Then
            lblErrorNDC.Visible = True
            MsgBox("NDC is required. Please fill in with 11 valid Digits", MsgBoxStyle.Critical, "Error!")
            mtbNDC.Select()
            Return
        End If

        'check for Empty DrugName
        If txtDrugName.Text = "" Then
            lblErrorDrugName.Visible = True
            MsgBox("Drug Name is required. Please fill in the box", MsgBoxStyle.Critical, "Error!")
            txtDrugName.Select()
            Return
        End If

        'check for Empty Dosage
        If txtDosage.Text = "" Then
            lblErrorDosage.Visible = True
            MsgBox("Dosage is required. Please fill in the box", MsgBoxStyle.Critical, "Error!")
            txtDosage.Select()
            Return
        End If

        'check for Empty Description
        If txtDesc.Text = "" Then
            lblErrorDesc.Visible = True
            MsgBox("Description is required. Please fill in the box", MsgBoxStyle.Critical, "Error!")
            txtDesc.Select()
            Return
        End If

        'check for Empty Unit Quantity
        If txtUnitQuantity.Text = "" Then
            lblErrorUnitQty.Visible = True
            MsgBox("Quantity Per Unit is required. Please fill in the box", MsgBoxStyle.Critical, "Error!")
            txtUnitQuantity.Select()
            Return
        End If

        'check for Empty unit Price
        If Not Decimal.TryParse(txtUnitPrice.Text, theDrug.unitPrice) Then
            lblErrorUnitP.Visible = True
            MsgBox("Please enter positive digits only", MsgBoxStyle.Critical, "Error!")
            txtUnitPrice.Select()
            Return
        End If

        'check for Empty Package Price
        If Not Decimal.TryParse(txtPackagePrice.Text, theDrug.packagePrice) Then
            lblErrorPkgP.Visible = True
            MsgBox("Please enter positive digits only", MsgBoxStyle.Critical, "Error!")
            txtPackagePrice.Select()
            Return
        End If

        'check for count
        If txtCount.Text = "" Then
            lblErrorCount.Visible = True
            MsgBox("Count is required. Please fill in the box", MsgBoxStyle.Critical, "Error!")
            txtCount.Select()
            Return
        End If


        'check for supplier
        If ddlSupplier.SelectedIndex = -1 Then
            lblErrorSupplier.Visible = True
            MsgBox("Supplier is required. Please pick one", MsgBoxStyle.Critical, "Error!")
            Return
        End If

        'check for category
        If lstCategory.SelectedIndex = -1 Then
            lblErrorCategory.Visible = True
            MsgBox("Category is required. Please pick one", MsgBoxStyle.Critical, "Error!")
            Return
        End If

        'check for RadioButtons
        If radAntibiotic.Checked = False AndAlso radAntidepressants.Checked = False _
            AndAlso radControlled.Checked = False AndAlso radOTC.Checked = False Then
            lblErrorDrugType.Visible = True
            MsgBox("Drug Type is required. Please choose one", MsgBoxStyle.Critical, "Error!")
            Return
        End If

        Dim drugTypeVal As String

        'for radio button value
        If radOTC.Checked Then
            drugTypeVal = radOTC.Text
        ElseIf radControlled.Checked Then
            drugTypeVal = radControlled.Text
        ElseIf radAntibiotic.Checked Then
            drugTypeVal = radAntibiotic.Text
        Else
            drugTypeVal = radAntidepressants.Text
        End If



        theDrug.NDC = mtbNDC.Text
        theDrug.drugName = txtDrugName.Text.Trim
        theDrug.dosage = txtDosage.Text.Trim
        theDrug.desc = txtDesc.Text.Trim
        theDrug.drugType = drugTypeVal
        theDrug.supplier = ddlSupplier.Text
        theDrug.category = lstCategory.Text
        theDrug.unitPrice = CDec(txtUnitPrice.Text)
        theDrug.packagePrice = CDec(txtPackagePrice.Text)
        theDrug.count = CInt(txtCount.Text)
        theDrug.dateReceived = CStr(dtpDateRecieved.Value.ToString("MM/dd/yyyy"))
        theDrug.unitQuantity = txtUnitQuantity.Text.Trim

        'for instock value
        If theDrug.count = 0 Then
            chkInStock.Checked = False
            theDrug.inStock = 0
        Else
            theDrug.inStock = 1
        End If

        Me.Close()

    End Sub
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        Dim msgBoxResult As Integer

            msgBoxResult = MessageBox.Show("Are you sure you want to Cancel? ", "Confirm?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

            If msgBoxResult = DialogResult.Yes Then
                Me.Close()
            End If



    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click

        mtbNDC.Text = ""
        txtDrugName.Text = ""
        txtDosage.Text = ""
        txtDesc.Text = ""
        txtUnitPrice.Text = ""
        txtPackagePrice.Text = ""
        txtCount.Text = ""
        radAntibiotic.Checked = False
        radOTC.Checked = False
        radAntidepressants.Checked = False
        radControlled.Checked = False
        lstCategory.SelectedIndex = -1
        ddlSupplier.SelectedIndex = -1
        chkInStock.Checked = False
        dtpDateRecieved.Value = Today
        txtUnitQuantity.Text = ""
        mtbNDC.Select()

        'hide label errors
        HideLblErrors()



    End Sub


    Private Sub HideLblErrors()

        lblErrorNDC.Visible = False
        lblErrorDrugName.Visible = False
        lblErrorDosage.Visible = False
        lblErrorDesc.Visible = False
        lblErrorSupplier.Visible = False
        lblErrorCategory.Visible = False
        lblErrorDrugType.Visible = False
        lblErrorUnitP.Visible = False
        lblErrorPkgP.Visible = False
        lblErrorInStock.Visible = False
        lblErrorCount.Visible = False
        lblErrorDate.Visible = False
        lblErrorUnitQty.Visible = False

    End Sub

End Class