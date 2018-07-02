﻿Imports System.Reflection

Public Class About
    Private Sub About_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblname.Text = AssemblyProduct
        lblversion.Text = AssemblyVersion
        lblcopyright.Text = AssemblyCopyright
        lblcompany.Text = AssemblyCompany
        lbldescription.Text = AssemblyDescription
    End Sub

#Region "Assembly Attribute Accessors"

    Public ReadOnly Property AssemblyTitle() As String
        Get
            ' Get all Title attributes on this assembly
            Dim attributes As Object() = System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttributes(GetType(AssemblyTitleAttribute), False)
            ' If there is at least one Title attribute
            If attributes.Length > 0 Then
                ' Select the first one
                Dim titleAttribute As AssemblyTitleAttribute = CType(attributes(0), AssemblyTitleAttribute)
                ' If it is not an empty string, return it
                If titleAttribute.Title <> "" Then
                    Return titleAttribute.Title
                End If
            End If
            ' If there was no Title attribute, or if the Title attribute was the empty string, return the .exe name
            Return System.IO.Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
        End Get
    End Property

    Public ReadOnly Property AssemblyVersion() As String
        Get
            Return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()
        End Get
    End Property

    Public ReadOnly Property AssemblyDescription() As String
        Get
            ' Get all Description attributes on this assembly
            Dim attributes As Object() = System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttributes(GetType(AssemblyDescriptionAttribute), False)
            ' If there aren't any Description attributes, return an empty string
            If attributes.Length = 0 Then
                Return ""
            End If
            ' If there is a Description attribute, return its value
            Return (CType(attributes(0), AssemblyDescriptionAttribute)).Description
        End Get
    End Property

    Public ReadOnly Property AssemblyProduct() As String
        Get
            ' Get all Product attributes on this assembly
            Dim attributes As Object() = System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttributes(GetType(AssemblyProductAttribute), False)
            ' If there aren't any Product attributes, return an empty string
            If attributes.Length = 0 Then
                Return ""
            End If
            ' If there is a Product attribute, return its value
            Return (CType(attributes(0), AssemblyProductAttribute)).Product
        End Get
    End Property

    Public ReadOnly Property AssemblyCopyright() As String
        Get
            ' Get all Copyright attributes on this assembly
            Dim attributes As Object() = System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttributes(GetType(AssemblyCopyrightAttribute), False)
            ' If there aren't any Copyright attributes, return an empty string
            If attributes.Length = 0 Then
                Return ""
            End If
            ' If there is a Copyright attribute, return its value
            Return (CType(attributes(0), AssemblyCopyrightAttribute)).Copyright
        End Get
    End Property

    Public ReadOnly Property AssemblyCompany() As String
        Get
            ' Get all Company attributes on this assembly
            Dim attributes As Object() = System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttributes(GetType(AssemblyCompanyAttribute), False)
            ' If there aren't any Company attributes, return an empty string
            If attributes.Length = 0 Then
                Return ""
            End If
            ' If there is a Company attribute, return its value
            Return (CType(attributes(0), AssemblyCompanyAttribute)).Company
        End Get
    End Property

    Private Sub BunifuImageButton1_Click_1(sender As Object, e As EventArgs) Handles BunifuImageButton1.Click
        Me.Close()
    End Sub

    Private Sub BunifuImageButton2_Click(sender As Object, e As EventArgs) Handles BunifuImageButton2.Click
        Process.Start("https://github.com/gaiththewolf")
    End Sub
#End Region

End Class