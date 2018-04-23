Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Net
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Shapes
Imports System.IO
Imports System.Reflection
Imports DevExpress.XtraRichEdit
Imports System.Globalization
Imports System.Windows.Media.Imaging
Imports DevExpress.XtraRichEdit.Utils
Imports DevExpress.Office.Utils
Imports DevExpress.Office.NumberConverters

Namespace AutoCorrectEvent
    Partial Public Class MainPage
        Inherits UserControl

        Public Sub New()
            InitializeComponent()
        End Sub
        Private Sub richEditControl1_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            richEditControl1.ApplyTemplate()
            richEditControl1.CreateNewDocument()
            AddHandler richEditControl1.AutoCorrect, AddressOf richEditControl1_AutoCorrect
        End Sub

        Private Function CreateImageFromResx(ByVal name As String) As OfficeImage
            Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim stream As Stream = assembly.GetManifestResourceStream("Images." & name)
            Dim im As OfficeImage = OfficeImage.CreateImage(stream)
            Return im
        End Function
        #Region "#autocorrect"
        Private Sub richEditControl1_AutoCorrect(ByVal sender As Object, ByVal e As DevExpress.XtraRichEdit.AutoCorrectEventArgs)
            Dim info As AutoCorrectInfo = e.AutoCorrectInfo
            e.AutoCorrectInfo = Nothing

            If info.Text.Length <= 0 Then
                Return
            End If
            Do
                If Not info.DecrementStartPosition() Then
                    Return
                End If

                If IsSeparator(info.Text(0)) Then
                    Return
                End If

                If info.Text(0) = "$"c Then
                    info.ReplaceWith = CreateImageFromResx("dollar_pic.png")
                    e.AutoCorrectInfo = info
                    Return
                End If

                If info.Text(0) = "%"c Then
                    Dim replaceString As String = CalculateFunction(info.Text)
                    If Not String.IsNullOrEmpty(replaceString) Then
                        info.ReplaceWith = replaceString
                        e.AutoCorrectInfo = info
                    End If
                    Return
                End If
            Loop
        End Sub
        #End Region ' #autocorrect
        Private Function CalculateFunction(ByVal name As String) As String
            name = name.ToLower()

            If name.Length > 2 AndAlso name.Chars(0) = "%"c AndAlso name.EndsWith("%") Then
                Dim value As Integer = Nothing
                If Int32.TryParse(name.Substring(1, name.Length - 2), value) Then
                    Dim converter As OrdinalBasedNumberConverter = OrdinalBasedNumberConverter.CreateConverter(NumberingFormat.CardinalText, LanguageId.English)
                    Return converter.ConvertNumber(value)
                End If
            End If

            Select Case name
                Case "%date%"
                    Return Date.Now.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern)
                Case "%time%"
                    Return Date.Now.ToString(CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern)
                Case Else
                    Return String.Empty
            End Select
        End Function
        Private Function IsSeparator(ByVal ch As Char) As Boolean
            Return ch <> "%"c AndAlso (ch = ControlChars.Cr OrElse ch = ControlChars.Lf OrElse Char.IsPunctuation(ch) OrElse Char.IsSeparator(ch) OrElse Char.IsWhiteSpace(ch))
        End Function

    End Class
End Namespace
