Partial Class ErrorPage
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Common As New Common
        Dim ErrorArgs As ErrorArgs
        Dim ErrorMessage As String

        ErrorArgs = Session("ErrorArgs")
        ErrorMessage = Replace(Request.QueryString("ErrorMessage"), "[sharp]", "#")
        ErrorMessage = Replace(ErrorMessage, "[sharp]", "#")
        ErrorMessage = Replace(ErrorMessage, "~", "<br>")

        litError.Text = ErrorArgs.HeaderMessage & "<br><br>" & ErrorMessage

        If InStr(ErrorMessage, "timed out") > 0 Then
            plClock.Visible = True
        Else
            plClock.Visible = False
        End If
        PageCaption.Text = Common.GetPageCaption
    End Sub

End Class