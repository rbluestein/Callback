
Partial Class DetectCaller
    Inherits System.Web.UI.Page

    Private Sub DetectCaller_Load(sender As Object, e As EventArgs) Handles Me.Load
        Response.Redirect("Index.aspx?Source=DetectCaller")
    End Sub
End Class
