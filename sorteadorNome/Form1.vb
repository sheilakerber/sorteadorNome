Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnSortear.Click
        ' valida preenchimento dos campos
        If (txtNome1.Text = Nothing Or txtNome2.Text = Nothing Or txtNome3.Text = Nothing Or txtNome4.Text = Nothing Or txtNome5.Text = Nothing) Then
            MsgBox("Há campos vazios. Por favor, verifique.", MsgBoxStyle.Critical, "Atenção!")
            Exit Sub
        End If

        ' se tudo estiver preenchido, cria o array nomes e guarda os inputs do usuario
        Dim nomes As Object = {txtNome1.Text, txtNome2.Text, txtNome3.Text, txtNome4.Text, txtNome5.Text}

        ' sorteia um index do array
        Dim nomeSorteado As Integer = CInt(Int(5 * Rnd() + 1))

        ' apresenta na tela o nome correspondente ao index sorteado
        MsgBox("Nome sorteado: " & nomes(nomeSorteado))

        'limpa os campos
        txtNome1.Text = Nothing
        txtNome2.Text = Nothing
        txtNome3.Text = Nothing
        txtNome4.Text = Nothing
        txtNome5.Text = Nothing
    End Sub
End Class
