Attribute VB_Name = "M�dulo9"
Sub InserirImagensAntes()

MsgBox ("INSERIR IMAGENS ANTES - Antes de prosseguir, cerifique-se de que a pasta selecionada possua apenas as imagens e nenhum outro arquivo. As imagens devem ter os nomes no padr�o -> (1), (2), (3), etc. tendo as extens�es JPG ou JPEG. Atualizado em 01/05/2024")

Dim fileNameAndPath As Variant
Dim image As Picture

Dim folderPicker As Variant
Dim picsFolder As Variant

Set folderPicker = Application.FileDialog(msoFileDialogFolderPicker)

'Selecionar pasta com as imagens
With folderPicker
    .Title = "Selecione a pasta com as imagens dos buracos"
    .AllowMultiSelect = True
    If .Show <> -1 Then Exit Sub
    picsFolder = .SelectedItems(1) & "\"
End With

'Obter a extens�o das imagens
Dim filesExt As String

filesExt = Application.InputBox("Digite a extens�o das imagens (JPEG ou JPG)")
filesExt = LCase(filesExt)
If Not (filesExt = "jpg" Or filesExt = "jpeg") Then
MsgBox ("Esta n�o � uma extens�o v�lida. Digite novamente")
filesExt = Application.InputBox("Digite a extens�o das imagens (Apenas JPEG ou JPG)")
    If Not (filesExt = "jpg" Or filesExt = "jpeg") Then 'Pergunta a extens�o pela segunda vez, e se for inv�lida novamente pede pro usu�rio conferir a pasta e executar novamente
    MsgBox ("Extens�o inv�lida. Verifique a sua pasta, todas as imagens devem ter a extens�o JPEG ou JPG. Quando estiver tudo certo, execute novamente.")
    Exit Sub
    End If
End If


'Declarar contador para as imagens de buracos
Dim imageCounter As Double
imageCounter = 0

'Instru��es e coleta da quantidade de imagens de buracos a serem inseridos
Dim buracosAmount As Double
MsgBox ("ATEN��O: Devido a limita��es t�cnicas o limite � de 100 imagens. Obrigado!")
buracosAmount = Application.InputBox("Informe a quantidade de imagens")
'Verificar se foi inserido valor superior a 100
If buracosAmount > 100 Then
    MsgBox ("Valor superior a 100, execute novamente.")
    Exit Sub
End If

'Declarando as dimens�es da imagem para 7x13cm
Dim centimetersWidth As Double
Dim centimetersHeight As Double
centimetersWidth = 13
centimetersHeight = 7

'Convertendo de cm para Pontos
Dim pointsW As Double
Dim pointsH As Double
pointsW = Application.CentimetersToPoints(centimetersWidth)
pointsH = Application.CentimetersToPoints(centimetersHeight)

'Verificando se a sele��o foi cancelada ou n�o obteve imagem nenhuma
If picsFolder = False Then Exit Sub

' Definindo vari�veis para controle de posicionamento
Dim leftPosition As Double
Dim topPosition As Double
leftPosition = ActiveSheet.Range("A1").Left + Application.CentimetersToPoints(0.4)
topPosition = ActiveSheet.Range("A17").Top + ActiveSheet.Cells(1, 1).Height

Dim colOffset As Integer
colOffset = 1 ' Come�amos na coluna A

Dim insertConfirmation As Variant

insertConfirmation = MsgBox("Voc� est� prestes a inser ir " & buracosAmount & " imagens. Deseja inserir?", vbYesNo)
If insertConfirmation = 7 Then Exit Sub

While imageCounter < buracosAmount
    imageCounter = imageCounter + 1
    'Verifica se existe a imagem do contador atual (imagem (1) ou (2) por exemplo) com o formato (jpg ou jpeg) que o usu�rio selecionou
    If Dir(picsFolder & "(" + CStr(imageCounter) + ")." + filesExt) = "" Then
        MsgBox "N�o foi poss�vel localizar a imagem (" + CStr(imageCounter) + ")." + filesExt + ". Verifique se esta imagem existe na sua pasta, se voc� selecionou a pasta correta, ou se est� com a extens�o correta."
        Exit Sub
    End If
    Set image = ActiveSheet.Pictures.Insert(picsFolder & "(" + CStr(imageCounter) + ")." + filesExt)
    With image
        ' Definindo como falso o bloqueio do aspecto da imagem
        .ShapeRange.LockAspectRatio = msoFalse
        ' Posicionando a imagem
        .Left = leftPosition
        .Top = topPosition
        ' Redimensionando a imagem
        .Width = pointsW
        .Height = pointsH
    End With

        leftPosition = ActiveSheet.Range("A1").Left + Application.CentimetersToPoints(0.4) ' Voltando para a coluna inicial + margem vertical da imagem de cima
        topPosition = topPosition + pointsH + (ActiveSheet.Cells(1, 1).Height * 4.775) ' Movendo para a pr�xima linha; Posicao a cima + altura da imagem + margem mais precisa possivel para alinhar a distanica entre as imagens da distancia entre os espacos destinados (4,775x a altura de uma celula)
        
Wend

End Sub
