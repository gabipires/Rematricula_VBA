Attribute VB_Name = "M�dulo3"
Sub Abrir()

Sorteio.Show

End Sub


Sub Impressao()

If MsgBox("Voc� deseja imprimir?", vbYesNo, "Aten��o!") = vbYes Then


Sheets("Imprimir").Select
Range("A1:N19").Select


copias = InputBox("Quantas c�pias?")

Application.Dialogs(xlDialogPrinterSetup).Show
Selection.PrintOut copies:=Int(copias), collate:=True

Sheets("Menu").Activate

MsgBox "Impress�o Efetuada com Sucesso!", vbOK, "ATEN��O!"


Else


End If


Exit Sub

End Sub


