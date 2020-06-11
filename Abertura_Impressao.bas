Attribute VB_Name = "Módulo3"
Sub Abrir()

Sorteio.Show

End Sub


Sub Impressao()

If MsgBox("Você deseja imprimir?", vbYesNo, "Atenção!") = vbYes Then


Sheets("Imprimir").Select
Range("A1:N19").Select


copias = InputBox("Quantas cópias?")

Application.Dialogs(xlDialogPrinterSetup).Show
Selection.PrintOut copies:=Int(copias), collate:=True

Sheets("Menu").Activate

MsgBox "Impressão Efetuada com Sucesso!", vbOK, "ATENÇÃO!"


Else


End If


Exit Sub

End Sub


