VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sorteio 
   Caption         =   "Rematrícula Anglo Morumbi 2019"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12312
   OleObjectBlob   =   "Formulario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sorteio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Atualizar_Click()

If MsgBox("Você deseja salvar as informações?", vbYesNo, "Atenção!") = vbYes Then

MsgBox "Os Dados Foram Atualizados com Sucesso!", vbCritical, "ERRO"


Dim linha As Integer
Dim RA As String


RA = RA_ALUN_CX
linha = 2


Do Until Sheets("Dados").Cells(linha, 1) = ""

If RA = Sheets("Dados").Cells(linha, 1) Then

Sheets("Dados").Cells(linha, 2).Value = NOME_ALUN_CX
Sheets("Dados").Cells(linha, 3).Value = CPF_RESP_CX
Sheets("Dados").Cells(linha, 4).Value = NOME_RESP_CX
Sheets("Dados").Cells(linha, 5).Value = RG_RESP_CX
Sheets("Dados").Cells(linha, 6).Value = END_CX
Sheets("Dados").Cells(linha, 7).Value = CIDADE_CX
Sheets("Dados").Cells(linha, 8).Value = UF_CX
Sheets("Dados").Cells(linha, 9).Value = CEP_CX
Sheets("Dados").Cells(linha, 10).Value = EMAIL_CX
Sheets("Dados").Cells(linha, 11).Value = TEL_CX



NOME_ALUN_CX.Enabled = False
CPF_RESP_CX.Enabled = False
NOME_RESP_CX.Enabled = False
RG_RESP_CX.Enabled = False
END_CX.Enabled = False
CIDADE_CX.Enabled = False
UF_CX.Enabled = False
CEP_CX.Enabled = False
EMAIL_CX.Enabled = False
TEL_CX.Enabled = False

Exit Sub

End If

linha = linha + 1

Loop



End If
End Sub



Private Sub Editar_Click()

RA_ALUN_CX.Enabled = False
NOME_ALUN_CX.Enabled = True
CPF_RESP_CX.Enabled = True
NOME_RESP_CX.Enabled = True
RG_RESP_CX.Enabled = True
END_CX.Enabled = True
CIDADE_CX.Enabled = True
UF_CX.Enabled = True
CEP_CX.Enabled = True
EMAIL_CX.Enabled = True
TEL_CX.Enabled = True

End Sub
Private Sub Imprimir_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Imprimir_Click()

Sheets("Imprimir").Range("B3").Value = Sorteio.RA_ALUN_CX
Sheets("Imprimir").Range("D3").Value = Sorteio.NOME_ALUN_CX
Sheets("Imprimir").Range("B6").Value = Sorteio.CPF_RESP_CX
Sheets("Imprimir").Range("G6").Value = Sorteio.RG_RESP_CX
Sheets("Imprimir").Range("B9").Value = Sorteio.END_CX
Sheets("Imprimir").Range("B12").Value = Sorteio.CIDADE_CX
Sheets("Imprimir").Range("H12").Value = Sorteio.CEP_CX
Sheets("Imprimir").Range("L12").Value = Sorteio.UF_CX
Sheets("Imprimir").Range("B15").Value = Sorteio.TEL_CX
Sheets("Imprimir").Range("F15").Value = Sorteio.EMAIL_CX
Sheets("Imprimir").Range("B18").Value = Sorteio.NOME_RESP_CX


Call Impressao


End Sub

Private Sub Lupa_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Lupa_Click()
Dim linha As Integer
Dim RA As Long


On Error GoTo Erro

linha = 2
RA = RA_ALUN_CX

Do Until Sheets("Dados").Cells(linha, 1) = ""

If RA = Sheets("Dados").Cells(linha, 1) Then

NOME_ALUN_CX = Sheets("Dados").Cells(linha, 2).Value
CPF_RESP_CX = Sheets("Dados").Cells(linha, 3).Value
NOME_RESP_CX = Sheets("Dados").Cells(linha, 4).Value
RG_RESP_CX = Sheets("Dados").Cells(linha, 5).Value
END_CX = Sheets("Dados").Cells(linha, 6).Value
CIDADE_CX = Sheets("Dados").Cells(linha, 7).Value
UF_CX = Sheets("Dados").Cells(linha, 8).Value
CEP_CX = Sheets("Dados").Cells(linha, 9).Value
EMAIL_CX = Sheets("Dados").Cells(linha, 10).Value
TEL_CX = Sheets("Dados").Cells(linha, 11).Value

NOME_ALUN_CX.Enabled = False
CPF_RESP_CX.Enabled = False
NOME_RESP_CX.Enabled = False
RG_RESP_CX.Enabled = False
END_CX.Enabled = False
CIDADE_CX.Enabled = False
UF_CX.Enabled = False
CEP_CX.Enabled = False
EMAIL_CX.Enabled = False
TEL_CX.Enabled = False

Exit Sub

End If

linha = linha + 1

Loop

Erro:
MsgBox "Aluno não encontrado, favor verificar o RA digitado", vbExclamation, "AVISO"

End Sub
Private Sub Limpar_Click()


Dim objeto As Control

For Each objeto In Me.Controls 'faz o looping percorrendo todos os objetos do Userform1
If TypeName(objeto) = "TextBox" Or TypeName(objeto) = "ComboBox" Then  ' se o tipo do objeto encontrado tiver o nome TEXTBOX
            objeto.Text = "" 'limpa o campo
            End If
Next objeto

    For Each bt In Sorteio.Controls
        If Left(bt.Name, 3) = "Opt" Then
            bt.Value = False
        End If
    Next
    
NOME_ALUN_CX.Enabled = True
CPF_RESP_CX.Enabled = True
NOME_RESP_CX.Enabled = True
RG_RESP_CX.Enabled = True
END_CX.Enabled = True
CIDADE_CX.Enabled = True
UF_CX.Enabled = True
CEP_CX.Enabled = True
EMAIL_CX.Enabled = True
TEL_CX.Enabled = True

End Sub

Private Sub Matricula_Click()
Call Limpar_Click

NOME_ALUN_CX.Enabled = True
CPF_RESP_CX.Enabled = True
NOME_RESP_CX.Enabled = True
RG_RESP_CX.Enabled = True
END_CX.Enabled = True

CIDADE_CX.Enabled = True
UF_CX.Enabled = True
CEP_CX.Enabled = True
EMAIL_CX.Enabled = True
TEL_CX.Enabled = True


End Sub

Private Sub Salvar_Click()
If MsgBox("Você deseja salvar as informações?", vbYesNo, "Atenção!") = vbYes Then


Dim linha As Integer
Dim RA As String


RA = RA_ALUN_CX
linha = 2


Do Until Sheets("Dados").Cells(linha, 1) = ""

If RA = Sheets("Dados").Cells(linha, 1) Then

Sheets("Dados").Cells(linha, 2).Value = NOME_ALUN_CX
Sheets("Dados").Cells(linha, 3).Value = CPF_RESP_CX
Sheets("Dados").Cells(linha, 4).Value = NOME_RESP_CX
Sheets("Dados").Cells(linha, 5).Value = RG_RESP_CX
Sheets("Dados").Cells(linha, 6).Value = END_CX
Sheets("Dados").Cells(linha, 7).Value = CIDADE_CX
Sheets("Dados").Cells(linha, 8).Value = UF_CX
Sheets("Dados").Cells(linha, 9).Value = CEP_CX
Sheets("Dados").Cells(linha, 10).Value = EMAIL_CX
Sheets("Dados").Cells(linha, 11).Value = TEL_CX



NOME_ALUN_CX.Enabled = False
CPF_RESP_CX.Enabled = False
NOME_RESP_CX.Enabled = False
RG_RESP_CX.Enabled = False
END_CX.Enabled = False
CIDADE_CX.Enabled = False
UF_CX.Enabled = False
CEP_CX.Enabled = False
EMAIL_CX.Enabled = False
TEL_CX.Enabled = False

Exit Sub

End If

linha = linha + 1

Loop

End If

End Sub


