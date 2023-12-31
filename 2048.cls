VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub nowa_gra()          'Generuje liczby na pocz�tku gry
For i = 2 To 5
    For j = 4 To 7
    Cells(i, j).Value = Empty
    Next
Next
Cells(2, 10).Value = 0
Arkusz2.Cells(1, 1).Value = 0

Call Generowanie
Call Generowanie
Call zapisanieruchu
Call kolory
End Sub

Sub kolory()            'Zmiana koloru kom�rek
For i = 2 To 5
    For j = 4 To 7
    Select Case Cells(i, j).Value
        Case Is = Empty
            Cells(i, j).Interior.Color = RGB(205, 193, 181)
            Cells(i, j).Font.Color = RGB(117, 113, 113)
        Case Is = 2
            Cells(i, j).Interior.Color = RGB(236, 228, 219)
            Cells(i, j).Font.Color = RGB(117, 113, 113)
        Case Is = 4
            Cells(i, j).Interior.Color = RGB(236, 226, 208)
            Cells(i, j).Font.Color = RGB(117, 113, 113)
        Case Is = 8
            Cells(i, j).Interior.Color = RGB(233, 180, 133)
            Cells(i, j).Font.Color = RGB(255, 255, 255)
        Case Is = 16
            Cells(i, j).Interior.Color = RGB(233, 155, 115)
            Cells(i, j).Font.Color = RGB(255, 255, 255)
        Case Is = 32
            Cells(i, j).Interior.Color = RGB(231, 132, 110)
            Cells(i, j).Font.Color = RGB(255, 255, 255)
        Case Is = 64
            Cells(i, j).Interior.Color = RGB(229, 107, 81)
            Cells(i, j).Font.Color = RGB(255, 255, 255)
        Case Is = 128
            Cells(i, j).Interior.Color = RGB(234, 214, 152)
            Cells(i, j).Font.Color = RGB(255, 255, 255)
        Case Is = 256
            Cells(i, j).Interior.Color = RGB(233, 212, 142)
            Cells(i, j).Font.Color = RGB(255, 255, 255)
        Case Is = 512
            Cells(i, j).Interior.Color = RGB(240, 213, 113)
            Cells(i, j).Font.Color = RGB(255, 255, 255)
        Case Is = 1024
            Cells(i, j).Interior.Color = RGB(232, 207, 121)
            Cells(i, j).Font.Color = RGB(255, 255, 255)
        Case Else
            Cells(i, j).Interior.Color = RGB(229, 197, 66)
            Cells(i, j).Font.Color = RGB(255, 255, 255)
    End Select
    Next
Next
End Sub

Sub zapisanieruchu()    'Zapisanie ruchu na wypadek cofni�cia
For i = 2 To 5
    For j = 4 To 7
    Arkusz2.Cells(i, j) = Arkusz1.Cells(i, j)
    Next
Next
Arkusz2.Cells(2, 10).Value = Arkusz1.Cells(2, 10).Value
End Sub

Sub cofni�cie()         'Cofni�cie ruchu
For i = 2 To 5          'Wstawianie zapisanych kom�rek
    For j = 4 To 7
    Arkusz1.Cells(i, j).Value = Arkusz2.Cells(i, j).Value
    Next
Next
Arkusz1.Cells(2, 10).Value = Arkusz2.Cells(2, 10).Value
Call kolory
End Sub

Sub sprawdzenieczykoniec()       'Sprawdzanie, czy koniec gry
Dim ilo��pustych As Integer
Dim sprawdzenie As Integer
ilo��pustych = 0
sprawdzenie = 0

For i = 2 To 5          'Sprawdzanie, czy wszystkie kom�rki s� wype�nione
    For j = 4 To 7
    If Cells(i, j).Value = Empty Then
       ilo��pustych = ilo��pustych + 1
    End If
    Next
Next

For i = 2 To 4          'Sprawdzanie, czy wyst�puje jakie� po��czenie kom�rek
    For j = 4 To 7
    If Cells(i, j).Value = Cells(i + 1, j).Value Then
    sprawdzenie = 1
    End If
    Next
Next

For i = 2 To 5          'Sprawdzanie, czy wyst�puje jakie� po��czenie kom�rek
    For j = 4 To 6
    If Cells(i, j).Value = Cells(i, j + 1).Value Then
    sprawdzenie = 1
    End If
    Next
Next

For i = 2 To 4          'Sprawdzanie, czy wyst�puje 2048
If Arkusz2.Cells(1, 1).Value = 1 Then
Exit For
End If
    For j = 4 To 7
    If Cells(i, j).Value = 2048 Then
    wygrana = MsgBox("Wygra�e�! Czy chcesz gra� dalej?", 4, "Wygrana")
        If wygrana = vbYes Then
            Arkusz2.Cells(1, 1).Value = 1
        Else
            Call nowa_gra
        End If
    End If
    Next
Next
If ilo��pustych = 0 And sprawdzenie = 0 Then
    przegrana = MsgBox("Koniec gry. Tw�j wynik to " & Cells(2, 10).Value & " punkt�w.", vbOKOnly, "Przegrana")
    Call nowa_gra
End If
End Sub

Sub Generowanie()       'Generowanie kom�rek, tak aby 2 zdarza�a si� cz�ciej ni� 4
Dim ilo��pustych As Integer
Dim sprawdzenie As Integer
Dim generowana As Double
Dim generowana_pozycja As Double
ilo��pustych = 0
sprawdzenie = 0

For i = 2 To 5          'Sprawdzenie, czy s� puste kom�rki
    For j = 4 To 7
    If Cells(i, j).Value = Empty Then
       ilo��pustych = ilo��pustych + 1
    End If
    Next
Next

generowana = Rnd()
generowana_pozycja = Int((ilo��pustych - 1 + 1) * Rnd + 1)
ilo��pustych = 0

If generowana < 0.8 Then
    generowana = 2
Else
    generowana = 4
End If

For i = 2 To 5          'Wstawienie wygenerowanej liczby w kom�rk�
    If sprawdzenie = 1 Then
        Exit For
    End If
    For j = 4 To 7
        If Cells(i, j).Value = Empty Then
            If (ilo��pustych + 1) = generowana_pozycja Then
                Cells(i, j).Value = generowana
                sprawdzenie = 1
                Exit For
            Else
                ilo��pustych = ilo��pustych + 1
            End If
        End If
    Next
Next
End Sub

Sub lewo()              'Przesuni�cie kom�rek w lewo i po��czenie je�li s� takie same
Call zapisanieruchu     'Zapisanie ruchu, aby m�c cofn��

For i = 2 To 5
    For j = 4 To 7
        If Cells(i, j).Value = Empty And Cells(i, j + 1).Value <> Empty Then
            Cells(i, j).Value = Cells(i, j + 1).Value
            Cells(i, j + 1).Value = Empty
        End If
    Next
Next

For i = 2 To 5
    For j = 4 To 7
        If Cells(i, j).Value = Empty And Cells(i, j + 1).Value <> Empty Then
            Cells(i, j).Value = Cells(i, j + 1).Value
            Cells(i, j + 1).Value = Empty
        End If
    Next
Next

For i = 2 To 5
    For j = 4 To 7
        If Cells(i, j).Value = Empty And Cells(i, j + 1).Value <> Empty Then
            Cells(i, j).Value = Cells(i, j + 1).Value
            Cells(i, j + 1).Value = Empty
        End If
    Next
Next

For i = 2 To 5
If Cells(i, 4).Value = Cells(i, 5).Value And Cells(i, 4).Value <> Empty Then
    Cells(i, 4).Value = Cells(i, 4).Value + Cells(i, 5).Value
    Cells(2, 10).Value = Cells(2, 10).Value + Cells(i, 4).Value
    Cells(i, 5).Value = Cells(i, 6).Value
    Cells(i, 6).Value = Cells(i, 7).Value
    Cells(i, 7).Value = Empty
    If Cells(i, 5).Value = Cells(i, 6).Value And Cells(i, 5).Value <> Empty Then
        Cells(i, 5).Value = Cells(i, 5).Value + Cells(i, 6).Value
        Cells(2, 10).Value = Cells(2, 10).Value + Cells(i, 5).Value
        Cells(i, 6).Value = Empty
    End If
ElseIf Cells(i, 5).Value = Cells(i, 6).Value Then
    Cells(i, 5).Value = Cells(i, 5).Value + Cells(i, 6).Value
    Cells(2, 10).Value = Cells(2, 10).Value + Cells(i, 5).Value
    Cells(i, 6).Value = Cells(i, 7).Value
    Cells(i, 7).Value = Empty
ElseIf Cells(i, 6).Value = Cells(i, 7).Value Then
    Cells(i, 6).Value = Cells(i, 6).Value + Cells(i, 7).Value
    Cells(2, 10).Value = Cells(2, 10).Value + Cells(i, 6).Value
    Cells(i, 7).Value = Empty
End If
Next

For i = 2 To 5
    For j = 4 To 7
         If Cells(i, j).Value = 0 Then
            Cells(i, j).Value = Empty
        End If
    Next
Next

Call Generowanie            'Generowanie kom�rek po ruchu
Call kolory                 'Zmiana koloru kom�rek
Call sprawdzenieczykoniec   'Sprawdzenie, czy koniec gry
End Sub

Sub prawo()             'Przesuni�cie kom�rek w prawo i po��czenie je�li s� takie same
Call zapisanieruchu     'Zapisanie ruchu, aby m�c cofn��

For i = 2 To 5
    For j = 4 To 6
        If Cells(i, j).Value <> Empty And Cells(i, j + 1).Value = Empty Then
            Cells(i, j + 1).Value = Cells(i, j).Value
            Cells(i, j).Value = Empty
        End If
    Next
Next

For i = 2 To 5
    For j = 4 To 6
        If Cells(i, j).Value <> Empty And Cells(i, j + 1).Value = Empty Then
            Cells(i, j + 1).Value = Cells(i, j).Value
            Cells(i, j).Value = Empty
        End If
    Next
Next

For i = 2 To 5
    For j = 4 To 6
        If Cells(i, j).Value <> Empty And Cells(i, j + 1).Value = Empty Then
            Cells(i, j + 1).Value = Cells(i, j).Value
            Cells(i, j).Value = Empty
        End If
    Next
Next


For i = 2 To 5
If Cells(i, 7).Value = Cells(i, 6).Value And Cells(i, 7).Value <> Empty Then
    Cells(i, 7).Value = Cells(i, 7).Value + Cells(i, 6).Value
    Cells(2, 10).Value = Cells(2, 10).Value + Cells(i, 7).Value
    Cells(i, 6).Value = Cells(i, 5).Value
    Cells(i, 5).Value = Cells(i, 4).Value
    Cells(i, 4).Value = Empty
    If Cells(i, 6).Value = Cells(i, 5).Value And Cells(i, 6).Value <> Empty Then
        Cells(i, 6).Value = Cells(i, 5).Value + Cells(i, 5).Value
        Cells(2, 10).Value = Cells(2, 10).Value + Cells(i, 6).Value
        Cells(i, 5).Value = Empty
    End If
ElseIf Cells(i, 6).Value = Cells(i, 5).Value Then
    Cells(i, 6).Value = Cells(i, 6).Value + Cells(i, 5).Value
    Cells(2, 10).Value = Cells(2, 10).Value + Cells(i, 6).Value
    Cells(i, 5).Value = Cells(i, 4).Value
    Cells(i, 4).Value = Empty
ElseIf Cells(i, 5).Value = Cells(i, 4).Value Then
    Cells(i, 5).Value = Cells(i, 5).Value + Cells(i, 4).Value
    Cells(2, 10).Value = Cells(2, 10).Value + Cells(i, 5).Value
    Cells(i, 4).Value = Empty
End If
Next

For i = 2 To 5
    For j = 4 To 6
        If Cells(i, j).Value = 0 Then
            Cells(i, j).Value = Empty
        End If
    Next
Next

Call Generowanie            'Generowanie kom�rek po ruchu
Call kolory                 'Zmiana koloru kom�rek
Call sprawdzenieczykoniec   'Sprawdzenie, czy koniec gry
End Sub

Sub d�()               'Przesuni�cie kom�rek w d� i po��czenie je�li s� takie same
Call zapisanieruchu     'Zapisanie ruchu, aby m�c cofn��

For j = 4 To 7
    For i = 1 To 4
        If Cells(i, j).Value <> Empty And Cells(i + 1, j).Value = Empty Then
            Cells(i + 1, j).Value = Cells(i, j).Value
            Cells(i, j).Value = Empty
        End If
    Next
Next

For j = 4 To 7
    For i = 1 To 4
        If Cells(i, j).Value <> Empty And Cells(i + 1, j).Value = Empty Then
            Cells(i + 1, j).Value = Cells(i, j).Value
            Cells(i, j).Value = Empty
        End If
    Next
Next

For j = 4 To 7
    For i = 1 To 4
        If Cells(i, j).Value <> Empty And Cells(i + 1, j).Value = Empty Then
            Cells(i + 1, j).Value = Cells(i, j).Value
            Cells(i, j).Value = Empty
        End If
    Next
Next

For j = 4 To 7
    If Cells(5, j).Value = Cells(4, j).Value And Cells(5, j).Value <> Empty Then
        Cells(5, j).Value = Cells(5, j).Value + Cells(4, j).Value
        Cells(2, 10).Value = Cells(2, 10).Value + Cells(5, j).Value
        Cells(4, j).Value = Cells(3, j).Value
        Cells(3, j).Value = Cells(2, j).Value
        Cells(2, j).Value = Empty
        If Cells(4, j).Value = Cells(3, j).Value And Cells(4, j).Value <> Empty Then
            Cells(4, j).Value = Cells(4, j).Value + Cells(3, j).Value
            Cells(2, 10).Value = Cells(2, 10).Value + Cells(4, j).Value
            Cells(3, j).Value = Empty
        End If
    ElseIf Cells(4, j).Value = Cells(3, j).Value Then
        Cells(4, j).Value = Cells(4, j).Value + Cells(3, j).Value
        Cells(2, 10).Value = Cells(2, 10).Value + Cells(4, j).Value
        Cells(3, j).Value = Cells(2, j).Value
        Cells(2, j).Value = Empty
    ElseIf Cells(3, j).Value = Cells(2, j).Value Then
        Cells(3, j).Value = Cells(3, j).Value + Cells(2, j).Value
        Cells(2, 10).Value = Cells(2, 10).Value + Cells(3, j).Value
        Cells(2, j).Value = Empty
    End If
Next

For j = 4 To 7
    For i = 1 To 4
        If Cells(i, j).Value = 0 Then
            Cells(i, j).Value = Empty
        End If
    Next
Next

Call Generowanie            'Generowanie kom�rek po ruchu
Call kolory                 'Zmiana koloru kom�rek
Call sprawdzenieczykoniec   'Sprawdzenie, czy koniec gry
End Sub

Sub g�ra()              'Przesuni�cie kom�rek w g�r� i po��czenie je�li s� takie same
Call zapisanieruchu     'Zapisanie ruchu, aby m�c cofn��
Dim sprawdzenie As Integer
For j = 4 To 7
    For i = 2 To 4
        If Cells(i, j).Value = Empty And Cells(i + 1, j).Value <> Empty Then
            Cells(i, j).Value = Cells(i + 1, j).Value
            Cells(i + 1, j).Value = Empty
        End If
    Next
Next

For j = 4 To 7
    For i = 2 To 4
        If Cells(i, j).Value = Empty And Cells(i + 1, j).Value <> Empty Then
            Cells(i, j).Value = Cells(i + 1, j).Value
            Cells(i + 1, j).Value = Empty
        End If
    Next
Next

For j = 4 To 7
    For i = 2 To 4
        If Cells(i, j).Value = Empty And Cells(i + 1, j).Value <> Empty Then
            Cells(i, j).Value = Cells(i + 1, j).Value
            Cells(i + 1, j).Value = Empty
        End If
    Next
Next

For j = 4 To 7
    If Cells(2, j).Value = Cells(3, j).Value And Cells(2, j).Value <> Empty Then
        Cells(2, j).Value = Cells(2, j).Value + Cells(3, j).Value
        Cells(2, 10).Value = Cells(2, 10).Value + Cells(2, j).Value
        Cells(3, j).Value = Cells(4, j).Value
        Cells(4, j).Value = Cells(5, j).Value
        Cells(5, j).Value = Empty
        If Cells(3, j).Value = Cells(4, j).Value And Cells(3, j).Value <> Empty Then
            Cells(3, j).Value = Cells(3, j).Value + Cells(4, j).Value
            Cells(2, 10).Value = Cells(2, 10).Value + Cells(3, j).Value
            Cells(4, j).Value = Empty
        End If
    ElseIf Cells(3, j).Value = Cells(4, j).Value Then
        Cells(3, j).Value = Cells(3, j).Value + Cells(4, j).Value
        Cells(2, 10).Value = Cells(2, 10).Value + Cells(3, j).Value
        Cells(4, j).Value = Cells(5, j).Value
        Cells(5, j).Value = Empty
    ElseIf Cells(4, j).Value = Cells(5, j).Value Then
        Cells(4, j).Value = Cells(4, j).Value + Cells(5, j).Value
        Cells(2, 10).Value = Cells(2, 10).Value + Cells(4, j).Value
        Cells(5, j).Value = Empty
    End If
Next

For j = 4 To 7
    For i = 2 To 4
        If Cells(i, j).Value = 0 Then
            Cells(i, j).Value = Empty
        End If
    Next
Next

Call Generowanie            'Generowanie kom�rek po ruchu
Call kolory                 'Zmiana koloru kom�rek
Call sprawdzenieczykoniec   'Sprawdzenie, czy koniec gry
End Sub
