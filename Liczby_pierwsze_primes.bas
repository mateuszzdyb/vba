Attribute VB_Name = "Module3"
Sub Liczby_pierwsze_primes()
Attribute Liczby_pierwsze_primes.VB_ProcData.VB_Invoke_Func = "O\n14"
' Liczby_pierwszePrimes Macro
' Keyboard Shortcut: Ctrl+Shift+O
                            'macro created by Mateusz Zdyb
                            '2018
                            'mateuszzdyb@yahoo.co.uk
                            'free to use for non-commercial purposes
Dim LiczbaPoczatkowa As Long
Dim LiczbaSprawdzana As Long
Dim Sprawdzacz As Long
Dim SprawdzaczPol As Long
Dim x As Long
Dim IleLiczbPierwszych As Long
LiczbaPoczatkowa = Range("b1")
LiczbaSprawdzana = LiczbaPoczatkowa
Cells(3, 2).Value = 2
Sprawdzacz = LiczbaSprawdzana / 3
Sprawdzacz = Int(Sprawdzacz)
SprawdzaczPol = Sprawdzacz + 1
Sprawdzacz = 2
x = 2
Line1:
Cells(3, 2).Value = Sprawdzacz
CzySieDzieli = LiczbaSprawdzana / Sprawdzacz
IntCzy = Int(CzySieDzieli)
Wynik = CzySieDzieli - IntCzy
    If Wynik = 0 Then
        LiczbaSprawdzana = LiczbaSprawdzana + 1
        Cells(1, 2).Value = LiczbaSprawdzana
    Else
        x = x + 1
        Sprawdzacz = Cells(x, 6).Value
        If Sprawdzacz >= SprawdzaczPol Then
            IleLiczbPierwszych = Range("d1")
            IleLiczbPierwszych = IleLiczbPierwszych + 1
            Cells(IleLiczbPierwszych, 6).Value = LiczbaSprawdzana
            Cells(IleLiczbPierwszych, 7).Value = Now()
            Cells(1, 4).Value = IleLiczbPierwszych
            LiczbaSprawdzana = LiczbaSprawdzana + 1
            Cells(1, 2).Value = LiczbaSprawdzana
        Else
            GoTo Line1
        End If
    End If
LiczbaSprawdzana = LiczbaSprawdzana + 2
Cells(2, 2).Value = LiczbaSprawdzana
Cells(5, 2).Value = x
End Sub

