Attribute VB_Name = "Module1"
Sub multitestprimes()
Attribute multitestprimes.VB_Description = "multi test for test sub"
Attribute multitestprimes.VB_ProcData.VB_Invoke_Func = "M\n14"
'multitest macro
'Keyboard Shortcut: ctrl+shift+M
                            'macro created by Mateusz Zdyb
                            '2018
                            'mateuszzdyb@yahoo.co.uk
                            'free to use for non-commercial purposes
For i = 1 To Range("b4")
Call Liczby_pierwsze_primes
Next i
End Sub
