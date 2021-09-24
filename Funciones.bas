Attribute VB_Name = "Funciones"
Public Function depurar()

Dim Cn As New ADODB.Connection
Dim rsd As New ADODB.Recordset
Cn.ConnectionString = Cadena
Cn.Open

'Dim fechaact As Date
fechaact = Format(Date, "yyyy-mm-dd")
'fechaact = fechaact - 8
fechaact = Format(fechaact, "yyyy-mm-dd")
'fechadep = Format(fechaact, "yyyy-mm-dd")

borra = "DELETE FROM depurados"
Cn.Execute borra

'copia = "INSERT INTO depurados SELECT * from nuevos WHERE NueEst = " & 1 & " AND NueFeM <= " & "'" & fechadep & "'"
copia = "INSERT INTO depurados SELECT * from nuevos WHERE NueEst = " & 1 & " AND NueFeD <= " & "'" & fechaact & "'"
Cn.Execute copia

rsd.CursorType = adOpenKeyset
rsd.LockType = adLockOptimistic
rsd.ActiveConnection = Cn
rsd.Source = "Select * from depurados"
rsd.Open

If Not rsd.EOF Then
    siexiste = 1
Else
    siexiste = 0
End If

End Function



