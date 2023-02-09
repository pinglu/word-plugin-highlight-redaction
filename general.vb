Option Explicit


Public Function is_in_array(value As String, test_array) As Boolean
    If Not (IsArray(test_array)) Then Exit Function
    If InStr(1, "'" & Join(test_array, "'") & "'", "'" & value & "'") > 0 _
        Then is_in_array = True
End Function

