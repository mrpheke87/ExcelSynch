Private Sub cmdSubmit_Click()
Call pushDataToGoogleSheet
End Sub

Private Sub UserForm_Activate()
cbKelas.AddItem "X IPA 1"
cbKelas.AddItem "X IPA 2"
cbKelas.AddItem "X IPA 3"
cbKelas.AddItem "X IPS 1"
cbKelas.AddItem "X IPS 2"

Laki_laki.Value = True

End Sub




