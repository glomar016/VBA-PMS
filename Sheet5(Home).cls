VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub cmdView_Click()
    viewDashboard.Show
End Sub

Private Sub cmdNew_Click()
    newEmployeeForm.Show
End Sub


Private Sub cmdUpdate_Click()
    updateEmployeeForm.Show
End Sub

Private Sub cmdGenerate_Click()
    generatePayslip.Show
End Sub

Private Sub CommandButton1_Click()
    generateFinalPayslip.Show
End Sub
