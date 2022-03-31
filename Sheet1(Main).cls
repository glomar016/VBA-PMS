VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdAddNewEmployee_Click()
    newEmployeeForm.Show
End Sub

Private Sub cmdGeneratePayslip_Click()
    generatePayslip.Show
End Sub

Private Sub cmdUpdateEmployee_Click()
    updateEmployeeForm.Show
End Sub

Private Sub CommandButton1_Click()
    viewDashboard.Show
End Sub

