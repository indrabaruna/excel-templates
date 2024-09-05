VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Progress: Prepopulate Data"
   ClientHeight    =   768
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5388
   OleObjectBlob   =   "ProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


 
Function Progress(Progress_Percentage As Single, Optional Unload_After_Completion As Boolean = True)

    Me.Show False
    
    VBA.DoEvents
    
    Dim Total_width, Current_W As Single
    
    Total_width = 280
    
    Current_W = (Total_width / 100) * Progress_Percentage
    
    Me.lbl_Progress.Width = Current_W
    Me.lbl_Value.Caption = Format(Progress_Percentage, "0") & "%"
        
    If Unload_After_Completion = True And Progress_Percentage = 100 Then Unload Me
     
End Function
  

Private Sub lbl_Progress_Click()

End Sub

Private Sub lbl_Value_Click()

End Sub
