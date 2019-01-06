Attribute VB_Name = "mGlobal"
Option Compare Database
Option Explicit


Public Property Get LibSys() As Form_fSystem
   Const FN As String = "fSystem"
   If Not CurrentProject.AllForms(FN).IsLoaded Then DoCmd.OpenForm FN, , , , , acHidden
   Set LibSys = Forms(FN)
End Property

Public Property Get dbs() As DAO.Database
   Set dbs = LibSys.Database
End Property
