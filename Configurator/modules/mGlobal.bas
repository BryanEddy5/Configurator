Option Compare Database
Option Explicit


Public Property Get LibSys() As Form_fSystem
   Const fn As String = "fSystem"
   If Not CurrentProject.AllForms(fn).IsLoaded Then DoCmd.OpenForm fn, , , , , acHidden
   Set LibSys = Forms(fn)
End Property

Public Property Get dbs() As DAO.Database
   Set dbs = LibSys.Database
End Property