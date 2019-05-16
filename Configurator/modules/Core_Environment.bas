Option Compare Database
Option Explicit
Private Const environmentArg = "Dev = "

Public Enum EnviromentType
    Prod = 0
    DEV = 1
End Enum


Public Function SetEnvironment(Environment As EnviromentType)

    Dim EvironmentArg As String
    EvironmentArg = environmentArg & Environment

    Call Application.SetOption("Conditional Compilation Arguments", EvironmentArg)


End Function

Private Property Get GetCompilationArguments() As String

    GetCompilationArguments = Application.GetOption("Conditional Compilation Arguments")

End Property

Public Property Get CurrentEnvironmentState() As EnviromentType

    Dim hold() As String
    Dim env
    
    hold = Split(GetCompilationArguments, environmentArg)
    
    env = CInt(hold(1))
    
    CurrentEnvironmentState = env

End Property