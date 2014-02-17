Attribute VB_Name = "Module1"
Option Explicit

Public ServerActiveAtStart As Boolean

Sub Load_Source_Server()

Worksheets("Sheet1").Activate
rinterface.StartRServer

If rinterface.RLibraryIsAvailable("tseries") = False Then
    rinterface.RRun "install.packages(c(""tseries""))"
End If

If rinterface.RLibraryIsAvailable("zoo") = False Then
    rinterface.RRun "install.packages(c(""zoo""))"
End If
    
rinterface.RRun "library(tseries)"
rinterface.RRun "library(zoo)"
rinterface.RRun "source(file.choose())"

End Sub

Sub Plot_Weekly_TS()

'Declare Frequency Object
rinterface.PutDataframe "fre", Range("Sheet1!C5:C6"), False, False
rinterface.RunRCall "function(Data)ts(data=Data,frequency=fre[1,])", Range("Sheet1!Data")

'Find Optimal Model of Data
rinterface.GetRApply "function(Data)find.best.arima(Data)", Range("Sheet1!L2"), Range("Sheet1!Data")
rinterface.PutDataframe "z", Range("Sheet1!L1:L7"), False, False
rinterface.RunRCall "function(Data)sarima(Data,z[1,],z[2,],z[3,],z[4,],z[5,],z[6,],fre[1,])", Range("Sheet1!Data")
rinterface.InsertCurrentRPlot Range("Sheet1!BG4"), widthrescale:=0.7, heightrescale:=0.7, closergraph:=True

'Forecast n Days Ahead
rinterface.PutDataframe "days", Range("Sheet1!C7:C8"), False, False
rinterface.GetRApply "function(Data)sarima.for(Data,days,z[1,],z[2,],z[3,],z[4,],z[5,],z[6,],fre[1,])", Range("Sheet1!M2"), Range("Sheet1!Data")
rinterface.InsertCurrentRPlot Range("Sheet1!AY4"), widthrescale:=0.7, heightrescale:=0.7, closergraph:=True

Application.DisplayAlerts = False
Range("Sheet1!M2").TextToColumns Destination:=Range("N2"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=True, Space:=False, Other:=True, OtherChar:= _
    "(", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), _
    Array(6, 1)), TrailingMinusNumbers:=True

Range("Sheet1!M3").TextToColumns Destination:=Range("N3"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=True, Space:=False, Other:=True, OtherChar:= _
    "(", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), _
    Array(6, 1)), TrailingMinusNumbers:=True
Application.DisplayAlerts = True

End Sub

Sub Clear_Forecasted_Values()

On Error Resume Next

ActiveSheet.Shapes.Range(Array("RPlot002")).Select
    Selection.Delete
ActiveSheet.Shapes.Range(Array("RPlot001")).Select
    Selection.Delete
    
Range("Sheet1!O2:AP3").ClearContents
Range("Sheet1!L2:L7").ClearContents
Range("Sheet1!K5:K40").ClearContents

End Sub




