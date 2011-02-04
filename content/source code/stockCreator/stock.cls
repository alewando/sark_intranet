VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function getStock() As String
    Dim objSource As getSource.cWX

    Set objSource = New getSource.cWX
    objSource.OpenURL ("http://quicken.excite.com/investments/quotes/?defview=QUICK&symbol=TEMFX%2CLSCRX%2COAKMX%2CLSBDX%2CSCDGX%2CTWCUX%2CJANSX")
    getStock = objSource.htmlWX

End Function


