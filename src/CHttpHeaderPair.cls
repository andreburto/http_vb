VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHttpHeaderPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==============================================================
' CHttpHeaderPair
'
' Name/Value holder returned by CHttpHeaders.Item when enumerating
' headers. Consumers read it; it is not meant to be created
' directly (CHttpHeaders owns and populates every instance).
'==============================================================

Public Name As String
Public Value As String
