Attribute VB_Name = "Module1"
Public acctnum As String
Public recordLength As Long
Public filename As String

Public Type ClientRecord
      accountNumber As Integer
      lastName As String * 15
      firstName As String * 15
      balance As Currency
End Type

Public mUdtClient As ClientRecord   ' user-defined type

'recordLength = LenB(mUdtClient)

