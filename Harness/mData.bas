Attribute VB_Name = "mData"
Option Explicit

'/* change the data structure to suit your application
'/* if you want it dynamic, unrem the subitem structure
'/* and rem the subitem1/subitem2 entries

Public Type HLISubItm
    lIcon       As Long
    Text()      As String
End Type

Public Type HLIStc
    Item()      As String
    lIcon()     As Long
    SubItem1()  As String
    SubItem2()  As String
    'SubItem()   As HLISubItm
End Type

Public Type HLIRtn
    RItem       As String
    RSubItem()  As String
End Type


