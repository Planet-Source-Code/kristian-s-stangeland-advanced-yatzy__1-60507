VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLocalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Copyright (C) 2004 Kristian. S.Stangeland

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Enum LOCALE_CONSTANTS
    LOCALE_SCURRENCY = &H14         'local symbol
    LOCALE_SINTLSYMBOL = &H15       'intl symbol
    LOCALE_SMONDECIMALSEP = &H16    'decimal separator
    LOCALE_SMONTHOUSANDSEP = &H17   'thousand separator
    LOCALE_SMONGROUPING = &H18      'grouping
    LOCALE_ICURRDIGITS = &H19       '# local digits
    LOCALE_IINTLCURRDIGITS = &H1A   '# intl digits
    LOCALE_ICURRENCY = &H1B         'pos currency mode
    LOCALE_INEGCURR = &H1C          'neg currency mode
    LOCALE_SSHORTDATE = &H1F        'short date mode
    LOCALE_IPOSSIGNPOSN = &H52      'pos sign position
    LOCALE_INEGSIGNPOSN = &H53      'neg sign position
    LOCALE_IPOSSYMPRECEDES = &H54   'mon sym precedes pos amt
    LOCALE_IPOSSEPBYSPACE = &H55    'mon sym sep by space from pos amt
    LOCALE_INEGSYMPRECEDES = &H56   'mon sym precedes neg amt
    LOCALE_INEGSEPBYSPACE = &H57    'mon sym sep by space from neg amt
    LOCALE_SENGCURRNAME = &H1007    'english name of currency
    LOCALE_SNATIVECURRNAME = &H1008 'native name of currency
    LOCALE_USER_DEFAULT = &H400     'default user-setting
End Enum

Public Function Clone() As clsLocalInfo

    Set Clone = New clsLocalInfo

End Function

Private Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As LOCALE_CONSTANTS) As String

    Dim sReturn As String, ret As Long
    
    ret = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
    If ret Then
     
        sReturn = Space$(ret)
        ret = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
          
        If ret Then
            GetUserLocaleInfo = Left$(sReturn, ret - 1)
        End If
    
    End If
    
End Function

Public Property Get LocalSymbol() As String

    LocalSymbol = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY)
    
End Property

Public Property Get IntlSymbol() As String

    IntlSymbol = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SINTLSYMBOL)

End Property

Public Property Get DecimalSeparator() As String
    
    DecimalSeparator = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONDECIMALSEP)

End Property

Public Property Get ThousandSeparator() As String
    
    ThousandSeparator = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHOUSANDSEP)

End Property

Public Property Get Grouping() As String
    
    Grouping = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONGROUPING)

End Property

Public Property Get LocalDigits() As String
    
    LocalDigits = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ICURRDIGITS)

End Property

Public Property Get IntlDigits() As String
    
    IntlDigits = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IINTLCURRDIGITS)

End Property

Public Property Get PosCurrencyMode() As String
    
    PosCurrencyMode = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ICURRENCY)

End Property

Public Property Get NegCurrencyMode() As String
    
    NegCurrencyMode = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_INEGCURR)

End Property

Public Property Get PosSignPosition() As String
    
    PosSignPosition = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IPOSSIGNPOSN)

End Property

Public Property Get NegSignPosition() As String
    
    NegSignPosition = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_INEGSIGNPOSN)

End Property

Public Property Get MonSymPrecedesPosAmt() As String
    
    MonSymPrecedesPosAmt = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IPOSSYMPRECEDES)

End Property

Public Property Get MonSymSepBySpaceFromPosAmt() As String
    
    MonSymSepBySpaceFromPosAmt = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IPOSSEPBYSPACE)

End Property

Public Property Get MonSymPrecedesNegAmt() As String
    
    MonSymPrecedesNegAmt = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_INEGSYMPRECEDES)

End Property

Public Property Get MonSymSepBySpaceFromNegAmt() As String
    
    MonSymSepBySpaceFromNegAmt = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_INEGSEPBYSPACE)

End Property

Public Property Get EnglishCurrencyName() As String
    
    EnglishCurrencyName = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGCURRNAME)

End Property

Public Property Get NativeCurrencyName() As String
    
    NativeCurrencyName = GetUserLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SNATIVECURRNAME)

End Property
