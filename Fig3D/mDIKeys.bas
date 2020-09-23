Attribute VB_Name = "mDIKeys"
Option Explicit

'cRVTDX.mDIKeys - a component of the rvtDX.dll
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'This Module Remaps VBKey constants to DIK Key constants so
'that a program doesnt need to know the DirectInput constants
'The DI constants are MUCH more extensive and are probably used best directly for smarter apps

Public Enum vbExtraKeyCodes
  vbKeyALT = vbKeyMenu
  vbKeyGrave = &HC0&
  vbKeyMinus = &HBD&
  vbKeyEquals = &HBB&
  vbKeyLBracket = &HDB&
  vbKeyRBracket = &HDD&
  vbKeyBackSlash = &HDC&
  vbKeySemiColon = &HBA&
  vbKeyApostrophe = &HDE&
  vbKeyComma = &HBC&
  vbKeyPeriod = &HBE&
  vbKeySlash = &HBF&
  vbKeyLWin = &H5B&
  vbRWin = &H5C&
  vbScrollLock = &H91&
End Enum

Public vbToDIK(0 To 255) As Byte

'This is called ONCE at cDX8 Init

Public Sub MakevbToDIK()

' vbtodik(vbKeyCancel    3 CANCEL key

  vbToDIK(vbKeyBack) = DIK_BACKSPACE
  vbToDIK(vbKeyTab) = DIK_TAB
  ' vbtodik(vbKeyClear    12 CLEAR key         ??
  vbToDIK(vbKeyReturn) = DIK_RETURN
  vbToDIK(vbKeyShift) = DIK_LSHIFT
  vbToDIK(vbKeyControl) = DIK_LCONTROL
  vbToDIK(vbKeyMenu) = DIK_LMENU
  vbToDIK(vbKeyPause) = DIK_PAUSE
  vbToDIK(vbKeyCapital) = DIK_CAPSLOCK
  vbToDIK(vbKeyEscape) = DIK_ESCAPE
  vbToDIK(vbKeySpace) = DIK_SPACE
  vbToDIK(vbKeyPageUp) = DIK_PGUP
  vbToDIK(vbKeyPageDown) = DIK_PGDN
  vbToDIK(vbKeyEnd) = DIK_END
  vbToDIK(vbKeyHome) = DIK_HOME
  vbToDIK(vbKeyLeft) = DIK_LEFTARROW
  vbToDIK(vbKeyUp) = DIK_UPARROW
  vbToDIK(vbKeyRight) = DIK_RIGHTARROW
  vbToDIK(vbKeyDown) = DIK_DOWNARROW
  ' vbtodik(vbKeySelect   41 SELECT key        ??
  vbToDIK(vbKeyPrint) = DIK_SYSRQ
  ' vbtodik(vbKeyExecute  43 EXECUTE key       ??
  vbToDIK(vbKeySnapshot) = DIK_SYSRQ
  vbToDIK(vbKeyInsert) = DIK_INSERT
  vbToDIK(vbKeyDelete) = DIK_DELETE
  ' vbtodik(vbKeyHelp     47 HELP key          ??
  vbToDIK(vbKey0) = DIK_0
  vbToDIK(vbKey1) = DIK_1
  vbToDIK(vbKey2) = DIK_2
  vbToDIK(vbKey3) = DIK_3
  vbToDIK(vbKey4) = DIK_4
  vbToDIK(vbKey5) = DIK_5
  vbToDIK(vbKey6) = DIK_6
  vbToDIK(vbKey7) = DIK_7
  vbToDIK(vbKey8) = DIK_8
  vbToDIK(vbKey9) = DIK_9

  vbToDIK(vbKeyA) = DIK_A
  vbToDIK(vbKeyB) = DIK_B
  vbToDIK(vbKeyC) = DIK_C
  vbToDIK(vbKeyD) = DIK_D
  vbToDIK(vbKeyE) = DIK_E
  vbToDIK(vbKeyF) = DIK_F
  vbToDIK(vbKeyG) = DIK_G
  vbToDIK(vbKeyH) = DIK_H
  vbToDIK(vbKeyI) = DIK_I
  vbToDIK(vbKeyJ) = DIK_J
  vbToDIK(vbKeyK) = DIK_K
  vbToDIK(vbKeyL) = DIK_L
  vbToDIK(vbKeyM) = DIK_M
  vbToDIK(vbKeyN) = DIK_N
  vbToDIK(vbKeyO) = DIK_O
  vbToDIK(vbKeyP) = DIK_P
  vbToDIK(vbKeyQ) = DIK_Q
  vbToDIK(vbKeyR) = DIK_R
  vbToDIK(vbKeys) = DIK_S
  vbToDIK(vbKeyT) = DIK_T
  vbToDIK(vbKeyU) = DIK_U
  vbToDIK(vbKeyV) = DIK_V
  vbToDIK(vbKeyW) = DIK_W
  vbToDIK(vbKeyX) = DIK_X
  vbToDIK(vbKeyY) = DIK_Y
  vbToDIK(vbKeyZ) = DIK_Z

  vbToDIK(vbKeyNumpad0) = DIK_NUMPAD0
  vbToDIK(vbKeyNumpad1) = DIK_NUMPAD1
  vbToDIK(vbKeyNumpad2) = DIK_NUMPAD2
  vbToDIK(vbKeyNumpad3) = DIK_NUMPAD3
  vbToDIK(vbKeyNumpad4) = DIK_NUMPAD4
  vbToDIK(vbKeyNumpad5) = DIK_NUMPAD5
  vbToDIK(vbKeyNumpad6) = DIK_NUMPAD6
  vbToDIK(vbKeyNumpad7) = DIK_NUMPAD7
  vbToDIK(vbKeyNumpad8) = DIK_NUMPAD8
  vbToDIK(vbKeyNumpad9) = DIK_NUMPAD9
  vbToDIK(vbKeyMultiply) = DIK_NUMPADSTAR

  vbToDIK(vbKeyAdd) = DIK_NUMPADPLUS
  vbToDIK(vbKeySeparator) = DIK_NUMPADENTER
  vbToDIK(vbKeySubtract) = DIK_NUMPADMINUS
  vbToDIK(vbKeyDecimal) = DIK_NUMPADPERIOD = 83
  vbToDIK(vbKeyDivide) = DIK_NUMPADSLASH
  vbToDIK(vbKeyF1) = DIK_F1
  vbToDIK(vbKeyF2) = DIK_F2
  vbToDIK(vbKeyF3) = DIK_F3
  vbToDIK(vbKeyF4) = DIK_F4
  vbToDIK(vbKeyF5) = DIK_F5
  vbToDIK(vbKeyF6) = DIK_F6
  vbToDIK(vbKeyF7) = DIK_F7
  vbToDIK(vbKeyF8) = DIK_F8
  vbToDIK(vbKeyF9) = DIK_F9
  vbToDIK(vbKeyF10) = DIK_F10
  vbToDIK(vbKeyF11) = DIK_F11
  vbToDIK(vbKeyF12) = DIK_F12
  vbToDIK(vbKeyF13) = DIK_F13
  vbToDIK(vbKeyF14) = DIK_F14
  vbToDIK(vbKeyF15) = DIK_F15
  'vbKeyF16 127 F16 key
  vbToDIK(vbKeyNumlock) = DIK_NUMLOCK

  vbToDIK(vbKeyALT) = DIK_LALT
  vbToDIK(vbKeyGrave) = DIK_GRAVE
  vbToDIK(vbKeyMinus) = DIK_MINUS
  vbToDIK(vbKeyEquals) = DIK_EQUALS
  vbToDIK(vbKeyLBracket) = DIK_LBRACKET
  vbToDIK(vbKeyRBracket) = DIK_RBRACKET
  vbToDIK(vbKeyBackSlash) = DIK_BACKSLASH
  vbToDIK(vbKeySemiColon) = DIK_SEMICOLON
  vbToDIK(vbKeyApostrophe) = DIK_APOSTROPHE
  vbToDIK(vbKeyComma) = DIK_COMMA
  vbToDIK(vbKeyPeriod) = DIK_PERIOD
  vbToDIK(vbKeySlash) = DIK_SLASH
  vbToDIK(vbKeyLWin) = DIK_LWIN
  vbToDIK(vbRWin) = DIK_RWIN
  vbToDIK(vbScrollLock) = DIK_SCROLL

End Sub

'Visual Basic Reference
'Key Code Constants
'Key Codes
'Constant Value Description
'vbKeyCancel    3 CANCEL key
'vbKeyBack      8 BACKSPACE key     DIK_BACKSPACE = 14    DIK_BACK = 14
'vbKeyTab       9 TAB key           DIK_TAB = 15
'vbKeyClear    12 CLEAR key         ??
'vbKeyReturn   13 ENTER key         DIK_RETURN = 28        '(&H1C)
'vbKeyShift    16 SHIFT key         DIK_LSHIFT = 42        '(&H2A) DIK_RSHIFT = 54        '(&H36)
'vbKeyControl  17 CTRL key          DIK_LCONTROL = 29      '(&H1D) DIK_RCONTROL = 157 '(&H9D)
'vbKeyMenu     18 MENU key          DIK_LMENU = 56         '(&H38) DIK_RMENU = 184        '(&HB8)
'vbKeyPause    19 PAUSE key         DIK_PAUSE = 197        '(&HC5)
'vbKeyCapital  20 CAPS LOCK key     DIK_CAPSLOCK = 58      '(&H3A) DIK_CAPITAL = 58       '(&H3A)
'vbKeyEscape   27 ESC key           DIK_ESCAPE = 1
'vbKeySpace    32 SPACEBAR key      DIK_SPACE = 57         '(&H39)
'vbKeyPageUp   33 PAGE UP key       DIK_PGUP = 201         '(&HC9)
'vbKeyPageDown 34 PAGE DOWN key     DIK_PGDN = 209         '(&HD1)
'vbKeyEnd      35 END key           DIK_END = 207          '(&HCF)
'vbKeyHome     36 HOME key          DIK_HOME = 199         '(&HC7)
'vbKeyLeft     37 LEFT ARROW key    DIK_LEFTARROW = 203    '(&HCB)    DIK_LEFT = 203         '(&HCB)
'vbKeyUp       38 UP ARROW key      DIK_UPARROW = 200      '(&HC8)    DIK_UP = 200           '(&HC8)
'vbKeyRight    39 RIGHT ARROW key   DIK_RIGHTARROW = 205   '(&HCD)    DIK_RIGHT = 205        '(&HCD)
'vbKeyDown     40 DOWN ARROW key    DIK_DOWNARROW = 208    '(&HD0)    DIK_DOWN = 208         '(&HD0)
'vbKeySelect   41 SELECT key        ??
'vbKeyPrint    42 PRINT SCREEN key  DIK_SYSRQ = 183        '(&HB7)
'vbKeyExecute  43 EXECUTE key       ??
'vbKeySnapshot 44 SNAPSHOT key      DIK_SYSRQ = 183        '(&HB7)
'vbKeyInsert   45 INS key           DIK_INSERT = 210       '(&HD2)
'vbKeyDelete   46 DEL key           DIK_DELETE = 211       '(&HD3)
'vbKeyHelp     47 HELP key          ??

'Key0 Through Key9 Are the Same as Their ASCII Equivalents: '0' Through '9
'Constant Value Description
'vbKey0 48 0 key    DIK_0 = 11
'vbKey1 49 1 key    DIK_1 = 2
'vbKey2 50 2 key    DIK_2 = 3
'vbKey3 51 3 key    DIK_3 = 4
'vbKey4 52 4 key    DIK_4 = 5
'vbKey5 53 5 key    DIK_5 = 6
'vbKey6 54 6 key    DIK_6 = 7
'vbKey7 55 7 key    DIK_7 = 8
'vbKey8 56 8 key    DIK_8 = 9
'vbKey9 57 9 key    DIK_9 = 10

'KeyA Through KeyZ Are the Same as Their ASCII Equivalents: 'A' Through 'Z'
'Constant Value Description
'vbKeyA 65 A key    DIK_A = 30             '(&H1E)
'vbKeyB 66 B key    DIK_B = 48             '(&H30)
'vbKeyC 67 C key    DIK_C = 46             '(&H2E)
'vbKeyD 68 D key    DIK_D = 32             '(&H20)
'vbKeyE 69 E key    DIK_E = 18             '(&H12)
'vbKeyF 70 F key    DIK_F = 33             '(&H21)
'vbKeyG 71 G key    DIK_G = 34             '(&H22)
'vbKeyH 72 H key    DIK_H = 35             '(&H23)
'vbKeyI 73 I key    DIK_I = 23             '(&H17)
'vbKeyJ 74 J key    DIK_J = 36             '(&H24)
'vbKeyK 75 K key    DIK_K = 37             '(&H25)
'vbKeyL 76 L key    DIK_L = 38             '(&H26)
'vbKeyM 77 M key    DIK_M = 50             '(&H32)
'vbKeyN 78 N key    DIK_N = 49             '(&H31)
'vbKeyO 79 O key    DIK_O = 24             '(&H18)
'vbKeyP 80 P key    DIK_P = 25             '(&H19)
'vbKeyQ 81 Q key    DIK_Q = 16             '(&H10)
'vbKeyR 82 R key    DIK_R = 19             '(&H13)
'vbKeyS 83 S key    DIK_S = 31             '(&H1F)
'vbKeyT 84 T key    DIK_T = 20             '(&H14)
'vbKeyU 85 U key    DIK_U = 22             '(&H16)
'vbKeyV 86 V key    DIK_V = 47             '(&H2F)
'vbKeyW 87 W key    DIK_W = 17             '(&H11)
'vbKeyX 88 X key    DIK_X = 45             '(&H2D)
'vbKeyY 89 Y key    DIK_Y = 21             '(&H15)
'vbKeyZ 90 Z key    DIK_Z = 44             '(&H2C)

'Keys on the Numeric Keypad
'Constant Value Description
'vbKeyNumpad0    96 0 key                       DIK_NUMPAD0 = 82       '(&H52)
'vbKeyNumpad1    97 1 key                       DIK_NUMPAD1 = 79       '(&H4F)
'vbKeyNumpad2    98 2 key                       DIK_NUMPAD2 = 80       '(&H50)
'vbKeyNumpad3    99 3 key                       DIK_NUMPAD3 = 81       '(&H51)
'vbKeyNumpad4   100 4 key                       DIK_NUMPAD4 = 75       '(&H4B)
'vbKeyNumpad5   101 5 key                       DIK_NUMPAD5 = 76       '(&H4C)
'vbKeyNumpad6   102 6 key                       DIK_NUMPAD6 = 77       '(&H4D)
'vbKeyNumpad7   103 7 key                       DIK_NUMPAD7 = 71       '(&H47)
'vbKeyNumpad8   104 8 key                       DIK_NUMPAD8 = 72       '(&H48)
'vbKeyNumpad9   105 9 key                       DIK_NUMPAD9 = 73       '(&H49)
'vbKeyMultiply  106 MULTIPLICATION SIGN (*) key DIK_NUMPADSTAR = 55    '(&H37)    DIK_MULTIPLY = 55      '(&H37)

'vbKeyAdd       107 PLUS SIGN (+) key           DIK_NUMPADPLUS = 78    '(&H4E)
'vbKeySeparator 108 ENTER (keypad) key          DIK_NUMPADENTER = 156  '(&H9C)
'vbKeySubtract  109 MINUS SIGN (-) key          DIK_NUMPADMINUS = 74   '(&H4A)    DIK_SUBTRACT = 74      '(&H4A)
'vbKeyDecimal   110 DECIMAL POINT(.) key        DIK_NUMPADPERIOD = 83  '(&H53)    DIK_DECIMAL = 83       '(&H53)
'vbKeyDivide    111 DIVISION SIGN (/) key       DIK_NUMPADSLASH = 181  '(&HB5)    DIK_DIVIDE = 181       '(&HB5)

'Function Keys
'Constant Value Description
'vbKeyF1  112  F1 key    DIK_F1 = 59            '(&H3B)
'vbKeyF2  113  F2 key    DIK_F2 = 60            '(&H3C)
'vbKeyF3  114  F3 key    DIK_F3 = 61            '(&H3D)
'vbKeyF4  115  F4 key    DIK_F4 = 62            '(&H3E)
'vbKeyF5  116  F5 key    DIK_F5 = 63            '(&H3F)
'vbKeyF6  117  F6 key    DIK_F6 = 64            '(&H40)
'vbKeyF7  118  F7 key    DIK_F7 = 65            '(&H41)
'vbKeyF8  119  F8 key    DIK_F8 = 66            '(&H42)
'vbKeyF9  120  F9 key    DIK_F9 = 67            '(&H43)
'vbKeyF10 121 F10 key    DIK_F10 = 68           '(&H44)
'vbKeyF11 122 F11 key    DIK_F11 = 87           '(&H57)
'vbKeyF12 123 F12 key    DIK_F12 = 88           '(&H58)
'vbKeyF13 124 F13 key    DIK_F13 = 100          '(&H64)
'vbKeyF14 125 F14 key    DIK_F14 = 101          '(&H65)
'vbKeyF15 126 F15 key    DIK_F15 = 102          '(&H66)
'vbKeyF16 127 F16 key

'NEW CODES                unshifted                         shifted
'vbKeyALT = vbKeyMenu DIK_LALT = 56          '(&H38)    DIK_RALT = 184         '(&HB8)
'vbKeyGrave      xC0  DIK_GRAVE = 41         '(&H29)    DIK_CIRCUMFLEX = 144   '(&H90)
'vbKeyMinus      xBD  DIK_MINUS = 12                    DIK_UNDERLINE = 147    '(&H93)
'vbKeyEquals     xBB  DIK_EQUALS = 13                   DIK_ADD = 78           '(&H4E)
'vbKeyLBracket   xDB  DIK_LBRACKET = 26      '(&H1A)
'vbKeyRBracket   xDD  DIK_RBRACKET = 27      '(&H1B)
'vbKeyBackSlash  xDC  DIK_BACKSLASH = 43     '(&H2B)
'vbKeySemiColon  xBA  DIK_SEMICOLON = 39     '(&H27)    DIK_COLON = 146        '(&H92)
'vbKeyApostrophe xDE  DIK_APOSTROPHE = 40    '(&H28)
'vbKeyComma      xBC  DIK_COMMA = 51         '(&H33)
'vbKeyPeriod     xBE  DIK_PERIOD = 52        '(&H34)
'vbKeySlash      xBF  DIK_SLASH = 53         '(&H35)
'vbKeyLWin       x5B  DIK_LWIN = 219         '(&HDB)
'vbRWin          x5C  DIK_RWIN = 220         '(&HDC)
'vbScrollLock    x91  DIK_SCROLL = 70        '(&H46)
'vbKeyNumlock    144  NUM LOCK key        DIK_NUMLOCK = 69       '(&H45)

'    DIK_ABNT_C1 = 115      '(&H73)
'    DIK_ABNT_C2 = 126      '(&H7E)
'    DIK_APPS = 221         '(&HDD)
'    DIK_AT = 145           '(&H91)
'    DIK_AX = 150           '(&H96)
'    DIK_CALCULATOR = 161   '(&HA1)
'    DIK_CONVERT = 121      '(&H79)
'    DIK_KANA = 112         '(&H70)
'    DIK_KANJI = 148        '(&H94)
'    DIK_MAIL = 236         '(&HEC)
'    DIK_MEDIASELECT = 237  '(&HED)
'    DIK_MEDIASTOP = 164    '(&HA4)
'    DIK_MUTE = 160         '(&HA0)
'    DIK_MYCOMPUTER = 235   '(&HEB)
'    DIK_NEXT = 209         '(&HD1)
'    DIK_NEXTTRACK = 153    '(&H99)
'    DIK_NOCONVERT = 123    '(&H7B)
'    DIK_NUMPADCOMMA = 179  '(&HB3)
'    DIK_NUMPADEQUALS = 141 '(&H8D)
'    DIK_OEM_102 = 86       '(&H56)
'    DIK_PLAYPAUSE = 162    '(&HA2)
'    DIK_POWER = 222        '(&HDE)
'    DIK_PREVTRACK = 144    '(&H90)
'    DIK_PRIOR = 201        '(&HC9)
'    DIK_SLEEP = 223        '(&HDF)
'    DIK_STOP = 149         '(&H95)
'    DIK_UNLABELED = 151    '(&H97)
'    DIK_VOLUMEDOWN = 174   '(&HAE)
'    DIK_VOLUMEUP = 176     '(&HB0)
'    DIK_WAKE = 227         '(&HE3)
'    DIK_WEBBACK = 234      '(&HEA)
'    DIK_WEBFAVORITES = 230 '(&HE6)
'    DIK_WEBFORWARD = 233   '(&HE9)
'    DIK_WEBHOME = 178      '(&HB2)
'    DIK_WEBREFRESH = 231   '(&HE7)
'    DIK_WEBSEARCH = 229    '(&HE5)
'    DIK_WEBSTOP = 232      '(&HE8)
'    DIK_YEN = 125          '(&H7D)

'Platform SDK: Windows User Interface
'Virtual-Key Codes
'The following table shows the symbolic constant names, hexadecimal values, and mouse or keyboard equivalents for the virtual-key codes used by the system. The codes are listed in numeric order.

'Symbolic constant name Value
'(hexadecimal) Mouse or keyboard equivalent
'VK_LBUTTON   01 Left mouse button
'VK_RBUTTON   02 Right mouse button
'VK_CANCEL    03 Control-break processing
'VK_MBUTTON   04 Middle mouse button (three-button mouse)
'VK_XBUTTON1  05 Win2K or later: X1 mouse button
'VK_XBUTTON2  06 Win2K or later: X2 mouse button
'—            07 Undefined
'VK_BACK      08 BACKSPACE key
'VK_TAB       09 TAB key
'—          0A–0B Reserved
'VK_CLEAR     0C CLEAR key
'VK_RETURN    0D ENTER key
'—          0E–0F Undefined
'VK_SHIFT     10 SHIFT key
'VK_CONTROL   11 CTRL key
'VK_MENU      12 ALT key
'VK_PAUSE     13 PAUSE key
'VK_CAPITAL   14 CAPS LOCK key
'VK_KANA      15 IME Kana mode
'VK_HANGUEL   15 IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
'VK_HANGUL    15 IME Hangul mode
'—            16 Undefined
'VK_JUNJA     17 IME Junja mode
'VK_FINAL     18 IME final mode
'VK_HANJA     19 IME Hanja mode
'VK_KANJI     19 IME Kanji mode
'—            1A Undefined
'VK_ESCAPE    1B ESC key
'VK_CONVERT   1C IME convert
'VK_NONCONVERT 1D IME nonconvert
'VK_ACCEPT    1E IME accept
'VK_MODECHANGE 1F IME mode change request
'VK_SPACE     20 SPACEBAR
'VK_PRIOR     21 PAGE UP key
'VK_NEXT      22 PAGE DOWN key
'VK_END       23 END key
'VK_HOME      24 HOME key
'VK_LEFT      25 LEFT ARROW key
'VK_UP        26 UP ARROW key
'VK_RIGHT     27 RIGHT ARROW key
'VK_DOWN      28 DOWN ARROW key
'VK_SELECT    29 SELECT key
'VK_PRINT     2A PRINT key
'VK_EXECUTE   2B EXECUTE key
'VK_SNAPSHOT  2C PRINT SCREEN key
'VK_INSERT    2D INS key
'VK_DELETE    2E DEL key
'VK_HELP      2F HELP key
'             30 0 key
'             31 1 key
'             32 2 key
'             33 3 key
'             34 4 key
'             35 5 key
'             36 6 key
'             37 7 key
'             38 8 key
'             39 9 key
'—          3A–40 Undefined
'             41  a Key
'             42  b Key
'             43  c Key
'             44  d Key
'             45  E Key
'             46  F Key
'             47  g Key
'             48  H Key
'             49  i Key
'             4A J key
'             4B K key
'             4C L key
'             4D M key
'             4E N key
'             4F O key
'             50  P Key
'             51  q Key
'             52  r Key
'             53  S Key
'             54  T Key
'             55  U Key
'             56  V Key
'             57  w Key
'             58  x Key
'             59  y Key
'             5A Z key

'VK_LWIN      5B Left Windows key (Microsoft® Natural® keyboard)
'VK_RWIN      5C Right Windows key (Natural keyboard)
'VK_APPS      5D Applications key (Natural keyboard)
'—            5E Reserved
'VK_SLEEP     5F Computer Sleep key
'VK_NUMPAD0   60 Numeric keypad 0 key
'VK_NUMPAD1   61 Numeric keypad 1 key
'VK_NUMPAD2   62 Numeric keypad 2 key
'VK_NUMPAD3   63 Numeric keypad 3 key
'VK_NUMPAD4   64 Numeric keypad 4 key
'VK_NUMPAD5   65 Numeric keypad 5 key
'VK_NUMPAD6   66 Numeric keypad 6 key
'VK_NUMPAD7   67 Numeric keypad 7 key
'VK_NUMPAD8   68 Numeric keypad 8 key
'VK_NUMPAD9   69 Numeric keypad 9 key
'VK_MULTIPLY  6A Multiply key
'VK_ADD       6B Add key
'VK_SEPARATOR 6C Separator key
'VK_SUBTRACT  6D Subtract key
'VK_DECIMAL   6E Decimal key
'VK_DIVIDE    6F Divide key

'VK_F1        70 F1 key
'VK_F2        71 F2 key
'VK_F3        72 F3 key
'VK_F4        73 F4 key
'VK_F5        74 F5 key
'VK_F6        75 F6 key
'VK_F7        76 F7 key
'VK_F8        77 F8 key
'VK_F9        78 F9 key
'VK_F10       79 F10 key
'VK_F11       7A F11 key
'VK_F12       7B F12 key
'VK_F13       7C F13 key
'VK_F14       7D F14 key
'VK_F15       7E F15 key
'VK_F16       7F F16 key

'VK_F17       80H F17 key
'VK_F18       81H F18 key
'VK_F19       82H F19 key
'VK_F20       83H F20 key
'VK_F21       84H F21 key
'VK_F22       85H F22 key
'VK_F23       86H F23 key
'VK_F24       87H F24 key
'—           88–8F Unassigned
'VK_NUMLOCK   90 NUM LOCK key
'VK_SCROLL    91 SCROLL LOCK key
'           92–96 OEM specific
'—          97–9F Unassigned

'VK_LSHIFT              A0 Left SHIFT key
'VK_RSHIFT              A1 Right SHIFT key
'VK_LCONTROL            A2 Left CONTROL key
'VK_RCONTROL            A3 Right CONTROL key
'VK_LMENU               A4 Left MENU key
'VK_RMENU               A5 Right MENU key
'VK_BROWSER_BACK        A6 Win2K or later: Browser Back key
'VK_BROWSER_FORWARD     A7 Win2K or later: Browser Forward key
'VK_BROWSER_REFRESH     A8 Win2K or later: Browser Refresh key
'VK_BROWSER_STOP        A9 Win2K or later: Browser Stop key
'VK_BROWSER_SEARCH      AA Win2K or later: Browser Search key
'VK_BROWSER_FAVORITES   AB Win2K or later: Browser Favorites key
'VK_BROWSER_HOME        AC Win2K or later: Browser Start and Home key
'VK_VOLUME_MUTE         AD Win2K or later: Volume Mute key
'VK_VOLUME_DOWN         AE Win2K or later: Volume Down key
'VK_VOLUME_UP           AF Win2K or later: Volume Up key

'VK_MEDIA_NEXT_TRACK    B0 Win2K or later: Next Track key
'VK_MEDIA_PREV_TRACK    B1 Win2K or later: Previous Track key
'VK_MEDIA_STOP          B2 Win2K or later: Stop Media key
'VK_MEDIA_PLAY_PAUSE    B3 Win2K or later: Play/Pause Media key
'VK_LAUNCH_MAIL         B4 Win2K or later: Start Mail key
'VK_LAUNCH_MEDIA_SELECT B5 Win2K or later: Select Media key
'VK_LAUNCH_APP1         B6 Win2K or later: Start Application 1 key
'VK_LAUNCH_APP2         B7 Win2K or later: Start Application 2 key
'—                    B8-B9 Reserved

'VK_OEM_1               BA Win2K or later: For the US standard keyboard, the ';:' key
'VK_OEM_PLUS            BB Win2K or later: For any country/region, the '+' key
'VK_OEM_COMMA           BC Win2K or later: For any country/region, the ',' key
'VK_OEM_MINUS           BD Win2K or later: For any country/region, the '-' key
'VK_OEM_PERIOD          BE Win2K or later: For any country/region, the '.' key
'VK_OEM_2               BF Win2K or later: For the US standard keyboard, the '/?' key
'VK_OEM_3               C0 Win2K or later: For the US standard keyboard, the '`~' key
'—                    C1–D7 Reserved
'—                    D8–DA Unassigned
'VK_OEM_4               DB Win2K or later: For the US standard keyboard, the '[{' key
'VK_OEM_5               DC Win2K or later: For the US standard keyboard, the '\|' key
'VK_OEM_6               DD Win2K or later: For the US standard keyboard, the ']}' key
'VK_OEM_7               DE Win2K or later: For the US standard keyboard, the 'single-quote/double-quote' key
'VK_OEM_8               DF

'—                      E0 Reserved
'                       E1 OEM specific
'VK_OEM_102             E2 Win2K or later: Either the angle bracket key or the backslash key on the RT 102-key keyboard
'                     E3–E4 OEM specific
'VK_PROCESSKEY          E5 Windows 95/98, Windows NT 4.0, Win2K or later: IME PROCESS key
'                       E6 OEM specific
'VK_PACKET              E7 Win2K or later: Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP
'—                      E8 Unassigned
'                     E9–F5 OEM specific
'VK_ATTN                F6 Attn key
'VK_CRSEL               F7 CrSel key
'VK_EXSEL               F8 ExSel key
'VK_EREOF               F9 Erase EOF key
'VK_PLAY                FA Play key
'VK_ZOOM                FB Zoom key
'VK_NONAME              FC Reserved for future use
'VK_PA1                 FD PA1 key
'VK_OEM_CLEAR           FE Clear key

'Platform SDK Release: February 2001

':) Ulli's VB Code Formatter V2.13.5 (31-Jan-03 12:21:34) 25 + 503 = 528 Lines
