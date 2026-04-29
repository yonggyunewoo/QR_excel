Option Explicit
Dim mat() As Byte ' matrix of QR

' QR Code 2005 bar code symbol creation according ISO/IEC 18004:2006
'   param text to encode
'   param level optional: quality level LMQH
'   param version optional: minimum version size (-3:M1, -2:M2, .. 1, .. 40)
'   creates QR and micro QR bar code symbol as shape in Excel cell.
'  Kanji mode needs the custom property 'kanji' of the Application.Caller sheet to convert from unicode to kanji
'   the string contains the 6879 chars of Kanji followed by the 6879 equivalent unicode chars
Public Function QRCodeImage(text As String, Optional level As String, Optional version As Integer = 1) As String
' 입력받은 텍스트를 QR 코드로 변환하는 함수
' QR 코드는 엑셀의 이미지 형식으로 생성 (temp 폴더에 임시파일 저장)
' 임시파일은 2진 bmp 형식으로 생성하여 용량 Issue 최소화
' ARG:
    ' text: QR 코드로 변환할 텍스트 문자열
    ' level: QR 코드의 오류 수정 수준 (L, M, Q, H 중 하나)
    ' version: QR 코드의 버전 (1부터 40까지, 또는 -3부터 -1까지의 마이크로 QR 버전)
On Error GoTo failed
Dim tStart As Double: tStart = Timer

If Not TypeOf Application.Caller Is Range Then Err.Raise 513, "QR code", "Call only from sheet"
Dim mode As Byte, lev As Byte, s As Long, a As Long, blk As Long, ec As Long
Dim i As Long, j As Long, k As Long, l As Long, c As Long, b As Long, txt As String
Dim w As Long, x As Long, y As Long, v As Double, el As Long, eb As Long
Dim shp As Shape, m As Long, p As Variant, ecw As Variant, ecb As Variant
Dim k1 As String, k2 As String, fColor As Long, bColor As Long, line As Long
Const alpha = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:"

fColor = vbBlack: bColor = vbBlack: line = xlHairline ' redraw graphic ?
For Each shp In Application.Caller.Parent.Shapes
    If shp.Name = Application.Caller.Address Then
        Dim prevTitle As String
        prevTitle = ""
        Dim prevAlt As String
        prevAlt = ""
        On Error Resume Next
        prevTitle = shp.Title
        prevAlt = shp.AlternativeText
        fColor = shp.Fill.ForeColor.RGB
        bColor = shp.line.ForeColor.RGB
        line = shp.line.Weight
        On Error GoTo failed
        If Left(prevAlt, Len(text) + 1) = text & "|" Then Exit Function ' same as prev ?
        shp.Delete
    End If
Next shp
For Each ecw In ActiveWorkbook.Worksheets
    For Each p In ecw.CustomProperties ' look for kanji conversion string
        If p.Name = "kanji" Then If Len(p.Value) > 10000 Then k1 = p.Value
    Next p
Next ecw
lev = (InStr("LMQHlmqh0123", level) - 1) And 3
For i = 1 To Len(text) ' compute mode
    c = AscW(Mid(text, i, 1))
    If c < 48 Or c > 57 Then
        If mode = 0 Then mode = 1 ' alphanumeric mode
        If InStr(alpha, ChrW(c)) = 0 Then
            If mode = 1 Then mode = 2 ' binary or kanji ?
            If c < 32 Or c > 126 Then
                If InStr(Len(k1) / 2 + 1, k1, ChrW(c)) = 0 Then mode = 2: Exit For ' binary
                mode = 3 ' kanji
            End If
        End If
    End If
Next i
txt = IIf(mode = 2, utf16to8(text), text) ' for reader conformity
l = Len(txt)
w = Int(l * Array(10 / 3, 11 / 2, 8, 13)(mode) + 0.5) ' 3 digits in 10 bits, 2 chars in 11 bits, 1 byte, 13 bits/byte
p = Array(Array(10, 12, 14), Array(9, 11, 13), Array(8, 16, 16), Array(8, 10, 12))(mode) ' # of bits of count indicator
' error correction words L,M,Q,H and blocks L,M,Q,H for all version sizes (99=N/A)
ecw = Array(Array(2, 5, 6, 8, 7, 10, 15, 20, 26, 18, 20, 24, 30, 18, 20, 24, 26, 30, 22, 24, 28, 30, 28, 28, 28, 28, 30, 30, 26, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30), _
    Array(99, 6, 8, 10, 10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26, 26, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28), _
    Array(99, 99, 99, 14, 13, 22, 18, 26, 18, 24, 18, 22, 20, 24, 28, 26, 24, 20, 30, 24, 28, 28, 26, 30, 28, 30, 30, 30, 30, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30), _
    Array(99, 99, 99, 99, 17, 28, 22, 16, 22, 28, 26, 26, 24, 28, 24, 28, 22, 24, 24, 30, 28, 28, 26, 28, 30, 24, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30))
ecb = Array(Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 4, 4, 4, 4, 4, 6, 6, 6, 6, 7, 8, 8, 9, 9, 10, 12, 12, 12, 13, 14, 15, 16, 17, 18, 19, 19, 20, 21, 22, 24, 25), _
    Array(1, 1, 1, 1, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31, 33, 35, 37, 38, 40, 43, 45, 47, 49), _
    Array(1, 1, 1, 1, 1, 1, 2, 2, 4, 4, 6, 6, 8, 8, 8, 10, 12, 16, 12, 17, 16, 18, 21, 20, 23, 23, 25, 27, 29, 34, 34, 35, 38, 40, 43, 45, 48, 51, 53, 56, 59, 62, 65, 68), _
    Array(1, 1, 1, 1, 1, 1, 2, 4, 4, 4, 5, 6, 8, 8, 11, 11, 16, 16, 18, 16, 19, 21, 25, 25, 25, 34, 30, 32, 35, 37, 40, 42, 45, 48, 51, 54, 57, 60, 63, 66, 70, 74, 77, 81))
version = IIf(version < mode - 3, mode - 3, version) - 1
Do ' compute QR size
    version = version + 1
    If version + 3 > UBound(ecb(0)) Then Err.Raise 515, "QRCode", "Message too long"
    s = version * IIf(version < 1, 2, 4) + 17 ' symbol size
    j = ecb(lev)(version + 3) * ecw(lev)(version + 3)   ' error correction
    a = IIf(version < 2, 0, version \ 7 + 2) ' # of align pattern
    el = (s - 1) * (s - 1) - (5 * a - 1) * (5 * a - 1) ' total bits - align - timing
    el = el - IIf(version < 1, 59, IIf(version < 2, 191, IIf(version < 7, 136, 172))) ' finder, version, format
    k = IIf(version < 1, version + (19 - 2 * mode) \ 3, p((version + 7) \ 17)) ' count indcator bits
    i = IIf(version < 1, version + (version And 1) * 4 + 3, 4) ' mode indicator bits, M1+M3: +4 bits
Loop While (el And -8) - 8 * j < w + i + k
For lev = lev To 2 ' increase security level if data still fits
    j = ecb(lev + 1)(version + 3) * ecw(lev + 1)(version + 3)
    If (el And -8) - 8 * j < w + i + k Then Exit For
Next lev
blk = ecb(lev)(version + 3) ' # of error correction blocks
ec = ecw(lev)(version + 3) ' # of error correction bytes
el = el \ 8 - ec * blk ' data capacity
w = el \ blk ' # of words in group 1
b = blk + w * blk - el ' # of blocks in group 1

ReDim enc(el + ec * blk) As Byte, mat(s - 1, s - 1) As Byte
c = 0 ' encode head indicator bits
If version > 0 Then v = 2 ^ mode: eb = 4 Else v = mode: eb = version + 3 ' mode indicator
eb = eb + k: v = v * 2 ^ k + l ' character count indicator
For i = 1 To l ' encode data
    Select Case mode
    Case 0: ' numeric
        v = v * IIf(i + 1 < l, 1024, IIf(i < l, 128, 16)) + Val(Mid(txt, i, 3))
        eb = eb + IIf(i + 1 < l, 10, 4 + 3 * (l - i)): i = i + 2
    Case 1: ' alphanumeric
        j = InStr(alpha, Mid(txt, i, 1)) - 1
        If i < l Then j = 45 * j + InStr(alpha, Mid(txt, i + 1, 1)) - 1
        v = v * IIf(i < l, 2048, 64) + j
        eb = eb + IIf(i < l, 11, 6): i = i + 1
    Case 2: ' binary
        v = v * 256 + AscW(Mid(txt, i, 1))
        eb = eb + 8
    Case 3: ' Kanji
        j = InStr(Len(k1) / 2 + 1, k1, Mid(txt, i, 1)) - Len(k1) / 2
        j = (AscW(Mid(k1, j, 1)) And &H3FFF) - 320 ' unicode to shift JIS X 2008
        v = v * 8192 + (j \ 256) * 192 + (j And 255) ' to 13 bit kanji
        eb = eb + 13
    End Select
    For eb = eb To 8 Step -8 ' add data to bit stream
        j = 2 ^ (eb - 8): enc(c) = v \ j
        v = v - enc(c) * j: c = c + 1
    Next eb
Next i
If el > c Then i = IIf(version > 0, 4, version + 6): v = v * 2 ^ i: eb = eb + i ' terminator
enc(c) = (v * 256) \ 2 ^ eb: c = c + 1: enc(c) = ((v * 65536) \ 2 ^ eb) And 255
If eb > 8 And el >= c Then c = c + 1 ' bit padding
If (version And -3) = -3 And el = c Then enc(c) = enc(c) \ 16 ' M1,M3: shift high bits to low nibble
i = 236
For c = c To el - 1 ' byte padding
    enc(c) = IIf((version And -3) = -3 And c = el - 1, 0, i)
    i = i Xor 236 Xor 17
Next c

ReDim rs(ec + 1) As Integer ' compute Reed Solomon error detection and correction
Dim lg(256) As Integer, ex(255) As Integer ' log/exp table
j = 1
For i = 0 To 254
    ex(i) = j: lg(j) = i ' compute log/exp table of Galois field
    j = j + j: If j > 255 Then j = j Xor 285 ' GF polynomial a^8+a^4+a^3+a^2+1 = 100011101b = 285
Next i
rs(0) = 1 ' compute RS generator polynomial
For i = 0 To ec - 1
    rs(i + 1) = 0
    For j = i + 1 To 1 Step -1
        rs(j) = rs(j) Xor ex((lg(rs(j - 1)) + i) Mod 255)
    Next j
Next i
eb = el: k = 0
For c = 1 To blk  ' compute RS correction data for each block
    For i = IIf(c <= b, 1, 0) To w
        x = enc(eb) Xor enc(k)
        For j = 1 To ec
            enc(eb + j - 1) = enc(eb + j) Xor IIf(x, ex((lg(rs(j)) + lg(x)) Mod 255), 0)
        Next j
        k = k + 1
    Next i
    eb = eb + ec
Next c

' fill QR matrix
For i = 8 To s - 1 ' timing pattern
    mat(i, IIf(version < 1, 0, 6)) = i And 1 Xor 3
    mat(IIf(version < 1, 0, 6), i) = i And 1 Xor 3
Next i
If version > 6 Then ' reserve version area
    For i = 0 To 17
        mat(i \ 3, s - 11 + i Mod 3) = 2
        mat(s - 11 + i Mod 3, i \ 3) = 2
    Next i
End If
If a < 2 Then a = IIf(version < 1, 1, 2)
For x = 1 To a ' layout finder/align pattern
    For y = 1 To a
        If x = 1 And y = 1 Then ' finder upper left
            i = 0: j = 0
            p = Array(383, 321, 349, 349, 349, 321, 383, 256, 511)
        ElseIf x = 1 And y = a Then  ' finder lower left
            i = 0: j = s - 8
            p = Array(256, 383, 321, 349, 349, 349, 321, 383)
        ElseIf x = a And y = 1 Then  ' finder upper right
            i = s - 8: j = 0
            p = Array(254, 130, 186, 186, 186, 130, 254, 0, 255)
        Else ' alignment grid
            c = 2 * Int(2 * (version + 1) / (1 - a)) ' pattern spacing
            i = IIf(x = 1, 4, s - 9 + c * (a - x))
            j = IIf(y = 1, 4, s - 9 + c * (a - y))
            p = Array(31, 17, 21, 17, 31) ' alignment pattern
        End If
        If version <> 1 Or x + y < 4 Then ' no align pattern for version 1
            For c = 0 To UBound(p) ' set fixed pattern, reserve space
                m = p(c): k = 0
                Do
                    mat(i + k, j + c) = (m And 1) Or 2
                    m = m \ 2: k = k + 1
                Loop While 2 ^ k <= p(0)
            Next c
        End If
    Next y
Next x
x = s: y = s - 1 ' layout codewords
For i = 0 To eb - 1
    c = 0: k = 0: j = w + 1 ' interleave data
    If i >= el Then
        c = el: k = el: j = ec ' interleave checkwords
    ElseIf i + blk - b >= el Then
        c = -b: k = c ' interleave group 2 last bytes
    ElseIf (i Mod blk) >= b Then
        c = -b ' interleave group 2
    Else
        j = j - 1 ' interleave group 1
    End If
    c = enc(c + ((i - k) Mod blk) * j + (i - k) \ blk) ' interleave data
    For j = IIf((-3 And version) = -3 And i = el - 1, 3, 7) To 0 Step -1 ' M1,M3: 4 bit
        k = IIf(version > 0 And x < 6, 1, 0) ' skip vertical timing pattern
        Do ' advance x,y
            x = x - 1
            If 1 And (x + 1) Xor k Then
                If s - x - k And 2 Then
                    If y > 0 Then y = y - 1: x = x + 2 ' up, top turn
                Else
                    If y < s - 1 Then y = y + 1: x = x + 2 ' down, bottom turn
                End If
            End If
        Loop While mat(x, y) And 2 ' skip reserved area
        If c And 2 ^ j Then mat(x, y) = 1
    Next j
Next i

m = 0: p = 1000000 ' data masking
For k = 0 To IIf(version < 1, 3, 7)
    If version < 1 Then ' penalty micro QR
        x = 1: y = 1
        For i = 1 To s - 1
            x = x - getPattern(i, s - 1, k, version)
            y = y - getPattern(s - 1, i, k, version)
        Next i
        j = IIf(x > y, 16 * x + y, x + 16 * y)
    Else ' penalty QR
        l = 0: k2 = "": j = 0
        For y = 0 To s - 1 ' horizontal
            c = 0: i = 0: k1 = "0000"
            For x = 0 To s - 1
                w = getPattern(x, y, k, version)
                l = l + w: k1 = k1 & w ' rule 4: count darks
                If c = w Then ' same as prev
                    i = i + 1
                    If x And Mid(k2, x + 4, 2) = c & c Then j = j + 3 ' rule 2: block 2x2
                Else
                    If i > 5 Then j = j + i - 2 ' rule 1: >5 adjacent
                    c = 1 - c: i = 1
                End If
            Next x
            If i > 5 Then j = j + i - 2 ' rule 1: >5 adjacent
            i = 0
            Do ' rule 3: like finder pattern
                i = InStr(i + 4, k1, "1011101")
                If i < 1 Then Exit Do
                If Mid(k1, i - 4, 4) = "0000" Or Mid(k1 & "0000", i + 7, 4) = "0000" Then j = j + 40
            Loop
            k2 = k1 ' rule 2: remember last line
        Next y
        For x = 0 To s - 1 ' vertical
            c = 0: i = 0: k1 = "0000"
            For y = 0 To s - 1
                w = getPattern(x, y, k, version)
                k1 = k1 & w ' vertical to string
                If c = w Then ' same as prev
                    i = i + 1
                Else
                    If i > 5 Then j = j + i - 2 ' rule 1: >5 adjacent
                    c = 1 - c: i = 1
                End If
            Next y
            If i > 5 Then j = j + i - 2 ' rule 1: >5 adjacent
            i = 0
            Do ' rule 3: like finder pattern
                i = InStr(i + 4, k1, "1011101")
                If i < 1 Then Exit Do
                If Mid(k1, i - 4, 4) = "0000" Or Mid(k1 & "0000", i + 7, 4) = "0000" Then j = j + 40
            Loop
        Next x
        j = j + Int(Abs(10 - 20 * l / (s * s))) * 10 ' rule 4: darks
    End If
    If j < p Then p = j: m = k ' take mask of lower penalty
Next k
' add format information, code level and mask
j = IIf(version = -3, m, IIf(version < 1, (2 * version + lev + 5) * 4 + m, ((5 - lev) And 3) * 8 + m))
j = j * 1024: k = j
For i = 4 To 0 Step -1 ' BCH error correction: 5 data, 10 error bits
    If j >= 1024 * 2 ^ i Then j = j Xor 1335 * 2 ^ i
Next i ' generator polynom: x^10+x^8+x^5+x^4+x^2+x+1 = 10100110111b = 1335
k = k Xor j Xor IIf(version < 1, 17477, 21522) ' XOR masking
For j = 0 To 14 ' layout format information
    If version < 1 Then
        mat(IIf(j < 8, 8, 15 - j), IIf(j < 8, j + 1, 8)) = k And 1 Xor 2 ' micro QR
    Else
        mat(IIf(j < 8, s - j - 1, IIf(j = 8, 7, 14 - j)), 8) = k And 1 Xor 2 ' QR horizontal
        mat(8, IIf(j < 6, j, IIf(j < 8, j + 1, s + j - 15))) = k And 1 Xor 2 ' vertical
    End If
    k = k \ 2
Next j
If version > 6 Then ' add version information
    k = version * 4096&
    For i = 5 To 0 Step -1 ' BCH error correction: 6 data, 12 error bits
        If k >= 4096 * 2 ^ i Then k = k Xor 7973 * 2 ^ i
    Next i ' generator polynom: x^12+x^11+x^10+x^9+x^8+x^5+x^2+1 = 1111100100101b = 7973
    k = k Xor (version * 4096&)
    For j = 0 To 17 ' layout version information
        mat(j \ 3, s + j Mod 3 - 11) = k And 1 Xor 2
        mat(s + j Mod 3 - 11, j \ 3) = k And 1 Xor 2
        k = k \ 2
    Next j
End If
' --- 1bpp BMP rendering ---
Const cellPx As Long = 10  ' pixels per module
Const margin As Long = 1   ' 0 =  no margin
Dim totalCells As Long: totalCells = s + margin * 2
Dim imgW As Long: imgW = totalCells * cellPx
Dim imgH As Long: imgH = totalCells * cellPx
Dim bmpStride As Long: bmpStride = ((imgW + 31) \ 32) * 4
Dim dataOffset As Long: dataOffset = 62
Dim dataSize As Long: dataSize = bmpStride * imgH
Dim fileSize As Long: fileSize = dataOffset + dataSize
Dim hdr(61) As Byte
' BITMAPFILEHEADER
hdr(0) = 66: hdr(1) = 77   ' "BM"
hdr(2) = fileSize And &HFF:          hdr(3) = (fileSize \ &H100) And &HFF
hdr(4) = (fileSize \ &H10000) And &HFF: hdr(5) = (fileSize \ &H1000000) And &HFF
hdr(6) = 0: hdr(7) = 0: hdr(8) = 0: hdr(9) = 0  ' reserved
hdr(10) = 62: hdr(11) = 0: hdr(12) = 0: hdr(13) = 0  ' bfOffBits = 62
' BITMAPINFOHEADER
hdr(14) = 40: hdr(15) = 0: hdr(16) = 0: hdr(17) = 0  ' biSize = 40
hdr(18) = imgW And &HFF:             hdr(19) = (imgW \ &H100) And &HFF
hdr(20) = (imgW \ &H10000) And &HFF: hdr(21) = (imgW \ &H1000000) And &HFF
hdr(22) = imgH And &HFF:             hdr(23) = (imgH \ &H100) And &HFF
hdr(24) = (imgH \ &H10000) And &HFF: hdr(25) = (imgH \ &H1000000) And &HFF
hdr(26) = 1: hdr(27) = 0   ' biPlanes = 1
hdr(28) = 1: hdr(29) = 0   ' biBitCount = 1
hdr(30) = 0: hdr(31) = 0: hdr(32) = 0: hdr(33) = 0  ' biCompression = BI_RGB
hdr(34) = dataSize And &HFF:          hdr(35) = (dataSize \ &H100) And &HFF
hdr(36) = (dataSize \ &H10000) And &HFF: hdr(37) = (dataSize \ &H1000000) And &HFF
hdr(38) = 0: hdr(39) = 0: hdr(40) = 0: hdr(41) = 0  ' XPelsPerMeter
hdr(42) = 0: hdr(43) = 0: hdr(44) = 0: hdr(45) = 0  ' YPelsPerMeter
hdr(46) = 2: hdr(47) = 0: hdr(48) = 0: hdr(49) = 0  ' biClrUsed = 2
hdr(50) = 2: hdr(51) = 0: hdr(52) = 0: hdr(53) = 0  ' biClrImportant = 2
' Color table: index 0 = black, index 1 = white
hdr(54) = 0:   hdr(55) = 0:   hdr(56) = 0:   hdr(57) = 0
hdr(58) = 255: hdr(59) = 255: hdr(60) = 255: hdr(61) = 0
Dim bmpPath As String
bmpPath = Environ("TEMP") & "\qr_" & Replace(Application.Caller.Address, "$", "") & ".bmp"
Dim bmpFile As Integer: bmpFile = FreeFile
Open bmpPath For Binary As #bmpFile
Put #bmpFile, , hdr
Dim bitMask(7) As Byte
bitMask(0) = 128: bitMask(1) = 64: bitMask(2) = 32: bitMask(3) = 16
bitMask(4) = 8:   bitMask(5) = 4:  bitMask(6) = 2:  bitMask(7) = 1
Dim rowBuf() As Byte: ReDim rowBuf(bmpStride - 1)
Dim px As Long, py As Long, qx As Long, qy As Long, ri As Long
For py = imgH - 1 To 0 Step -1  ' BMP stores rows bottom-up
    For ri = 0 To bmpStride - 1: rowBuf(ri) = 0: Next ri
    For px = 0 To imgW - 1
        qx = (px \ cellPx) - margin
        qy = (py \ cellPx) - margin
        If qx >= 0 And qx < s And qy >= 0 And qy < s Then
            If getPattern(qx, qy, m, version) = 0 Then  ' white module
                rowBuf(px \ 8) = rowBuf(px \ 8) Or bitMask(px Mod 8)
            End If
        Else  ' quiet zone = white
            rowBuf(px \ 8) = rowBuf(px \ 8) Or bitMask(px Mod 8)
        End If
    Next px
    Put #bmpFile, , rowBuf
Next py
Close #bmpFile
x = Application.Caller.MergeArea.Width
y = Application.Caller.MergeArea.Height
If x > y Then x = y
x = x * s / (s + 2)
Dim shpNew As Shape
Set shpNew = Application.Caller.Parent.Shapes.AddPicture( _
    bmpPath, False, True, _
    Application.Caller.Left + (Application.Caller.MergeArea.Width - x) / 2, _
    Application.Caller.Top + (Application.Caller.MergeArea.Height - x) / 2, _
    x, x)
' Kill bmpPath
With shpNew
     .Name = Application.Caller.Address
     .AlternativeText = text & "|QuickResponse barcode, level " & Mid("LMQH", lev + 1, 1) & _
        ", version " & IIf(version < 1, "M" & (version + 4), version) & _
        ", mode " & Array("digit", "alpha", "binary", "kanji")(mode) & _
        ", " & s & "x" & s & " cells"
     .LockAspectRatio = True
     .Placement = xlMove
End With

' MsgBox "수행 시간 : " & Format(Timer - tStart, "0.000") & " 초", vbInformation, "QRCodeImage"
 Exit Function
 failed:
 If Err.Number Then QRCodeImage = "ERROR QRCodeImage [" & Err.Number & "] " & Err.Description & " | bmpPath=" & bmpPath

End Function

' get QR pattern mask
Private Function getPattern(ByVal x As Long, ByVal y As Long, ByVal m As Integer, ByVal version As Integer) As Integer
Dim i As Integer, j As Long
If version < 1 Then m = Array(1, 4, 6, 7)(m) ' mask pattern of micro QR
i = mat(x, y)
If i < 2 Then
    Select Case m
    Case 0: j = (x + y) And 1
    Case 1: j = y And 1
    Case 2: j = x Mod 3
    Case 3: j = (x + y) Mod 3
    Case 4: j = (x \ 3 + y \ 2) And 1
    Case 5: j = ((x * y) And 1) + (x * y) Mod 3
    Case 6: j = (x * y + (x * y) Mod 3) And 1
    Case 7: j = (x + y + (x * y) Mod 3) And 1
    End Select
    If j = 0 Then i = i Xor 1 ' invert only data according mask
End If
getPattern = i And 1
End Function

Private Function utf16to8(text As String) As String
Dim i As Integer, c As Long
utf16to8 = text
For i = Len(text) To 1 Step -1
    c = AscW(Mid(text, i, 1)) And 65535
    If c > 127 Then
        If c > 4095 Then
            utf16to8 = Left(utf16to8, i - 1) + ChrW(224 + c \ 4096) + ChrW(128 + (c \ 64 And 63)) + ChrW(128 + (c And 63)) & Mid(utf16to8, i + 1)
        Else
            utf16to8 = Left(utf16to8, i - 1) + ChrW(192 + c \ 64) + ChrW(128 + (c And 63)) & Mid(utf16to8, i + 1)
        End If
    End If
Next i
End Function
