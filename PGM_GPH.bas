Attribute VB_Name = "PGM_GPH"
Option Explicit
Global gGphDtMax%         '?f?[?^?? ???ﾍiaf

Public Sub MyEditGph(obj As Object)
Dim i%, j%, itp%, icp%, ijp%, ja%, iaf%
Dim aatemp%, btemp%, az%, bz%, inq%, apre%, bpre%, k%
Dim com$
Dim x!, y!, xs!, ys!
Dim x0!, x1!, y0!, y1!
Dim ChrW!, ChrH!
Dim xMax!, yZMax!, yTMax!, yPMax!
Dim xMin!, yZMin!, yTMin!, yPMin!
Dim xZero!, yZZero!, yTZero!, yPZero!
Dim cor0!, cor1!, cor2!, cor3!, cor4!, cor5!
  '
  cor0 = QBColor(0)
  cor1 = QBColor(1)
  cor2 = QBColor(2)
  cor3 = QBColor(3)
  cor4 = vbGreen    'QBColor(4)     '04.10.30
  cor5 = QBColor(5)
  iaf = gGphDtMax                       '?f?[?^?? ???ﾍiaf
  yPZero = 0: yTZero = 10: yZZero = 20
 
  obj.Scale (30, 70)-(330, 380)      ' ?譁ﾊ?̑傫??(X0,Y0)-(X1,Y1)
  obj.DrawWidth = 1
'------------------------- ?ڐ?
  For j = 1 To iaf + 1
    x0 = 30 + 300 * j / (iaf + 2): y0 = 70
    x1 = 30 + 300 * j / (iaf + 2): y1 = 380
    MYEdit.Line1(j - 1).Visible = True
    MYEdit.Line1(j - 1).x1 = x0
    MYEdit.Line1(j - 1).X2 = x1
  Next j
  For j = iaf + 2 To 25   '  21 -> 25 2012.11.24
    MYEdit.Line1(j - 1).Visible = False
  Next j
  ' ---------------------- ???艫R?}???h
  j = 0
  For i = 0 To 200
    Select Case Left(scom(i), 1)
    Case "S"
      'obj.ForeColor = cor4
      aatemp = sisub(i)
      j = j + 1
      itp = itp + 1
      com = "T" & Format(itp, "0")
      x0 = 30 + 300 * (j + 1) / (iaf + 2): y0 = 298 - aatemp / 5
      obj.CurrentX = x0: obj.CurrentY = y0: obj.Print com
    Case "T"
      'obj.ForeColor = cor4
      icp = icp + 1
      com = "C" & Format(icp, "0")
      x0 = 20 + 300 * (j + 1) / (iaf + 2): y0 = 312 - aatemp / 5
      obj.CurrentX = x0: obj.CurrentY = y0: obj.Print com
    Case "J"
      'obj.ForeColor = cor1
      j = j + 1
      ijp = ijp + 1
      com = "J" & Format(ijp, "0")
      x0 = 20 + 300 * (j + 1) / (iaf + 2): y0 = 298 - aatemp / 5
      obj.CurrentX = x0: obj.CurrentY = y0: obj.Print com
    Case "P"
      If Mid(scom(i), 2, 1) = "R" And sisub(i) = 1 Then
        ja = ja + 1
        If ic(ja - 1) = 2 Then GoTo Pend
        j = j + 1
        inq = inq + 1
        az = z(ja - 1)
        'obj.ForeColor = cor0
        com = "Z" & Format(inq, "0")
        x0 = 20 + 300 * (j + 1) / (iaf + 2): y0 = 220 - Int(0.75 * az)
        obj.CurrentX = x0: obj.CurrentY = y0: obj.Print com
        If ic(ja - 1) <> 1 Then GoTo Pend
        'obj.ForeColor = cor5
        apre = pres(ja - 1)
        com = "P" & Format(inq, "0")
        x0 = 20 + 300 * (j + 1) / (iaf + 2): y0 = 358 - apre / 10
        obj.CurrentX = x0: obj.CurrentY = y0: obj.Print com
      End If
      If Mid(scom(i), 2, 1) = "W" And sisub(i) = 4 Then
        ja = ja + 1
        If ic(ja - 1) = 2 Then GoTo Pend
        j = j + 1
        inq = inq + 1
        az = z(ja - 1)
        'obj.ForeColor = cor0
        com = "Z" & Format(inq, "0")
        x0 = 20 + 300 * (j + 1) / (iaf + 2): y0 = 220 - Int(0.75 * az)
        obj.CurrentX = x0: obj.CurrentY = y0: obj.Print com
        If ic(ja - 1) <> 1 Then GoTo Pend
        'obj.ForeColor = cor5
        apre = pres(ja - 1)
        com = "P" & Format(inq, "0")
        x0 = 20 + 300 * (j + 1) / (iaf + 2): y0 = 358 - apre / 10
        obj.CurrentX = x0: obj.CurrentY = y0: obj.Print com
      End If
Pend:
    Case "E"
      Exit For
    End Select
    If k <> j Then
      x0 = 30 + 300 * j / (iaf + 2): y0 = 310 - btemp / 5
      x1 = 30 + 300 * (j + 1) / (iaf + 2): y1 = 310 - aatemp / 5
      obj.Line (x0, y0)-(x1, y1), cor4
      btemp = aatemp
      '
      x0 = 30 + 300 * j / (iaf + 2): y0 = 220 - Int(0.75 * bz)
      x1 = 30 + 300 * (j + 1) / (iaf + 2): y1 = 220 - Int(0.75 * az)
      obj.Line (x0, y0)-(x1, y1), cor0
      bz = az
      If ic(ja) <> 2 Then az = z(ja)
      '
      x0 = 30 + 300 * j / (iaf + 2): y0 = 370 - bpre / 10
      x1 = 30 + 300 * (j + 1) / (iaf + 2): y1 = 370 - apre / 10
      obj.Line (x0, y0)-(x1, y1), cor5
      bpre = apre
      apre = pres(ja)
      k = j
    End If
  Next i
  '
  
End Sub

Public Sub MoniGraph(obj As Object, ifst%, ifin%)
'------------------------ ???`?????j?^?O???t?̏??匀ｻ
'
'?@?@?@?@update 2002.8.20  ?u?ʒu?v?@???F?ɕύX
'?@?@?@?@update 2004.10.30  ?u???x?v?@?΂ɕύX
'
Dim i%
Dim x!, y!, xs!, ysZ!, ysT!, ysP!
Dim x0!, x1!, y0!, y1!
Dim ChrW!, ChrH!
Dim xMax!, yZMax!, yTMax!, yPMax!
Dim xMin!, yZMin!, yTMin!, yPMin!
Dim xZero!, yZZero!, yTZero!, yPZero!
Dim cor0!, cor1!, cor2!, cor3!, cor4!
Dim xVmax%, xVmin%, yVmax%, yVmin%
  '
  cor0 = vbYellow
  cor1 = vbRed
  cor2 = vbGreen
  cor3 = &HC0C0FF     '?s???N  vbBlack
  cor4 = &HFFFFC0     '???F
  '
  If ifin = 0 Then obj.Cls
  '
  xVmin = 0: xVmax = 1000
  yVmin = 0: yVmax = 1000
  obj.Scale (xVmin, yVmax)-(xVmax, yVmin)   ' ?譁ﾊ?̑傫??(X0,Y0)-(X1,Y1)
  '
  yZMin = InitDat(1)  '?O???t?X?P?[???W (Min)
  yZMax = InitDat(2)  '?O???t?X?P?[???W (Max)
  ysZ = (yVmax - yVmin) / (yZMax - yZMin)
  '
  yPMin = InitDat(3)  '?O???t?X?P?[???^?ｳ (Min)
  yPMax = InitDat(4)  '?O???t?X?P?[???^?ｳ (Max)
  ysP = (yVmax - yVmin) / (yPMax - yPMin)
  '
  yTMin = InitDat(5)  '?O???t?X?P?[???^???x (Min)
  yTMax = InitDat(6)  '?O???t?X?P?[???^???x (Max)
  ysT = (yVmax - yVmin) / (yTMax - yTMin)
  '
  xMin = InitDat(7) * 60 '?O???t?X?P?[???o?ߎ??ﾔ (Min)
  xMax = InitDat(8) * 60 '?O???t?X?P?[???o?ߎ??ﾔ (Max)
  xs = (yVmax - yVmin) / (xMax - xMin)
  '
  '---------------- ?W
  i = ifst
  x0 = (TPass(i) - xMin) * xs + xMin        '?o?ߎ??ﾔ(?b)
  y0 = (ZAxis(i) - yZMin) * ysZ + yZMin     '?W?iZ-???j
  For i = ifst + 1 To ifin
    x1 = (TPass(i) - xMin) * xs + xMin        '?o?ߎ??ﾔ(?b)
    y1 = (ZAxis(i) - yZMin) * ysZ + yZMin     '?W?iZ-???j
    obj.Line (x0, y0)-(x1, y1), cor0
    x0 = x1: y0 = y1
  Next i
  '---------------- ?^?ｳ
  i = ifst
  x0 = (TPass(i) - xMin) * xs + xMin        '?o?ߎ??ﾔ(?b)
  y0 = (Press(i) - yPMin) * ysP + yPMin     '
  For i = ifst + 1 To ifin
    x1 = (TPass(i) - xMin) * xs + xMin        '?o?ߎ??ﾔ(?b)
    y1 = (Press(i) - yPMin) * ysP + yPMin     '
    obj.Line (x0, y0)-(x1, y1), cor1
    x0 = x1: y0 = y1
  Next i
  '---------------- ?^???x?i?????v?j
  i = ifst
  x0 = (TPass(i) - xMin) * xs + xMin        '?o?ߎ??ﾔ(?b)
  y0 = (Templ(i) - yTMin) * ysT + yTMin     '
  For i = ifst + 1 To ifin
    x1 = (TPass(i) - xMin) * xs + xMin        '?o?ߎ??ﾔ(?b)
    y1 = (Templ(i) - yTMin) * ysT + yTMin     '
    obj.Line (x0, y0)-(x1, y1), cor2
    x0 = x1: y0 = y1
  Next i
  '---------------- ?^???x?i?繻^?j
  i = ifst
  x0 = (TPass(i) - xMin) * xs + xMin        '?o?ߎ??ﾔ(?b)
  y0 = (Templu(i) - yTMin) * ysT + yTMin     '
  For i = ifst + 1 To ifin
    x1 = (TPass(i) - xMin) * xs + xMin        '?o?ߎ??ﾔ(?b)
    y1 = (Templu(i) - yTMin) * ysT + yTMin     '
    obj.Line (x0, y0)-(x1, y1), cor3
    x0 = x1: y0 = y1
  Next i
  '---------------- ?^???x?i???^?j
  i = ifst
  x0 = (TPass(i) - xMin) * xs + xMin        '?o?ߎ??ﾔ(?b)
  y0 = (Templd(i) - yTMin) * ysT + yTMin     '
  For i = ifst + 1 To ifin
    x1 = (TPass(i) - xMin) * xs + xMin        '?o?ߎ??ﾔ(?b)
    y1 = (Templd(i) - yTMin) * ysT + yTMin     '
    obj.Line (x0, y0)-(x1, y1), cor4
    x0 = x1: y0 = y1
  Next i
End Sub
