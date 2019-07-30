Attribute VB_Name = "PGM_S_Main"
Global VccDt!(0 To 10)        '?^?v?̃f?[?^
Global VccCnt%                '?^?v?̃f?[?^?z?ﾌ?J?E???^
Global VccTm0!
Global VccTm1!

Public Sub hijyou()
'/* ?・竡~?V?[?P???T?ɂ????・*/
'outp(DIO_P,0x10);   /* ?・竡~?V?[?P???T?ֈ˗? */
  EmgFlg = True
  DioOut 5, 1     '?・竡~
  SeikeiOFF       '???`OFF?@?ҋ@??
  ServoOFF
  OrgOFF
'  VacuumOFF      '2006.12.21 ?폜 s.f
  HeatOFF
  CoolOFF
  FrmEmg.Show 1
End Sub

'Public Sub cal_pid(m_sa!, m_p!, m_lim!)
''  float  m_sa,     /* ?ݒ舳?ﾍ */
''         m_p,      /* ?ݒ閧o?l */
''         m_lim;    /* ?ݒ胊?~?b?g?l */
'Dim i%, nout%
'Dim pa!, per!
'  pa = r_pres()     '/* ???ﾍ */
'
'  If pa > 1000 Then '/* ?w?舳?́{?Q?O?O?j???ȏ繧ﾅ?・竡~ */
'    hijyou
'    Exit Sub
'  End If
'
''/* ?o?h?c???Z */
'
'  per = 5 * (m_sa - pa) * Abs(m_sa - pa) / (m_p * m_p)
'  If per > m_lim Then per = m_lim
'  If per < (-1 * m_lim) Then per = -1 * m_lim
'  inout = Int(40.95 * per) + &H800
'
'  'outp(ADPORT,(nout%256));
'  'outp(ADPORT+1,0x20|(nout/256));
'
'End Sub

Public Sub err_sign(ic%, id%)
Dim i%, ih%
'
' ic
' B0-B3 : ?z???G???[ ?` ?T?[?{?G???[
' B4    : ?z???G???[?㉺?؊?
' id
' ?G???[ ?P?`?P?Q
  Err_ic = ic: Err_id = id
  'frmerr_sign.Show
  'Call frmerr_sign.DispErr
  
End Sub


Public Sub ErrBitRd(er1%, er2%, er3%)
Dim ch%
Dim er%(0 To 32), hdt%
'
  er1 = 0: er2 = 0: er3 = 0
  For ch = 5 To 8
    DioInput ch, hdt
    If hdt = 1 Then er1 = BitBSet(er1, ch - 5)
  Next ch
  For ch = 17 To 24
    DioInput ch, hdt
    If hdt = 1 Then er2 = BitBSet(er2, ch - 17)
  Next ch
  For ch = 25 To 32
    DioInput ch, hdt
    If hdt = 1 Then er3 = BitBSet(er3, ch - 25)
  Next ch
  '
End Sub
Public Function BitRd%(ch%)
Dim hdt%, sts%
'
  sts = 0
  Select Case ch
  Case 0
    For ch = 1 To 8
      DioInput ch, hdt
      If hdt = 1 Then sts = BitBSet(sts, ch - 1)
    Next ch
  Case 1
    For ch = 9 To 16
      DioInput ch, hdt
      If hdt = 1 Then sts = BitBSet(sts, ch - 9)
    Next ch
  Case 2
    For ch = 17 To 24
      DioInput ch, hdt
      If hdt = 1 Then sts = BitBSet(sts, ch - 17)
    Next ch
  Case 3
    For ch = 25 To 32
      DioInput ch, hdt
      If hdt = 1 Then sts = BitBSet(sts, ch - 25)
    Next ch
  End Select
  '
  BitRd = sts
  '
End Function

Public Sub outp(ch%, hdt%)
Dim sts%
  Select Case ch
  Case 0
    For ch = 1 To 8
      sts = BitBTest(hdt, ch)
      DioOut ch, sts
    Next ch
  Case 1
    For ch = 9 To 16
      sts = BitBTest(hdt, ch - 9)
      DioOut ch, sts
    Next ch
  Case 2
    For ch = 17 To 24
      sts = BitBTest(hdt, ch - 17)
      DioOut ch, sts
    Next ch
  Case 3
    For ch = 25 To 32
      sts = BitBTest(hdt, ch - 25)
      DioOut ch, sts
    Next ch
  End Select

End Sub
'Public Sub LS21S_Monitor()
'Dim dt!(0 To 4), i%
'Dim flg As Long
'  VccTm0 = Timer
'  If Int(VccTm0 * 10) = Int(VccTm1 * 10) Then Exit Sub
'  VccTm1 = VccTm0
'  flg = 0
'  'AdRead dt(), flg
'  '------ ?V?[?P???T?Ő^?v?̓d???Ƃ`?m?c(2002.6.11)
'  'If dt(4) < VccHi Or gVumFlg = 1 Then
'  '  gVumFlg = 0                       '?^?󓞒B=1
'    VacuumON                           '?^?󓞒B?M?????M
'  'Else
'    VacuumOFF                           '?^?󖢓??B?M?????M
'  'End If
'  Exit Sub
'  '
'  If dt(4) <= 0.2 Then Exit Sub
'
'  VccCnt = VccCnt + 1
'  If VccCnt < 0 Then VccCnt = 0
'  If VccCnt > 3 Then VccCnt = 0
'  VccDt(VccCnt) = dt(4)
'  '---- ?^?v?@?k?????W?@?P?@?`?@?R
'  '  ?^?ERO?@?@?@?@?@?@?@?@?@?^?󓞒B?_
'  For i = 1 To 3
'    If VccLw < VccDt(i) And VccDt(i) < VccHi Then flg = flg + 1
'  Next i
'  If flg = 3 Or gVumFlg = 1 Then
'    gVumFlg = 0                       '?^?󓞒B=1
'    VacuumON                 '    '2006.12.21 ?폜 s.f
'    For i = 0 To 4: VccDt(i) = 10: Next i
'    WaitSec 1
'    VacuumOFF          '    '2006.12.21 ?폜 s.f
'  Else
'    VacuumOFF
'  End If
'End Sub


Public Function DispCtrlCommand$(i%)
Dim sdt$
    sdt = Right("     " & Format(i, "0"), 4)
    sdt = sdt & "  " & Right("     " & Format(seg_num(i), "0"), 4)   ' /* ?Z?O?????g?ԍ? */
    sdt = sdt & "  " & Right("     " & Format(ic(i), "0"), 4)        ' /* ???苺・? 1,2,3,8,9 */
    sdt = sdt & "  " & Right("         " & Format(z(i), "0.000"), 7) ' /* ?ڕW?ʒu */
    sdt = sdt & "  " & Right("         " & Format(vel(i), "0.0"), 7) ' /* ???x */
    sdt = sdt & "  " & Right("       " & Format(pres(i), "0"), 6)    ' /* ?v???X???ﾍ */
    sdt = sdt & "  " & Right("     " & Format(t0(i), "0.0"), 4)      ' /* ?^?C???A?E?g?l */
    sdt = sdt & "  " & Right("     " & Format(p(i), "0.0"), 4)       ' /* ?o?h?c?@?o */
  DispCtrlCommand = sdt
End Function
