Attribute VB_Name = "PGM_KTD"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    PGM_KTD
'
'         update: 2002.6.29  s.f   difftime
'         update: 2002.10.5  s.f   difftime!
'         update: 2002.12.03 s.f   RecDtsave0, RecDtsave ?ǉﾁ
'         update: 2002.12.07 s.f   RecDtsave0(icnt) ?֕ύX
'         update: 2002.12.09 s.f   cooloff, heatoff ???凬潟Z?b?g?@?ǉﾁ
'         update: 2004. 3. 8 s.f   RecEmgDtsave ?・竡~???b?Z?[?W?̕ۑ?  2004.3.8'
'         update: 2004. 3.12 s.f   ???x?w?ߓd???@Global ?錾
'         update: 2004. 3.30 s.f   ?・竡~ү???ރo?O?C??
'         update: 2004. 5. 5 s.f   ???x?W???A?・␳???[?`???@?ǉﾁ  PGM_KTD,My_lib,MYEDIT, LS21_SC, LS21_TC
'         update: 2005. 9.27 s.f   ?ۉ??竡~???[?h?@?ǉﾁ
'         update: 2005. 9.28 s.f   T?W???@?\???F?ύX
'         update: 2005.11. 6 s.f   ?I?[?o?[?t???[?΍・idc65536,idc256,ddc05
'         update: 2006.04.14 s.f   on error goto
'         update: 2006.04.15 s.f   error ?\??
'         update: 2006.05.15 s.f   data???????݁A?ǂݍ??ݎ??@?hL"?@?ǉﾁ
'       Ver.3.33R_070927 2007.09.27 s.f  Z?␳?@?w?肵????ﾞﾒﾝﾄNo.?ց@?ł??驍・??ɂ??・
'       Ver.3.33R_071113 2007.11.13 s.f  ?u?????\?[?N?v?????A?@?u1?񐬌`?venable=False?ﾖ
'       Ver.3.33R_071119 2007.11.19 s.f  ?H????Ԑ??艨@?o?O?C???iedit???A?f?[?^?p???j?A???ϒlAND?ŐV?l?Ł@?X?V???閧ﾖ
'?@?@?@?@?@?@?@?@?@?@?@?@?@?@?@?@?@?@?@?@?H????Ԑ??艨@ON???́A?@T?W???f?[?^?́@?t?@?C?????轤ﾌ?ǂݍ??݂??Ȃ?
'       Ver.3.33R_071120 2007.11.20 s.f  ?o?O?C???A?@?󐬌`-?r?o?@?ǉAA?@?A?????`?ĊJ?@?ǉﾁ
'       Ver.3.33R_071203 2007.12.03 s.f  ?・竡~?@???b?Z?[?W?ǉAA?ύX
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'       Ver.NQD_70_080320 2008. 03.20 s.f  C870v1?・ｯ　ReadyWait?@???ׂĂɓ・・・
'       Ver.NQD_71_080910 2008. 9.10 s.f  ?VQD?@?Q???@???繧ﾉ?????@???[?^?[?p???X?A?]?綷・?????
'       Ver.NQD_71_090817 s.f  NQD6???繧ﾉ??????????
'       Ver.NQD_71_140111 s.f  TBK&TE????
'///////////////////////////////////////////////////////
'?@?@?@TBK&TE?@?????@?@?@Keyword=TBK/TE?@?@?@1?ӏ?
'///////////////////////////////////////////////////////
'
Option Explicit
'
Global iFlg_hijyou%                    ' ?・竡~FLG?@'090817?@NQD70_SC???辷ﾚ?ﾐ
'                                        0: ???・1:2ѱ?ﾌﾟ    2:?????・     3:4:5:DC?G???[
'                                        6: r_pres?@1ton?z??  7:    8:     9:
'                                       10:?@???ݻ?
'
Global InitDat!(0 To 50)               '?ۑ??f?[?^
Global InitStr$(0 To 50)
'
Global TPass!(0 To 12000)                '?o?ߎ??ﾔ(?b)
Global ZAxis!(0 To 12000)                '?W?iZ-???j
Global Press!(0 To 12000)                '?^?ｳ
Global Templ!(0 To 12000)                '?^???x
Global Templu!(0 To 12000)               '?^???x ?・
Global Templd!(0 To 12000)               '?^???x ??
Global Const ResDtSize = 12000
Global BrdFlg$
Global StartTime!                       'Debug?p
Global GCnt0%                           '???`???f?[?^?J?E???^
Global GCnt1%
Global Const H24Hr = 24# * 3600!        'Timer?p ?齠・̕b??
Global EmgFlg%                          '?・竡~
Global gOrgFlg%                         '???_???A????=TRUE
'
Global Err_ic%, Err_id%                 'ERROR
Global pv_ch!        '/* ?}?j???A?????̑??x?^?ʒu?؂芷???l*/
Global VccLw!                           '?^?ero
Global VccHi!                           '?^?󓞒B?_
'
Global FrmMenuFlg%                      '???j???[???甲???驍ﾆ??false
Global gM_sa!                           '???j???[?̑??x???艪ﾌ???ﾌ/* ?ݒ舳?ﾍ */
Global gM_p!                           '???j???[?̑??x???艪ﾌ???ﾌ/* ?ݒ閧o?l */
Global gM_lim!                           '???j???[?̑??x???艪ﾌ???ﾌ/* ?ݒ胊?~?b?g?l */
Global ViewFlg%                         '?譁ﾊ?ԍ?
'/////////////////////////////////////
'      TBK/TE
' /TBK/'  ////////
'Global Const gDirect! = -1            'S.M?̉]?綷・(+1 or -1)'08.3.24 tsubaki
'Global Const gRev2Disp As Double = 8000  '1?]?????閧ﾌ?p???X?? KD2002ED(ﾂﾊ޷)???[?^?[  2000*4=8000
' /TE/'  //////////
Global Const gDirect! = 1            'S.M?̉]?綷・(+1 or -1)  '08.09.10???h
Global Const gRev2Disp As Double = 24000   '1?]?????閧ﾌ?p???X?? ???h?T?[?{???[?^?[?@???]???o 24000Pulse/rev.  ?̂ﾌLS?@
'//////////////////////////////////////////////////////////////////
Global gTimeUpCnt%                      '?^?C???A?b?v?̃J?E???^
Global gVumFlg%                         '?^?󓞒B=1
Global Const idc16777216 As Long = 16777216  '?@?I?[?o?[?t???[?΍ﾅ?ǉﾁ  2005.11.22
Global Const idc8388607 As Long = 8388607  '?@?I?[?o?[?t???[?΍ﾅ?ǉﾁ  2005.11.22
Global Const idc65536  As Long = 65536  '?@?I?[?o?[?t???[?΍ﾅ?ǉA@???̉??R?s?@2005.11.6?@???D??
Global Const idc256 As Long = 256
Global Const ddc05 As Double = 0.5
Global Const dc0 As Double = 0#
Global Const LongTime As Long = 80000     '80000?b?i1?・86400?b?j?@?ُ墲ﾉ???????ԂƂ??驕B?@100116
Global sdt1$, sdt2$, sdt3$             '?G???[?\???p  2006.4.14
' -----------------  2005.5.?ǉﾁ
Global versionNo$, ppos$      '?@version?@No.?@PGM_Menu?́@Label1(13), PresentPosition(???݈ʒu?j
Global CmndColoff!(0 To 9)  '?R?}???h?t?̐F off???i?̐F
Global CmndColon!(0 To 9)  '?R?}???h?t?̐F?@on?@?????ꂽ?Ƃ??̐F
Global T_keisuCol!(0 To 4)  '???x?W???A?・␳?\???ﾌbackColor
Global kkno%               ' ?H????Ԑ??艨@?^No
'--------------- ?????Y???`?@?v???O????
' 2001?N3??
'
Global gcoxFlName$       '?R???g???[???f?[?^?t?@?C????
Global gcoxFldir$        '?f?B???N?g??
'
Global gCoxFlDtMax%
Global gCoxDlDt(0 To 200) As String       'cox?t?@?C???̓ǂ񂾂܂܂̃f?[?^
Global scom(0 To 200) As String       '
Global sisub(0 To 200) As Long        '
Global sjsub(0 To 200) As Long        '
Global sksub(0 To 200) As Long        '
Global slsub(0 To 200) As Long        '
Global hcomm(0 To 3) As String        '
Global dcomm(0 To 200) As String      '
Global seg_num(0 To 100) As Integer   '?Z?O?????g?ԍ?
Global ic(0 To 100) As Integer        '???苺・?
Global pres(0 To 100) As Single      '?v???X???ﾍ
Global z(0 To 100) As Single          '?ڕW?ʒu
Global vel(0 To 100) As Single        '???x
Global t0(0 To 100) As Single         'Time Out
Global p(0 To 100) As Single          'PID P
Global ptime%                         '???莞?ﾔ ??
Global ytemp%                         '?\?ﾁ?M ?x
'
' ----------   ???x?W???A?・␳?f?[?^     2004.4.30

Global ishu%                      ' ???`?@?H?・ﾚ
Global ishu_bkup%                 ' ???`?@?H?・ﾚ?ﾌbackup
Global T_keisuCont%(0 To 3)        '???x?W???R???g???[??
Global T_keisu!(0 To 9)          '???x?W???f?[?^
Global T_keisu_dum!
Global Z3_HoseiCont%(0 To 2)          '???x?W???R???g???[??
Global Z3_Hosei!(0 To 9)          '???x?W???f?[?^
Global DkatJ!(0 To 1)           '?H????ԖڕW?l
Global AkatJ!(0 To 1)           '?H????ԃA???[???ݒ閨@?繻ﾀ???ﾀ
Global Acp!(0 To 1)             'Cp?ʒu?A???[???ݒ閨@?繻ﾀ???ﾀ
Global Almdisp$                     '?A???[???\???@?H????Ԃ????b???ʒu
Global AlmON%                   ' alarm On/Off flg
Global iSeikeiTorF_flg%         ' ???`?@?L?򏂒???e?k?f
Global katCflag%                '?H????ﾔ ???????艫t???O
Global katDflag%                '?H????ԁ@???艨@?ύX?_???e?k?f
Global Henkou_No%               '?H????Ԏ??????范p ?^?ύX?燉e?@?ύX?Ȃ??A???炷?A???₷?A?・・ւ?
Global kaatsuJ!(0 To 10, 0 To 5)
Global iflgKataTorF%(0 To 9)     ' ?^?t???O?@true=?{???@?@False=dammy
Global iPltMax%                  '1?񐬌`???̌^?i?p???b?g?j??
Global Saikaiflg As Boolean     ' ?ĊJ?t???O2007.11.19 tsuika
Global lSokuFlg%                  '?????\?[?N?^?C??
Global Karauchiflg As Boolean   ' ?󐬌`?|?r?o?@?t???O
'
Global gDate$                         '???ʃO???t?喆t??
Global gTime$                         '???ʃO???t???ﾔ
Global gGphDtNum%                     '???ʃO???t?f?[?^??
Global gResFlName$                    '???ʃf?[?^?t?@?C????
Global gResFldir$                     '?f?B???N?g??
Global FlNmRecDt$                     '???`?f?[?^?t?@?C????
Global Rec_of_Mold$                   '???`?f?[?^?@?????ϐ?
'
Global gErrMsg$(0 To 1, 0 To 36)      '?G???[???Z?[?W
Global gemgmsg                        '?G???[???Z?[?W
'
Global kataNo$(0 To 10)                 ' ?^?̃i???o?[?@?@?@'2007.11.12?@tsuika
Global kataNoHyj$(0 To 36)                    ' ?^?m???D?@?\???p?????O?o?b?t?@
Global kataNoPnt As Integer                     '?@?^No.?@?|?C???^?[
Global katamax                          ' ???`?@?烽ﾌ?@?X?e?[?V??????
'--------------- [QD61]LS21_S.C ?Œ闍`???Ă??髟ﾏ??
Global KeikaTime%(0 To 12000)
Global atemp!(0 To 12000, 0 To 2)    '1801 ??12000?@?֕ύX?@130425
Global aposi!(0 To 12000)   '1801 ??12000?@?֕ύX?@130425
Global apre!(0 To 12000)   '1801 ??12000?@?֕ύX?@130425
Global roz!(2)               '?@?˓????`ﾊﾟﾗﾒ?@??,???ﾔ
Global ivd%, id_0%, id_1%, id_2%
'--------------- ?蓮?̈ʒu???䑬?x?ݒ阯p
Global gHiSpeed!                      '?蓮?̈ʒu???䑬?x
Global gLwSpeed!                      '?蓮?̈ʒu???䑬?x
Global r_pres_kousei!                  '???͓ǂݎ謔・l?̃[???Z??

Global nout%, v!                      'cal_pid?́@?ϐ??@???x?w?ߓd??

Global gOrgIL As Boolean              '???_?C???^?[???b?N
Global gOrgStartFlg As Boolean        '???񌴓_???A?????t???O
Global wTm0!, wTm1!                   '?o?ߎ??Ԍv?Z?p     2004.5.12 ?ǉﾁ  "??ﾊﾞ?ﾌﾛ?" ?΍・
'  -----        2009.8.17 ?ǉﾁ
Global fintime!, Tm2f1!, Tm2f2!
Public Sub Main()
Dim i%
' On Error GoTo errHandler:
  CmndColoff(1) = &H8000000F     '?I???R?}???h?t?̐F?@?@?@?@?@?D
  CmndColoff(3) = &HC0FFC0       'V?G?f?B?g?̃R?}???h?t?̐F?@?@???ﾎ
  CmndColoff(9) = &HC0C0FF       '?ۉ??竡~?̃R?}???h?t?̐F?@?@?s???N
  CmndColoff(0) = &HFFFFC0       '5???竡~?̃R?}???h?t?̐F?@?@???F
  CmndColon(1) = vbRed '&HFF&    '?R?}???h?t?̐F on?̂Ƃ??@?ԁ@?@?@?@?ﾔ
  CmndColon(3) = &HC0FFC0         '?R?}???h?t?̐F on?̂Ƃ??@???ﾎ
  CmndColon(9) = &HC0C0FF         '?R?}???h?t?̐F on?̂Ƃ??@?s???N
  CmndColon(0) = vbBlue        '?R?}???h?t?̐F on?̂Ƃ??@ao
  T_keisuCol!(0) = &HFFFFFF    '???x?W???A?・␳?@?\??backcolor?@off ?D?F
  T_keisuCol!(1) = &HFFFFC0    '???x?W???A?・␳?@?\??backcolor?@on?@???F
  T_keisuCol!(2) = &H800012    '???x?W???A?・␳?@?\??forecolor?@on?@??
  T_keisuCol!(3) = &HFF00FF    '???x?W???A?・␳?@?\??forecolor?@on point???@?@?s???N
  T_keisuCol!(4) = &HE0E0E0    '???x?W???A?・␳?@?\??backcolor?@dummy  ???F
  lSokuFlg = False        '?????\?[?N?^?C??   ?ʏ펞?́@OFF
  katCflag = False      ' ?v???O?????J?n???́A?K???H????膂FF
  Karauchiflg = False      ' ?v???O?????J?n???́A?齟Ufalse
  Saikaiflg = False         '?v???O?????J?n???́A?齟Ufalse
  katamax = 6           ' STATION SSU = 6
  ishu = 1              ' 1?T?ڂ??轣@?X?^?[?g

'
    For i = 0 To 9
        kataNo(i) = Format(i + 1, "##")     ' ?^?m???D?̏??匀ｻ
    Next i
    kataNo(10) = " 0"    ' ?^?m???D???????̏??0
'    ﾀﾞа?^?w?閧ﾌ?@reset
  For i = 0 To 9
    iflgKataTorF(i) = True
  Next i
'
ppos = "KTD"
'
  InitDtLoad
  cfileLoad
  coxDtRead gcoxFldir & gcoxFlName
  coxDtSet
  BoardInit
  ResetOFF          '/* ???Z?b?g?@?????? */
  SetErrMsg         '?A???[?????b?Z?[?W
  'DebugData         'Debug
  gResFlName = "*.mpr"                  '???ʃf?[?^?t?@?C????
  gResFldir = App.path & "\..\data\"  '?f?B???N?g??
  'ADMain.Show
  InitStr(2) = "roz.con"                    '???{?b?g?f?[?^?t?@?C????
  InitStr(3) = App.path & "\..\robo\"       '?f?B???N?g??
  'IOChk.Show '
  ViewFlg = 1
  gOrgFlg = False                       '???_???A????=TRUE
  gTimeUpCnt = 0                    '?^?C???A?b?v?̃J?E???^
  gVumFlg = 0                       '?^?󓞒B=1
  
'  VacuumOFF                        '2006.1221 ?폜?@s.f.
  SeikeiOFF                         '2006.12.21 ?V?K s.f
'
  CoolOFF
  HeatOFF
'
  ReadyFrm.Show
  'PGM_Menu.Show
  Exit Sub
'
''
End Sub

Public Sub coxFlLoad()
Dim fDir$, fname$, rflg%
    
    fname = gcoxFlName        '?R???g???[???f?[?^?t?@?C????
    fDir = gcoxFldir          '?f?B???N?g??
    rflg = False
    Call GenFile.SetCtrl("?t?@?C???Ǎ?", "?Ǎ?", "?謠ﾁ")
    Call GenFile.SetFile(cLoad, fDir, fname, "*.cox")
    GenFile.Show vbModal
    Call GenFile.GetFile(rflg, fDir, fname)
    Set GenFile = Nothing
    If rflg Then
      Screen.MousePointer = 11
      '
      coxDtRead fDir$ & fname
      gcoxFlName = fname      '?R???g???[???f?[?^?t?@?C????
      gcoxFldir = fDir        '?f?B???N?g??
      '
      Screen.MousePointer = 0
    End If
End Sub

Public Sub coxDtRead(fl$)
'?@080312:?@NQD?Ή?
Dim i%, fnum%, l%
Dim dmy$, dt$, com$, dta$(0 To 4)
Dim iaf%, ja%
Dim isub As Long
Dim jsub As Long
Dim ksub As Long
Dim lsub As Long

  fnum = FreeFile
  Open fl For Input As #fnum
    For l = 0 To 7
      Line Input #fnum, gCoxDlDt(l)
    Next l
    '
    For l = 0 To 2: hcomm(l) = gCoxDlDt(l): Next l
    l = 4: ptime = Val(gCoxDlDt(l))      '???莞?ﾔ
    l = 6: ytemp = Val(gCoxDlDt(l))      '?\?ﾁ?M???x
    l = 7
    '???쓮???艫R?}???h?̓Ǎ?
    For i = 0 To 100
      Line Input #fnum, dt
      l = l + 1
      gCoxDlDt(l) = dt
      seg_num(i) = Val(Mid(dt, 1, 2))
      ic(i) = Val(Mid(dt, 4, 4))
      z(i) = Val(Mid(dt, 9, 9))
      vel(i) = Val(Mid(dt, 19, 10))
      pres(i) = Val(Mid(dt, 30, 8))
      t0(i) = Val(Mid(dt, 39, 8))
      p(i) = Val(Mid(dt, 48, 6))
      If ic(i) = 9 Then Exit For
    Next i
    '?f?[?^?ﾇ?ݎ謔・
    Input #fnum, dmy
    l = l + 1
    gCoxDlDt(l) = dmy
    ja = 0
    For i = 0 To 200
      Line Input #fnum, dt
      l = l + 1
      gCoxDlDt(l) = dt
      scom(i) = Mid(dt, 1, 2)
      isub = Val(Mid(dt, 4, 5))
      com = Left(scom(i), 1)
      Select Case com
      Case "S", "L"                  ' 2006.5.15 "L" ?ǉﾁ s.f
        iaf = iaf + 1
        jsub = Val(Mid(dt, 10, 5))
        ksub = Val(Mid(dt, 16, 5))
        lsub = Val(Mid(dt, 22, 5))
      Case "J"
        iaf = iaf + 1
      Case "H"                      ' 2007.11.12 "H" ?ǉﾁ s.f
        iaf = iaf + 1
      Case "P"
        ja = ja + 1
        If Right(scom(i), 1) = "R" And isub = 1 And ic(ja - 1) <> 2 Then iaf = iaf + 1
        If Right(scom(i), 1) = "W" And isub = 4 And ic(ja - 1) <> 2 Then iaf = iaf + 1
      Case "E"
        Exit For
      End Select
      sisub(i) = isub
      sjsub(i) = jsub
      sksub(i) = ksub
      slsub(i) = lsub
    Next i
'  -- ???x?W???A?・␳?f?[?^
    Input #fnum, dmy
    l = l + 1
    gCoxDlDt(l) = dmy
    Input #fnum, T_keisuCont(0), T_keisuCont(1)
    l = l + 1
    gCoxDlDt(l) = "  " & Format(T_keisuCont(0), "0.000") & ",  " & Format(T_keisuCont(1), "0.000")
    If (katCflag = True) Then               ' katcflag?i?H????艫t???O?j??True?Ȃ轣@T?W???f?[?^?t?@?C?????逑ﾇ?ݍ??܂Ȃ?
            Line Input #fnum, dmy           ' ?P?s?ǂݔﾎ??
        Else
            Input #fnum, T_keisu(0), T_keisu(1), T_keisu(2), T_keisu(3), T_keisu(4)
    End If
    l = l + 1
    dt = "  " & Format(T_keisu(0), "0.000")
    For i = 1 To 4
      dt = dt & ",  " & Format(T_keisu(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
    If (katCflag = True) Then
            Line Input #fnum, dmy                 ' katcflag?i?H????艫t???O?j??True?Ȃ轣@T?W???f?[?^?t?@?C?????逑ﾇ?ݍ??܂Ȃ?
        Else
            Input #fnum, T_keisu(5), T_keisu(6), T_keisu(7), T_keisu(8), T_keisu(9)
    End If
    l = l + 1
    dt = "  " & Format(T_keisu(5), "0.000")
    For i = 6 To 9
      dt = dt & ",  " & Format(T_keisu(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
 '
    Input #fnum, dmy
    l = l + 1
    gCoxDlDt(l) = dmy
    Input #fnum, Z3_HoseiCont(0), Z3_HoseiCont(1), Z3_HoseiCont(2)
    l = l + 1
    gCoxDlDt(l) = "  " & Format(Z3_HoseiCont(0), "0.000") & ",  " & Format(Z3_HoseiCont(1), "0.000") & ",  " & Format(Z3_HoseiCont(2), "0.000")
    Input #fnum, Z3_Hosei(0), Z3_Hosei(1), Z3_Hosei(2), Z3_Hosei(3), Z3_Hosei(4)
    l = l + 1
    dt = "  " & Format(Z3_Hosei(0), "0.000")
    For i = 1 To 4
      dt = dt & ",  " & Format(Z3_Hosei(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
    Input #fnum, Z3_Hosei(5), Z3_Hosei(6), Z3_Hosei(7), Z3_Hosei(8), Z3_Hosei(9)
    l = l + 1
    dt = "  " & Format(Z3_Hosei(5), "0.000")
    For i = 6 To 9
      dt = dt & ",  " & Format(Z3_Hosei(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
 '
    Input #fnum, dmy                  '  ?H????Ԑ??䂍?????A?@???????̓ǂݍ??ﾝ
    l = l + 1
    gCoxDlDt(l) = dmy
    Input #fnum, DkatJ(1), DkatJ(0)
    l = l + 1
    gCoxDlDt(l) = "  " & Format(DkatJ(1), "000.0") & ",  " & Format(DkatJ(0), "000.0")
'
    Input #fnum, dmy                  '  ?^No.?@?f?[?^?@?ǂݍ??ﾝ
    l = l + 1
    gCoxDlDt(l) = dmy
    Input #fnum, kataNo(0), kataNo(1), kataNo(2), kataNo(3), kataNo(4), kataNo(5), kataNo(6), kataNo(7), kataNo(8)
    l = l + 1
    dt = "  " & kataNo(0)
    For i = 1 To 8
        dt = dt + "  " & kataNo(i)
    Next i
    gCoxDlDt(l) = dt
 '
    Input #fnum, dmy                  '  ?H????ﾔ&Cp?l?@ALARM ???????A?@???????̓ǂݍ??ﾝ
    l = l + 1
    gCoxDlDt(l) = dmy
'
    Input #fnum, AkatJ(1), AkatJ(0)
    l = l + 1
    gCoxDlDt(l) = "  " & Format(AkatJ(1), "0") & ",  " & Format(AkatJ(0), "0")
'
    Input #fnum, Acp(1), Acp(0)
    l = l + 1
    gCoxDlDt(l) = "  " & Format(Acp(1), "0.000") & ",  " & Format(Acp(0), "0.000")
'
'
  Close fnum
  gCoxFlDtMax = l
  gGphDtMax = iaf       '?f?[?^?? ???ﾍiaf
End Sub

Public Sub InitDtLoad()
Dim i%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
  flNm = "PGM.ini"
  Open fDir & flNm For Input As #fnum
  For i = 0 To 50
    Input #fnum, InitDat(i), InitStr(i)
  Next i
  Close #fnum
  'gcoxFlName = InitStr(0)       '?R???g???[???f?[?^?t?@?C????
  'gcoxFldir = InitStr(1)        '?f?B???N?g??
  'InitDat(10)=???`?J?E???^
  'InitDat(11)=???`?J?E???^?g?E?^??
End Sub
Public Sub InitDtSave()
Dim i%, fnum%
Dim fDir$, flNm$
  InitStr(0) = gcoxFlName    '?R???g???[???f?[?^?t?@?C????
  InitStr(1) = gcoxFldir     '?f?B???N?g??
  fnum = FreeFile
  fDir = App.path & "\..\Data\"
  flNm = "PGM.ini"
  Open fDir & flNm For Output As #fnum
  For i = 0 To 50
    Write #fnum, InitDat(i), InitStr(i)
  Next i
  Close #fnum
End Sub
Public Sub RecDtSave0(icnt!)                     '???`?f?[?^?t?@?C???̍쐬
Dim j%, fnum%, sdt$
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
'  FlNmRecDt = "LS" & Mid(Date, 6, 2) & Mid(Date, 9, 2) & Mid(Time, 1, 2) & Mid(Time, 4, 2) & Format(Int(icnt), "0") & ".lsl"
  FlNmRecDt = "LS" & Format$(Now, "yymmddhhmmss") & Format(Int(icnt), "0") & ".lsl"
  sdt = " No.     Z3         ct1    ct2"
  sdt = sdt & "      cc1     cc2    cc3"
  sdt = sdt & "    cc3-2     cp         8ﾄ     T?W??    Z3?␳"
  Open fDir & FlNmRecDt For Output As #fnum
     Write #fnum, gcoxFlName & "   " & Date$ & "   " & Time$
     Write #fnum, sdt
  Close #fnum
End Sub
Public Sub RecDtSave999()            '???`?f?[?^?̃Z?[?u?̏I???????@?@?R???g???[???f?[?^?ﾇ?A@?@2009.9.12?ǉﾁ
Dim j%, fnum%, l%
Dim fDir$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
  Open fDir & FlNmRecDt For Append As #fnum
    For l = 0 To gCoxFlDtMax
     Write #fnum, gCoxDlDt(l)
    Next l
  Close #fnum
End Sub
Public Sub RecDtSave(Rec_of_Mold$)            '???`?f?[?^?̃Z?[?u
Dim j%, fnum%
Dim fDir$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
  Open fDir & FlNmRecDt For Append As #fnum
     Write #fnum, Rec_of_Mold & "   " & Time$
  Close #fnum
End Sub
Public Sub RecEmgDtSave(sdt1$, sdt2$, sdt3$)            '?・竡~?f?[?^?̃Z?[?u  2004.3.8 ?ǉA@s.f
Dim j%, fnum%
Dim fDir$, emgmsg$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
  flNm = "emgmsg.txt"
     emgmsg = ArmEmgMsgChk$()
  Open fDir & flNm For Append As #fnum
     Write #fnum, Date$ & " " & Time$ & "  " & emgmsg$
     Write #fnum, "  " & sdt1
     Write #fnum, "  " & sdt2
     Write #fnum, "  " & sdt3 & ppos
  Close #fnum
End Sub
Public Sub ResDtSave(i_s%, i%)
Dim j%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\data\"
  flNm = Format$(Now, "yymmddhhmmss") & Trim(Str(i_s)) & "d.mpr"
  Open fDir & flNm For Output As #fnum
  Write #fnum, Date, gcoxFlName
  Write #fnum, Time
  Write #fnum, i
  For j = 0 To i
    Write #fnum, Format(Int(KeikaTime(j) / 60), "  0??") & Format(Int(KeikaTime(j)) Mod 60, " 0?b"), atemp(j, 0), atemp(j, 1), atemp(j, 2), apre(j), aposi(j)
  Next j
  Close #fnum
End Sub
Public Sub ResDtLoad(fDir$, flNm$)
Dim j%, fnum%, i%
  fnum = FreeFile
  Open fDir & flNm For Input As #fnum
  Input #fnum, gDate
  Input #fnum, gTime
  Input #fnum, gGphDtNum
  i = gGphDtNum
  For j = 0 To i
    Input #fnum, atemp(j, 0), atemp(j, 1), atemp(j, 2), apre(j), aposi(j)
  Next j
  Close #fnum
End Sub
Public Sub ResFlLoad()
Dim fDir$, fname$, rflg%
    
    fname = gResFlName        '???ʃf?[?^?t?@?C????
    fDir = gResFldir          '?f?B???N?g??
    rflg = False
    Call GenFile.SetCtrl("?t?@?C???Ǎ?", "?Ǎ?", "?謠ﾁ")
    Call GenFile.SetFile(cLoad, fDir, fname, "*.mpr")
    GenFile.Show vbModal
    Call GenFile.GetFile(rflg, fDir, fname)
    Set GenFile = Nothing
    If rflg Then
      Screen.MousePointer = 11
      '
      ResDtLoad fDir, fname
      gResFlName = fname      '?R???g???[???f?[?^?t?@?C????
      gResFldir = fDir        '?f?B???N?g??
      '
      Screen.MousePointer = 0
    End If
End Sub
' ---------------------------------------------------------
Public Sub coxDtSet()       '?@cox ?f?[?^?́@?ۑ??p?u1???C???̕????f?[?^?v?ϊ?
Dim i%, fnum%, l%
Dim dmy$, dt$, com$
Dim iaf%, ja%
Dim isub As Long
Dim jsub As Long
Dim ksub As Long
Dim lsub As Long

    For l = 0 To 2: gCoxDlDt(l) = hcomm(l): Next l
    l = 4: gCoxDlDt(l) = ptime    '???莞?ﾔ
    l = 6: gCoxDlDt(l) = ytemp    '?\?ﾁ?M???x
    l = 7
    '???쓮???艫R?}???h?̓Ǎ?
    For i = 0 To 100
      l = l + 1
      dt = gCoxDlDt(l)
      Mid(dt, 1, 2) = Right("  " & Str(seg_num(i)), 2)
      Mid(dt, 4, 4) = Right("    " & Str(ic(i)), 4)
      Mid(dt, 9, 9) = Right("         " & Format(z(i), "0.000"), 9)
      Mid(dt, 19, 10) = Right("        " & Format(vel(i), "0.00"), 10)
      Mid(dt, 30, 8) = Right("      " & Str(pres(i)), 8)
      Mid(dt, 39, 8) = Right("      " & Format(t0(i), "0.0"), 8)
      Mid(dt, 48, 6) = Right("      " & Format(p(i), "0.0"), 6)
      '
      gCoxDlDt(l) = dt
      If ic(i) = 9 Then Exit For
    Next i
    '?f?[?^?ﾇ?ݎ謔・
    l = l + 1
    '
    ja = 0
    For i = 0 To 200
      isub = sisub(i)
      jsub = sjsub(i)
      ksub = sksub(i)
      lsub = slsub(i)
      l = l + 1
      dt = gCoxDlDt(l)
      scom(i) = Mid(dt, 1, 2)
      Mid(dt, 4, 5) = Right("     " & Format(isub, "0"), 5)
      com = Left(scom(i), 1)
      Select Case com
      Case "S", "L"                    ' 2006.5.15 "L" ?ǉﾁ s.f
        Mid(dt, 10, 5) = Right("     " & Format(jsub, "0"), 5)
        Mid(dt, 16, 5) = Right("     " & Format(ksub, "0"), 5)
        Mid(dt, 22, 5) = Right("     " & Format(lsub, "0"), 5)
      Case "J"

      Case "H"
      
      Case "P"

      Case "E"
        Exit For
      End Select

      gCoxDlDt(l) = dt
    Next i
'  -- ???x?W???A?・␳?f?[?^
    l = l + 1   ' ?R?????g?s
    l = l + 1
    gCoxDlDt(l) = "  " & Format(T_keisuCont(0), "0.000") & ",  " & Format(T_keisuCont(1), "0.000")
    l = l + 1
    dt = "  " & Format(T_keisu(0), "0.000")
    For i = 1 To 4
      dt = dt & ",  " & Format(T_keisu(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
    l = l + 1
    dt = "  " & Format(T_keisu(5), "0.000")
    For i = 6 To 9
      dt = dt & ",  " & Format(T_keisu(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
 '
    l = l + 1  '  ?R?????g?s
    l = l + 1
    gCoxDlDt(l) = "  " & Format(Z3_HoseiCont(0), "0.000") & ",  " & Format(Z3_HoseiCont(1), "0.000") & ",  " & Format(Z3_HoseiCont(2), "0.000")
    l = l + 1
    dt = "  " & Format(Z3_Hosei(0), "0.000")
    For i = 1 To 4
      dt = dt & ",  " & Format(Z3_Hosei(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
    l = l + 1
    dt = "  " & Format(Z3_Hosei(5), "0.000")
    For i = 6 To 9
      dt = dt & ",  " & Format(Z3_Hosei(i), "0.000")
    Next i
    gCoxDlDt(l) = dt
'
'
  '  ?H????Ԑ??䂍?????A?@???????̏??????ﾝ
    l = l + 1   ' ?R?????g?s
    l = l + 1
    dt = "  " & Format(DkatJ(1), "000.0") & ",  " & Format(DkatJ(0), "000.0")
    gCoxDlDt(l) = dt
'
  '  ?^No.?@?f?[?^?@?̏??????ﾝ
    l = l + 1   ' ?R?????g?s
    l = l + 1
    dt = "  " & kataNo(0)
    For i = 1 To 8
        dt = dt + ",  " & kataNo(i)
    Next i
    gCoxDlDt(l) = dt
'
'  ?H????ԁ@???@?b???l?@?`?????????@???????A?@???????̏??????ﾝ
    l = l + 1   ' ?R?????g?s
    l = l + 1
    dt = "  " & Format(AkatJ(1), "0") & ",  " & Format(AkatJ(0), "0")
    gCoxDlDt(l) = dt
'
    l = l + 1
    dt = "  " & Format(Acp(1), "0.000") & ",  " & Format(Acp(0), "0.000")
    gCoxDlDt(l) = dt
''
  Close fnum
End Sub
Public Sub coxDtSave(fl$)
Dim l%, fnum%
  fnum = FreeFile
  Open fl For Output As #fnum
    For l = 0 To gCoxFlDtMax
      Print #fnum, gCoxDlDt(l)
    Next l
  Close #fnum
End Sub

Private Sub DebugData()
Dim i%
Dim z!, p!, t!, x!
'
  For i = 0 To ResDtSize
    TPass(i) = i                '?o?ߎ??ﾔ(?b)
    ZAxis(i) = 50 + 40 * Sin(i / 57.325)              '?W?iZ-???j
    Press(i) = i / 2000              '?^?ｳ
    Templ(i) = 500 + 100 * Sin(i / 57.325)       '?^???x
  Next i
End Sub

Public Sub BoardInit()
Dim flg%
    flg = 1
    Select Case flg
    Case 0
        BrdFlg = "OFF"
    Case 1
        BrdFlg = "ON"
        '--------------- D/A Board
        DeviceDaName
        'DvcDaOpen
        '--------------- A/D Board
        DvcAdOpen
        DeviceAdName
        '--------------- DIO Board
        DvcDioOpen
        '--------------- C-870V1
        Ready_Wait
        C870Open
    End Select
End Sub
Public Sub BoardClose()
Dim flg%
    flg = 1
    Select Case flg
    Case 0
        BrdFlg = "OFF"
    Case 1
        BrdFlg = "ON"
        '--------------- D/A Board
        'DeviceDaName
        DvcDaClose
        '--------------- A/D Board
        DvcAdClose
        'DeviceAdName
        '--------------- DIO Board
        'DvcDioClose
        '--------------- C-870V1
        Ready_Wait
        C870Close
    End Select
End Sub

Public Sub rozFileLoad()
Dim i%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = InitStr(3)
  flNm = InitStr(2)
  Open fDir & flNm For Input As #fnum
    Input #fnum, pv_ch                  '?ʒu?E???x???[?h?؊??_
    Input #fnum, roz(0), roz(1)         '?˓????`ﾊﾟﾗﾒ?@???A???ﾔ (???ﾔmax180?j
    Input #fnum, VccLw, VccHi           '?s???j?Q?[?W?p
    Input #fnum, gM_sa, gM_p, gM_lim    '???x???艪ﾌ?p?????[?^
    Input #fnum, gHiSpeed, gLwSpeed     '?蓮?̈ʒu???䑬?x
    Input #fnum, r_pres_kousei          '???͓ǎ謦l?@?O?Z??
  Close #fnum
'gM_sa!     '???j???[?̑??x???艪ﾌ???ﾌ/* ?ݒ舳?ﾍ */
'gM_p!      '???j???[?̑??x???艪ﾌ???ﾌ/* ?ݒ閧o?l */
'gM_lim!    '???j???[?̑??x???艪ﾌ???ﾌ/* ?ݒ胊?~?b?g?l */
End Sub
Public Sub rozFileSave()
Dim i%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = InitStr(3)
  flNm = InitStr(2)
  Open fDir & flNm For Output As #fnum
    Write #fnum, pv_ch
    Write #fnum, roz(0), roz(1)        '?˓????`ﾊﾟﾗﾒ?@???A???ﾔ
    Write #fnum, VccLw, VccHi
    Write #fnum, gM_sa, gM_p, gM_lim
    Write #fnum, gHiSpeed, gLwSpeed    '?蓮?̈ʒu???䑬?x
    Write #fnum, r_pres_kousei          '???͓ǎ謦l?@?O?Z??
  Close #fnum
End Sub
Public Sub ExecMemo(DDir$, flNm$)
Dim ExecFl$, fl$
Dim r!
  fl = DDir$ & flNm
  ExecFl = "C:\WINDOWS\NOTEPAD.EXE " & fl
'-------- ???????ﾅfl?J??
  r = Shell(ExecFl, 1)
  AppActivate r, True     '???????????驍ﾜ?ő҂ﾂ
End Sub
Public Function diffTime!(wTm1!, wTm0!)  '  '02.6.29  abs ?O??   !?・・・10/4 sf
'Dim wTm0!, wTm1!
'-------------- ?o?@wTm1?i???݁j?|?@wTm0(?ߋ?) ?p???Ԃec?Ōv?Z
  If wTm0 > wTm1 Then
    diffTime = wTm1 + H24Hr - wTm0
  Else
    diffTime = wTm1 - wTm0
    'diffTime = Abs(wTm1 - wTm0)
  End If
End Function

Public Function BitBSet%(dl%, bit%)
'
  BitBSet = dl Or (2 ^ bit%)

End Function
Public Function BitBReSet%(dl%, bit%)
'
  BitBReSet = dl And (&HFFFF - 2 ^ bit)

End Function
Public Function BitBTest%(dl%, bit%)
Dim sts%
'
  sts = 0
  If dl And 2 ^ bit Then sts = 1  '&h1
  BitBTest = sts
End Function
Public Sub cfileLoad()
Dim i%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\cont\"
  flNm = "cfile.con"
  Open fDir & flNm For Input As #fnum
    Input #fnum, gcoxFlName       '?R???g???[???f?[?^?t?@?C????
    Input #fnum, gcoxFldir        '?f?B???N?g??
  Close #fnum
End Sub
Public Sub cfileSave()
Dim i%, fnum%
Dim fDir$, flNm$
  fnum = FreeFile
  fDir = App.path & "\..\cont\"
  flNm = "cfile.con"
  Open fDir & flNm For Output As #fnum
    Write #fnum, gcoxFlName       '?R???g???[???f?[?^?t?@?C????
    Write #fnum, gcoxFldir        '?f?B???N?g??
  Close #fnum
End Sub
Public Sub WaitSec(t As Single)
'?P?ﾊ ?b
Dim tm!, InTm!, NTm!
  tm = 0
  InTm = Timer
  Do
    NTm = Timer
    DoEvents
    If NTm >= InTm Then
      tm = NTm - InTm
    Else
      tm = H24Hr - InTm + NTm
    End If
    'If gDurPauseFlg <> 0 Then Exit Do
    If tm > t Then Exit Do
  Loop
End Sub

Public Sub SetErrMsg()
Dim ErrNo%, EmgArm%
  EmgArm = 0          '?・竡~
  ErrNo = 0: gErrMsg$(EmgArm, ErrNo) = "System not ready" '
  ErrNo = 1: gErrMsg$(EmgArm, ErrNo) = "?o?b???・竡~" '?G???[???Z?[?W
  ErrNo = 2: gErrMsg$(EmgArm, ErrNo) = "?{?́@?・竡~?r?v"
  ErrNo = 3: gErrMsg$(EmgArm, ErrNo) = "???苳ﾕ?@?・竡~?r?v" '?@?f08.3?@?\???燉e?・ﾖ
  ErrNo = 4: gErrMsg$(EmgArm, ErrNo) = "???・g?h?g?d???`?k?l?@???`??"
  ErrNo = 5: gErrMsg$(EmgArm, ErrNo) = "???`???繻^?q?[?^?[ALM"
  ErrNo = 6: gErrMsg$(EmgArm, ErrNo) = "?T?[?{???[?^?ُ・
  ErrNo = 7: gErrMsg$(EmgArm, ErrNo) = "?`?????o???ُ・
  ErrNo = 8: gErrMsg$(EmgArm, ErrNo) = "?y???W?????ُ・
  ErrNo = 9: gErrMsg$(EmgArm, ErrNo) = "???`???@?????墲`?k?l"
  ErrNo = 10: gErrMsg$(EmgArm, ErrNo) = "?????ﾉORG???؂ꂽ"
  ErrNo = 11: gErrMsg$(EmgArm, ErrNo) = "?\?ﾁ?M?@?g?e?d???`?k?l"
  ErrNo = 12: gErrMsg$(EmgArm, ErrNo) = "?\?ﾁ?M?@?????墲`?k?l"
  ErrNo = 13: gErrMsg$(EmgArm, ErrNo) = "?\?ﾁ?M?A?h?g?d???`?k?l"
  ErrNo = 14: gErrMsg$(EmgArm, ErrNo) = "?\?ﾁ?M?A?????墲`?k?l"
  ErrNo = 15: gErrMsg$(EmgArm, ErrNo) = "???`?????^?q?[?^?[ALM"
  EmgArm = 1          '?A???[??
  ErrNo = 0: gErrMsg$(EmgArm, ErrNo) = "?`?k?l?@?O?@?i???g?p?j" '?G???[???Z?[?W
  ErrNo = 1: gErrMsg$(EmgArm, ErrNo) = "?y???W???????B" '?G???[???Z?[?W
  ErrNo = 2: gErrMsg$(EmgArm, ErrNo) = "?e?[?u???????B"
  ErrNo = 3: gErrMsg$(EmgArm, ErrNo) = "?p???b?g?R?????B"
  ErrNo = 4: gErrMsg$(EmgArm, ErrNo) = "?p???b?g?S?????B"
  ErrNo = 5: gErrMsg$(EmgArm, ErrNo) = "?p???b?g?Q?????B"
  ErrNo = 6: gErrMsg$(EmgArm, ErrNo) = "?p???b?g?P?????B"
  ErrNo = 7: gErrMsg$(EmgArm, ErrNo) = "?`?k?l?V?i???g?p?j"
  ErrNo = 8: gErrMsg$(EmgArm, ErrNo) = "???`?????x?ُ・
  ErrNo = 9: gErrMsg$(EmgArm, ErrNo) = "?\?ﾁ?M?Q???x?ُ・
  ErrNo = 10: gErrMsg$(EmgArm, ErrNo) = "?\?ﾁ?M?P???x?ُ・
  ErrNo = 11: gErrMsg$(EmgArm, ErrNo) = "???站p?V?????_?[???????B"
  ErrNo = 12: gErrMsg$(EmgArm, ErrNo) = "?^?󖢓??B"
  ErrNo = 13: gErrMsg$(EmgArm, ErrNo) = "???站p?V?????_?[?㖢???B"
  ErrNo = 14: gErrMsg$(EmgArm, ErrNo) = "?\?ﾁ?M???ړ??????B"
  ErrNo = 15: gErrMsg$(EmgArm, ErrNo) = "?\?ﾁ?M?繹ﾚ???????B"
End Sub
Public Sub DispCenter(frmObj As Form)
  Dim dmy As Long

  If frmObj.WindowState <> 0 Then frmObj.WindowState = 0
  dmy = Screen.Width - frmObj.Width
  If 1 < dmy Then
    frmObj.Left = dmy \ 2
  Else
    frmObj.Left = 0
  End If
  dmy = Screen.Height - frmObj.Height
  If 1 < dmy Then
    frmObj.Top = dmy \ 2
  Else
    frmObj.Top = 0
  End If
End Sub
