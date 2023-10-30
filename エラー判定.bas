Attribute VB_Name = "エラー判定"
Option Explicit
Sub テスト()
    Dim 判定, 文
    '引数：年齢, 接種日, ワクチン名, 回数, 前回接種日, 前回年齢E文
    判定 = 接種判定(5, #2/1/2022#, "ファイザー（５から１１歳用）", 2, #1/31/2022#, "")
    文 = 判定(0) & vbCrLf & 判定(1)
    MsgBox 文
End Sub
Function 接種判定(年齢 As Variant, 接種日 As Variant, ワクチン名 As String, 回数 As Long, 前回接種日 As Variant, 前回年齢E文 As String) As Variant()
    Select Case ワクチン名
    '現行ワクチン
        Case "コミナティ（ＸＢＢ．１．５）": 接種判定 = XBBファイザー判定(年齢, 接種日, 回数, 前回接種日)
        Case "コミナティ５から１１歳用ＸＢＢ．１．５": 接種判定 = 小児XBBファイザー判定(年齢, 接種日, 回数, 前回接種日, 前回年齢E文)
        Case "コミナティ６か月から４歳用ＸＢＢ．１．５": 接種判定 = 乳幼児XBBファイザー判定(年齢, 接種日, 回数, 前回接種日, 前回年齢E文)
        Case "スパイクバックス（ＸＢＢ．１．５）": 接種判定 = XBBモデルナ判定(年齢, 接種日, 回数, 前回接種日)
        Case "スパイクバックス６～１１歳ＸＢＢ．１．５": 接種判定 = 小児XBBモデルナ判定(年齢, 接種日, 回数, 前回接種日, 前回年齢E文)
        Case "スパイクバックス６月～５歳ＸＢＢ．１．５": 接種判定 = 乳幼児XBBモデルナ判定(年齢, 接種日, 回数, 前回接種日, 前回年齢E文)
        Case "ノババックス": 接種判定 = ノババックス判定(年齢, 接種日, 回数, 前回接種日)
    '終了ワクチン
        Case "ファイザー": 接種判定 = ファイザー判定(年齢, 接種日, 回数, 前回接種日)
        Case "武田／モデルナ": 接種判定 = モデルナ判定(年齢, 接種日, 回数, 前回接種日)
        Case "アストラゼネカ": 接種判定 = アストラゼネカ判定(年齢, 接種日, 回数, 前回接種日)
        Case "ファイザー（５から１１歳用）": 接種判定 = 小児ファイザー判定(年齢, 接種日, 回数, 前回接種日, 前回年齢E文)
        Case "コミナティ（５から１１歳用ＢＡ．４／５）": 接種判定 = 小児BA5ファイザー判定(年齢, 接種日, 回数, 前回接種日, 前回年齢E文)
        Case "コミナティ（２価：ＢＡ．１）": 接種判定 = BA1ファイザー判定(年齢, 接種日, 回数, 前回接種日)
        Case "コミナティ（２価：ＢＡ．４／５）": 接種判定 = BA5ファイザー判定(年齢, 接種日, 回数, 前回接種日)
        Case "スパイクバックス（２価：ＢＡ．１）": 接種判定 = BA1モデルナ判定(年齢, 接種日, 回数, 前回接種日)
        Case "スパイクバックス（２価：ＢＡ．４／５）": 接種判定 = BA5モデルナ判定(年齢, 接種日, 回数, 前回接種日)
        Case "モデルナ（６から１１歳用ＢＡ．４／５）": 接種判定 = 小児BA5モデルナ判定(年齢, 接種日, 回数, 前回接種日)
        Case "コミナティ（６か月から４歳用）": 接種判定 = 乳幼児ファイザー判定(年齢, 接種日, 回数, 前回接種日, 前回年齢E文)
    End Select
End Function
Function 年齢判定(年齢 As Variant, 下限 As Long, 上限 As Long) As String
    If 年齢 = "" Then
        年齢判定 = "年齢不明"
        Exit Function
    End If
    If 年齢 < 下限 Then 年齢判定 = 下限 & "歳未満"
    If 上限 <> 0 Then
        If 年齢 > 上限 Then 年齢判定 = 上限 + 1 & "歳以上"
    End If
End Function
Function 間隔判定(接種日 As Variant, 前回接種日 As Variant, 設定値 As Long, 単位 As String) As String
    If 前回接種日 = "" Then
        間隔判定 = "前回不明"
        Exit Function
    End If
    Select Case 単位
        Case "日"
            If 接種日 - 前回接種日 < 設定値 Then
                間隔判定 = 設定値 & 単位 & "未満"
            End If
        Case "月"
            If DateAdd("m", -設定値, 接種日) < 前回接種日 Then
                間隔判定 = 設定値 & 単位 & "未満"
            End If
    End Select
End Function
Function XBBファイザー判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種開始
            Select Case 回数
                Case 1: XBBファイザー判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: XBBファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case 3, 4, 5, 6, 7: XBBファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: XBBファイザー判定 = Array("", "法定外回数")
            End Select
        Case Else: XBBファイザー判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function 小児XBBファイザー判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant, 前回年齢E文 As String) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種開始
            Select Case 回数
                Case 1: 小児XBBファイザー判定 = Array(年齢判定(年齢, 5, 11), "")
                Case 2:
                    If 前回年齢E文 = "" Then
                        小児XBBファイザー判定 = Array("", 間隔判定(接種日, 前回接種日, 19, "日"))
                        Else: 小児XBBファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 19, "日"))
                    End If
                Case 3, 4, 5, 6: 小児XBBファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: 小児XBBファイザー判定 = Array("", "法定外回数")
            End Select
        Case Else: 小児XBBファイザー判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function 乳幼児XBBファイザー判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant, 前回年齢E文 As String) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種開始
            Select Case 回数
                Case 1: 乳幼児XBBファイザー判定 = Array(年齢判定(年齢, 0, 4), "")
                Case 2
                    If 前回年齢E文 = "" Then
                        乳幼児XBBファイザー判定 = Array("", 間隔判定(接種日, 前回接種日, 19, "日"))
                        Else: 乳幼児XBBファイザー判定 = Array(年齢判定(年齢, 0, 4), 間隔判定(接種日, 前回接種日, 19, "日"))
                    End If
                Case 3
                    If 前回年齢E文 = "" Then
                        乳幼児XBBファイザー判定 = Array("", 間隔判定(接種日, 前回接種日, 56, "日"))
                        Else: 乳幼児XBBファイザー判定 = Array(年齢判定(年齢, 0, 4), 間隔判定(接種日, 前回接種日, 56, "日"))
                    End If
                Case 4: 乳幼児XBBファイザー判定 = Array(年齢判定(年齢, 0, 4), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: 乳幼児XBBファイザー判定 = Array("", "法定外回数")
            End Select
        Case Else: 乳幼児XBBファイザー判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function XBBモデルナ判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #11/1/2023# '1・2回目接種開始
            Select Case 回数
                Case 1: XBBモデルナ判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: XBBモデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3, 4, 5, 6, 7: XBBモデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: XBBモデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #9/25/2023# '接種開始
            Select Case 回数
                Case 3, 4, 5, 6, 7: XBBモデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: XBBモデルナ判定 = Array("", "法定外回数")
            End Select
        Case Else: XBBモデルナ判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function 小児XBBモデルナ判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant, 前回年齢E文 As String) As Variant()
    Select Case 接種日
        Case Is >= #11/1/2023# '1・2回目接種開始
            Select Case 回数
                Case 1: 小児XBBモデルナ判定 = Array(年齢判定(年齢, 6, 11), "")
                Case 2:
                    If 前回年齢E文 = "" Then
                        小児XBBモデルナ判定 = Array("", 間隔判定(接種日, 前回接種日, 21, "日"))
                        Else: 小児XBBモデルナ判定 = Array(年齢判定(年齢, 6, 11), 間隔判定(接種日, 前回接種日, 21, "日"))
                    End If
                Case 3, 4, 5, 6: 小児XBBモデルナ判定 = Array(年齢判定(年齢, 6, 11), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: 小児XBBモデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #9/25/2023# '接種開始
            Select Case 回数
                Case 3, 4, 5, 6: 小児XBBモデルナ判定 = Array(年齢判定(年齢, 6, 11), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: 小児XBBモデルナ判定 = Array("", "法定外回数")
            End Select
        Case Else: 小児XBBモデルナ判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function 乳幼児XBBモデルナ判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant, 前回年齢E文 As String) As Variant()
    Select Case 接種日
        Case Is >= #11/1/2023# '1・2回目接種開始
            Select Case 回数
                Case 1: 乳幼児XBBモデルナ判定 = Array(年齢判定(年齢, 0, 5), "")
                Case 2:
                    If 前回年齢E文 = "" Then
                        乳幼児XBBモデルナ判定 = Array("", 間隔判定(接種日, 前回接種日, 21, "日"))
                        Else: 乳幼児XBBモデルナ判定 = Array(年齢判定(年齢, 0, 5), 間隔判定(接種日, 前回接種日, 21, "日"))
                    End If
                Case Else: 乳幼児XBBモデルナ判定 = Array("", "法定外回数")
            End Select
        Case Else: 乳幼児XBBモデルナ判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function ノババックス判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '7回目接種開始
            Select Case 回数
                Case 1: ノババックス判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ノババックス判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3, 4, 5, 6, 7: ノババックス判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 6, "月"))
                Case Else: ノババックス判定 = Array("", "法定外回数")
            End Select
        Case Is >= #5/8/2023# '6回目接種開始
            Select Case 回数
                Case 1: ノババックス判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ノババックス判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3, 4, 5, 6: ノババックス判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 6, "月"))
                Case Else: ノババックス判定 = Array("", "法定外回数")
            End Select
        Case Is >= #3/8/2023# '3～5回目年齢引き下げ
            Select Case 回数
                Case 1: ノババックス判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ノババックス判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3, 4, 5: ノババックス判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 6, "月"))
                Case Else: ノババックス判定 = Array("", "法定外回数")
            End Select
        Case Is >= #11/8/2022# '4・5回目接種開始（R4秋開始接種枠に移行）
            Select Case 回数
                Case 1: ノババックス判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ノババックス判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3, 4, 5: ノババックス判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 6, "月"))
                Case Else: ノババックス判定 = Array("", "法定外回数")
            End Select
        Case Is >= #7/22/2022# '初回のみ年齢引き下げ
            Select Case 回数
                Case 1: ノババックス判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ノババックス判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3: ノババックス判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 6, "月"))
                Case Else: ノババックス判定 = Array("", "法定外回数")
            End Select
        Case Is >= #5/25/2022# '1～3回目接種開始
            Select Case 回数
                Case 1: ノババックス判定 = Array(年齢判定(年齢, 18, 0), "")
                Case 2: ノババックス判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3: ノババックス判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 6, "月"))
                Case Else: ノババックス判定 = Array("", "法定外回数")
            End Select
        Case Else: ノババックス判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function BA5ファイザー判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種終了
            BA5ファイザー判定 = Array("", "法定外ワクチン")
        Case Is >= #8/7/2023# '1・2回目接種開始
            Select Case 回数
                Case 1: BA5ファイザー判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: BA5ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case 3, 4, 5, 6: BA5ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA5ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #5/8/2023# '6回目接種開始
            Select Case 回数
                Case 3, 4, 5, 6: BA5ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA5ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #10/21/2022# '間隔短縮
            Select Case 回数
                Case 3, 4, 5: BA5ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA5ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #10/13/2022# '3～5回目接種開始
            Select Case 回数
                Case 3, 4, 5: BA5ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 5, "月"))
                Case Else: BA5ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Else: BA5ファイザー判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function BA1ファイザー判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種終了
            BA1ファイザー判定 = Array("", "法定外ワクチン")
        Case Is >= #8/7/2023# '1・2回目接種開始
            Select Case 回数
                Case 1: BA1ファイザー判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: BA1ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case 3, 4, 5, 6: BA1ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA1ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #5/8/2023# '6回目接種開始
            Select Case 回数
                Case 3, 4, 5, 6: BA1ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA1ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #10/21/2022# '間隔短縮
            Select Case 回数
                Case 3, 4, 5: BA1ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA1ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #9/20/2022# '3～5回目接種開始
            Select Case 回数
                Case 3, 4, 5: BA1ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 5, "月"))
                Case Else: BA1ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Else: BA1ファイザー判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function ファイザー判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種終了
            ファイザー判定 = Array("", "法定外ワクチン")
        Case Is >= #4/1/2023# '3・4回目使用終了
            Select Case 回数
                Case 1: ファイザー判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case Else: ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #10/21/2022# '3・4回目間隔短縮
            Select Case 回数
                Case 1: ファイザー判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case 3: ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case 4: ファイザー判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #5/25/2022# '4回目接種開始＆間隔短縮
            Select Case 回数
                Case 1: ファイザー判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case 3: ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 5, "月"))
                Case 4: ファイザー判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 5, "月"))
                Case Else: ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #3/25/2022# '3回目年齢引き下げ
            Select Case 回数
                Case 1: ファイザー判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case 3: ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 6, "月"))
                Case Else: ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #12/17/2021# '3回目間隔短縮
            Select Case 回数
                Case 1: ファイザー判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case 3: ファイザー判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 6, "月"))
                Case Else: ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #12/1/2021# '3回目接種開始
            Select Case 回数
                Case 1: ファイザー判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case 3: ファイザー判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 8, "月"))
                Case Else: ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #6/1/2021# '年齢引き下げ
            Select Case 回数
                Case 1: ファイザー判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: ファイザー判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case Else: ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #2/17/2021# '当初
            Select Case 回数
                Case 1: ファイザー判定 = Array(年齢判定(年齢, 16, 0), "")
                Case 2: ファイザー判定 = Array(年齢判定(年齢, 16, 0), 間隔判定(接種日, 前回接種日, 19, "日"))
                Case Else: ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Else: ファイザー判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function 小児BA5ファイザー判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant, 前回年齢E文 As String) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種終了
            小児BA5ファイザー判定 = Array("", "法定外ワクチン")
        Case Is >= #8/7/2023# '1・2回目接種開始
            Select Case 回数
                Case 1: 小児BA5ファイザー判定 = Array(年齢判定(年齢, 5, 11), "")
                Case 2:
                    If 前回年齢E文 = "" Then
                        小児BA5ファイザー判定 = Array("", 間隔判定(接種日, 前回接種日, 19, "日"))
                        Else: 小児BA5ファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 19, "日"))
                    End If
                Case 3, 4, 5: 小児BA5ファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: 小児BA5ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #5/8/2023# '5回目接種開始
            Select Case 回数
                Case 3, 4, 5: 小児BA5ファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: 小児BA5ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #3/8/2023# '3・4回目接種開始
            Select Case 回数
                Case 3, 4: 小児BA5ファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: 小児BA5ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Else: 小児BA5ファイザー判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function 小児ファイザー判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant, 前回年齢E文 As String) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種終了
            小児ファイザー判定 = Array("", "法定外ワクチン")
        Case Is >= #4/1/2023# '3回目使用終了
            Select Case 回数
                Case 1: 小児ファイザー判定 = Array(年齢判定(年齢, 5, 11), "")
                Case 2:
                    If 前回年齢E文 = "" Then
                        小児ファイザー判定 = Array("", 間隔判定(接種日, 前回接種日, 19, "日"))
                        Else: 小児ファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 19, "日"))
                    End If
                Case Else: 小児ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #3/8/2023# '3回目間隔短縮
            Select Case 回数
                Case 1: 小児ファイザー判定 = Array(年齢判定(年齢, 5, 11), "")
                Case 2:
                    If 前回年齢E文 = "" Then
                        小児ファイザー判定 = Array("", 間隔判定(接種日, 前回接種日, 19, "日"))
                        Else: 小児ファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 19, "日"))
                    End If
                Case 3: 小児ファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: 小児ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #9/6/2022# '3回目接種開始
            Select Case 回数
                Case 1: 小児ファイザー判定 = Array(年齢判定(年齢, 5, 11), "")
                Case 2:
                    If 前回年齢E文 = "" Then
                        小児ファイザー判定 = Array("", 間隔判定(接種日, 前回接種日, 19, "日"))
                        Else: 小児ファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 19, "日"))
                    End If
                Case 3: 小児ファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 5, "月"))
                Case Else: 小児ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Is >= #2/1/2022# '接種開始
            Select Case 回数
                Case 1: 小児ファイザー判定 = Array(年齢判定(年齢, 5, 11), "")
                Case 2:
                    If 前回年齢E文 = "" Then
                        小児ファイザー判定 = Array("", 間隔判定(接種日, 前回接種日, 19, "日"))
                        Else: 小児ファイザー判定 = Array(年齢判定(年齢, 5, 11), 間隔判定(接種日, 前回接種日, 19, "日"))
                    End If
                Case Else: 小児ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Else: 小児ファイザー判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function 乳幼児ファイザー判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant, 前回年齢E文 As String) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種終了
            乳幼児ファイザー判定 = Array("", "法定外ワクチン")
        Case Is >= #10/24/2022# '接種開始
            Select Case 回数
                Case 1: 乳幼児ファイザー判定 = Array(年齢判定(年齢, 0, 4), "")
                Case 2
                    If 前回年齢E文 = "" Then
                        乳幼児ファイザー判定 = Array("", 間隔判定(接種日, 前回接種日, 19, "日"))
                        Else: 乳幼児ファイザー判定 = Array(年齢判定(年齢, 0, 4), 間隔判定(接種日, 前回接種日, 19, "日"))
                    End If
                Case 3
                    If 前回年齢E文 = "" Then
                        乳幼児ファイザー判定 = Array("", 間隔判定(接種日, 前回接種日, 56, "日"))
                        Else: 乳幼児ファイザー判定 = Array(年齢判定(年齢, 0, 4), 間隔判定(接種日, 前回接種日, 56, "日"))
                    End If
                Case Else: 乳幼児ファイザー判定 = Array("", "法定外回数")
            End Select
        Case Else: 乳幼児ファイザー判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function BA5モデルナ判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種終了
            BA5モデルナ判定 = Array("", "法定外ワクチン")
        Case Is >= #5/8/2023# '6回目接種開始
            Select Case 回数
                Case 3, 4, 5, 6: BA5モデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA5モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #12/14/2022# '年齢引き下げ
            Select Case 回数
                Case 3, 4, 5: BA5モデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA5モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #11/28/2022# '3～5回目接種開始
            Select Case 回数
                Case 3, 4, 5: BA5モデルナ判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA5モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Else: BA5モデルナ判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function BA1モデルナ判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種終了
            BA1モデルナ判定 = Array("", "法定外ワクチン")
        Case Is >= #5/8/2023# '6回目接種開始
            Select Case 回数
                Case 3, 4, 5, 6: BA1モデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA1モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #12/14/2022# '年齢引き下げ
            Select Case 回数
                Case 3, 4, 5: BA1モデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA1モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #10/21/2022# '間隔短縮
            Select Case 回数
                Case 3, 4, 5: BA1モデルナ判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: BA1モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #9/20/2022# '3～5回目接種開始
            Select Case 回数
                Case 3, 4, 5: BA1モデルナ判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 5, "月"))
                Case Else: BA1モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Else: BA1モデルナ判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function 小児BA5モデルナ判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #9/20/2023# '接種終了
            小児BA5モデルナ判定 = Array("", "法定外ワクチン")
        Case Is >= #8/7/2023# '3～5回目接種開始
            Select Case 回数
                Case 3, 4, 5: 小児BA5モデルナ判定 = Array(年齢判定(年齢, 6, 11), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: 小児BA5モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Else: 小児BA5モデルナ判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function モデルナ判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #2/12/2023# '接種終了
            モデルナ判定 = Array("", "法定外ワクチン")
        Case Is >= #12/14/2022# '3回目年齢引き下げ
            Select Case 回数
                Case 1: モデルナ判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: モデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3: モデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case 4: モデルナ判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #10/21/2022# '3・4回目間隔短縮
            Select Case 回数
                Case 1: モデルナ判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: モデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3, 4: モデルナ判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 3, "月"))
                Case Else: モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #5/25/2022# '4回目接種開始＆間隔短縮
            Select Case 回数
                Case 1: モデルナ判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: モデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3, 4: モデルナ判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 5, "月"))
                Case Else: モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #12/17/2021# '3回目接種開始
            Select Case 回数
                Case 1: モデルナ判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: モデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case 3: モデルナ判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 6, "月"))
                Case Else: モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #8/2/2021# '年齢引き下げ
            Select Case 回数
                Case 1: モデルナ判定 = Array(年齢判定(年齢, 12, 0), "")
                Case 2: モデルナ判定 = Array(年齢判定(年齢, 12, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case Else: モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Is >= #5/22/2021#
            Select Case 回数
                Case 1: モデルナ判定 = Array(年齢判定(年齢, 18, 0), "")
                Case 2: モデルナ判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 21, "日"))
                Case Else: モデルナ判定 = Array("", "法定外回数")
            End Select
        Case Else: モデルナ判定 = Array("", "法定外ワクチン")
    End Select
End Function
Function アストラゼネカ判定(年齢 As Variant, 接種日 As Variant, 回数 As Long, 前回接種日 As Variant) As Variant()
    Select Case 接種日
        Case Is >= #10/13/2022# '接種終了
            アストラゼネカ判定 = Array("", "法定外ワクチン")
        Case Is >= #8/2/2021# '接種開始
            Select Case 回数
                Case 1: アストラゼネカ判定 = Array(年齢判定(年齢, 18, 0), "")
                Case 2: アストラゼネカ判定 = Array(年齢判定(年齢, 18, 0), 間隔判定(接種日, 前回接種日, 28, "日"))
                Case Else: アストラゼネカ判定 = Array("", "法定外回数")
            End Select
        Case Else: アストラゼネカ判定 = Array("", "法定外ワクチン")
    End Select
End Function
