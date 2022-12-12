#Persistent
#SingleInstance, Force
#NoEnv
#UseHook
#InstallKeybdHook
#InstallMouseHook
#HotkeyInterval, 2000
#MaxHotkeysPerInterval, 200
Process, Priority,, Realtime
SendMode, Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2

KOJI_NAME := "開建：岩内幹線用水路上清川工区工"
YYMM := "R4-04"

Return

;;;;;;ツールチップ関係;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
my_tooltip_function(str, delay) {
  ToolTip, %str%
  SetTimer, remove_tooltip, -%delay%
}

remove_tooltip:
  ToolTip
Return

remove_tooltip_all:
  SetTimer, remove_tooltip, Off
  Loop, 20
  ToolTip, , , , % A_Index
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; 改行コード除去 ;;;
rm_crlf(str) {
  str := RegExReplace(str, "\n", "")
  str := RegExReplace(str, "\r", "")
  Return str
}
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;修飾キー  + ➾ Shift		^ ➾ Control		! ➾ Alt		#  Windowsロゴキー
;
;　設定済みキー一覧
; 
;　vk1D(無変換)＋α
;　+^!F10:: 
;　+^!F11:: 
;　^+!d::
;　^+!g::
;　^vkBB:(:Ctrl + ; )
;　!vk1D
;　
;
;
;

; ^+!G::Run, https://www.google.com/　;Ctrl+Shift+Alt+GでGoogleを開く





vkF0::Return ;英数キー(CapsLock)無効

;;;;;;無変換キー　＋　α　;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
; ↑
vk1D & i::Send, {Blind}{Up}
vk1D & w::Send, {Blind}{Up}


; ↓
vk1D & k::Send, {Blind}{Down}
vk1D & s::Send, {Blind}{Down}

; ←
vk1D & j::Send, {Blind}{Left}
vk1D & a::Send, {Blind}{Left}

; →
vk1D & l::Send, {Blind}{Right}
vk1D & d::Send, {Blind}{Right}

; Home
vk1D & h::Send, {Blind}{Home}

; End
vk1D & vkBB::Send, {Blind}{End}

; ↑↑↑↑
vk1D & u::Send, {Blind}{Up 4}

; ↓↓↓↓
vk1D & ,::Send, {Blind}{Down 4}

; →→→→
vk1D & .::Send, {Blind}{Right 4}

; ←←←←
vk1D & m::Send, {Blind}{Left 4}

; Enter
vk1D & Space::Send, {Blind}{Enter}

; Backspace
vk1D & n::Send, {Blind}{Backspace}

; Delete
vk1D & /::Send,{Blind}{Delete}

; 行挿入(Ctrl押下状況により、挿入位置を前後で選択)
vk1D & Enter::
  If (GetKeyState("Ctrl", "P")) {
    Send, {Up}{End}{Enter}
  } Else {
    Send, {End}{Enter}
  }
Return

; 半角英数
vk1D & vkF2::Send, {vkF2}{vkF3}

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;;; 矢印入力 ;;;
vk1D & Up::
  Clipboard = ↑
  Send, ^v
Return

vk1D & Down::
  Clipboard = ↓
  Send, ^v
Return

vk1D & Left::
  Clipboard = ←
  Send, ^v
Return

vk1D & Right::
  Clipboard = →
  Send, ^v
Return
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; google検索 ;;;
^+!g::
    stash := ClipboardAll
    Clipboard :=
    Send, ^c
    ClipWait, 0.05
    clip := Clipboard
    Clipboard := stash
    clip := rm_crlf(clip)
    Run, https://www.google.co.jp/search?q=%clip%
    clip := ""
Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;; deepl検索 ;;;
^+!d::
    stash := ClipboardAll
    Clipboard :=
    Send, ^c
    ClipWait, 0.05
    clip := Clipboard
    Clipboard := stash
    clip := rm_crlf(clip)
    Run, https://www.deepl.com/translator#en/ja/%clip%
    clip := ""
  Return

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
F18:: 

    stash :=ClipboardAll
    Clipboard :=
    Send, ^c
    ClipWait, 0.5
    clip :=Clipboard
    Clipboard :=stash
  
    ;クリップボードの文字列を先頭10文字取り出す  
    StringLeft, clip, clip, 10
    Clipboard := clip
    ClipWait, 0.5
    
    ;CSVエディタのウインドウをアクティブにする     
    WinActivate, ahk_exe Cassava.exe
    sleep, 1000
    
    Send,{Blind}{F5}
    ;MsgBox, %clip% ; 
    sleep, 500
    Send,^f
    sleep, 300
    Send,^v
    sleep, 200
    Send,{Enter}
    sleep, 200
    Send,{Blind}{F3}
    sleep, 200
    Send,{Blind}{F3}
    sleep, 200
    Send,{Blind}{F3}
    sleep, 500
    my_tooltip_function("過去利用した類似の単価を確認して下さい", 2000)
      
Return

;;; 事前準備： 工事ツリーと「単語」リボンの右端を揃える
; GAIA画面最大起動時に内訳表の1行目の「？」をダブルクリック
F19::
  CoordMode, Mouse,Screen
  Click,307,520,2
Return

;;; 事前準備： 工事ツリーと「単語」リボンの右端を揃える
; GAIA画面最大起動時に「上の階層に移動」をクリック
F20:: 
   CoordMode, Mouse,Screen
;  Click,197,245 ; 5K用
  Click,158,212 ; 4K用
  sleep, 800
;  MouseMove,307,680 ; 5K用
  MouseMove,232,512 ; 4K用 上から３つ目を選ぶ
Return

;!vk1D:: Send, {Backspace} ;Alt + 無変換 を BackSpace　に置き換え


;;; 単価登録-見積参考資料 ;;;
F21::
    stash := ClipboardAll
    Clipboard :=
    Send, ^c
    ClipWait, 0.5
    clip := Clipboard
    Clipboard := stash
    FileAppend, %clip%`t, GAIA.txt
    clip := ""
    my_tooltip_function("単価名を登録しました", 500)
  Return


;;; 単価登録-GAIA文字列 ;;; コピーする前にキーボード処理が入る E　→　E
F22::
    Send, e
    sleep, 100
    Send, e
    sleep, 100
    stash := ClipboardAll
    Clipboard :=
    Send, ^c
    ClipWait, 0.5
    clip := Clipboard
    Clipboard := stash
    FileAppend, %KOJI_NAME%`t%YYMM%`t%clip%, GAIA.txt
    clip := ""
    my_tooltip_function("採用単価の詳細情報を登録しました", 500)
Return

;;; 単価登録-GAIA文字列 ;;;キャンセル
F23::
    text := "Sキャンセル"
    FileAppend, %text%`n, GAIA.txt
    my_tooltip_function("キャンセル", 500)
Return
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
F24:: 

    stash :=ClipboardAll
    Clipboard :=
    Send, ^c
    ClipWait, 0.5
    clip :=Clipboard
    Clipboard :=stash

    ;歩掛文字列変換   ATLUS用
    StringReplace, clip, clip, SP土砂等運搬  _ダンプトラック10ｔ積級,SP　土砂等運搬　
    StringReplace, clip, clip, 不整地運搬_不整地運搬車,不整地運搬　特装運搬

    ;不要文字のスペース置換   GAIA用
    StringReplace, clip, clip, 水道用硬質ポリ塩化ビニール管ＲＲ継手片受ベンド,ＲＲ　継手　片受　ベンド　
    StringReplace, clip, clip, 暗渠排水管 ポリエチレンパイプ波状管,ポリエチレンパイプ　波状管　
    StringReplace, clip, clip, ﾌﾗﾝｼﾞ接合材,ﾌﾗﾝｼﾞ　接合　
    StringReplace, clip, clip, フランジ接合材,ﾌﾗﾝｼﾞ　接合　
    StringReplace, clip, clip, 強さ, 　
    StringReplace, clip, clip, または, 　
    StringReplace, clip, clip, ＫＮ, 　
    StringReplace, clip, clip, kg用, 　
    StringReplace, clip, clip, Ｂ Ｎ合金製, 　
    StringReplace, clip, clip, φ, 　
    StringReplace, clip, clip, ｋｇｆ, 　
    StringReplace, clip, clip, ｃｍ２, 　
    StringReplace, clip, clip, ｍ２, 　
    StringReplace, clip, clip, T=, 　
    StringReplace, clip, clip, L=, 　
    StringReplace, clip, clip, 呼径, 　
    StringReplace, clip, clip, ｃｍ, 　×
    StringReplace, clip, clip, ×, 　
    StringReplace, clip, clip, ×, 　
    StringReplace, clip, clip, ×, 　
    StringReplace, clip, clip, モノフィラメント系,　
    StringReplace, clip, clip, ｍｍ, 　
    StringReplace, clip, clip, ダクタイルライニング直管接合材,ダクタイル 接合 水道用　
    StringReplace, clip, clip, ダクタイルライニング直管,ダクタイル ライニング　
    StringReplace, clip, clip, ダクタイル異形管,ダクタイル　異形管　
    StringReplace, clip, clip, Ｔ形, 　Ｔ　
    StringReplace, clip, clip, Ｋ形, 　Ｋ　
    StringReplace, clip, clip, 離脱防止付き,　
    StringReplace, clip, clip, ｋｇ用,　
    StringReplace, clip, clip, かんがい専用, かんがい　
    StringReplace, clip, clip, コンクリート空積割増, 空積　帯広　
    StringReplace, clip, clip,°,　
    StringReplace, clip, clip,゜,　    
    StringReplace, clip, clip,～,　
    StringReplace, clip, clip,～,　
    StringReplace, clip, clip, ｍ ,  
    StringReplace, clip, clip, ｍ　, 
    StringReplace, clip, clip, ｍ ,  
    StringReplace, clip, clip, ｍ　,  
    StringReplace, clip, clip, ｍ ,  
    StringReplace, clip, clip, ｍ　,  
    StringReplace, clip, clip, L=, 
    StringReplace, clip, clip, 種 , 
    StringReplace, clip, clip, 種　, 
    StringReplace, clip, clip, 型 , 
    StringReplace, clip, clip, 型　, 
    
     ;全角数字の半角化
    zenkakuS = １２３４５６７８９．０
    hankakuS = 123456789.0

    loop, parse, zenkakuS
        zenkakuS%a_index% = %a_loopfield%3
    loop, parse, hankakuS
        hankakuS%a_index% = %a_loopfield%
    loop, 11
    clip := RegExReplace(clip, zenkakuS%a_index%, hankakuS%a_index%)
    


    ;前後が数字の漢字の削除　
     FoundPos := RegExMatch(clip,"幅[0-9]")
      If ( FonndPos != 0 )
      {
        x :=FoundPos-1
        y :=FoundPos+1


        StringMid, text_1, clip, 1, x
        StringMid, text_2, clip, y,
        clip :=text_1 . text_2
      }
      
     FoundPos := RegExMatch(clip,"長[0-9]")
      If ( FonndPos != 0 )
      {
        x :=FoundPos-1
        y :=FoundPos+1


        StringMid, text_1, clip, 1, x
        StringMid, text_2, clip, y,
        clip :=text_1 . text_2
      }
  
      FoundPos := RegExMatch(clip,"高[0-9]")
      If ( FonndPos != 0 )
      {
        x := FoundPos-1        x := FoundPos-1

        y := FoundPos+1

        StringMid, text_1, clip, 1, x
        StringMid, text_2, clip, y,
        clip :=text_1 . text_2
      }
      
      ;MsgBox, %clip% ;
      ;MsgBox, %FoundPos% ; 
      Clipboard := clip 
      ClipWait, 0.5 
      Send, ^v 
      my_tooltip_function("文字列を加工しました", 500)
      
Return
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

#IfWinActive, ahk_class XLMAIN
F13::   ExcelMojiColor(255, 0, 0)        ; 赤
F14::   ExcelMojiColor(0, 0, 255)        ; 青
F15::                ; 自動(黒)
    Send, {alt}
    Send, h
    Send, fc
    Send, {Enter}
    return

; パラメータで指定された文字色に変更する
ExcelMojiColor(r, g, b)
{
    Send, {alt}
    Send, h
    Send, fc
    Send, m
    Send, {Right}    ; ユーザー設定を選択
    Send, !r

    Send, %r%       ; R
    Send, {tab}
    Send, %g%       ; G
    Send, {tab}
    Send, %b%       ; B
    sleep, 100      ; 別にいらないけど一瞬くらい表示させてやろう
    Send, {Enter}
}
#IfWinActive

#IfWinActive, ahk_class ahk_class WindowsForms10.Window.8.app.0.3ce0bb8_r7_ad1
F13::   
   CoordMode, Mouse,Screen
  
  Click,1131,94 ; 4K用
  sleep, 200
  Click,1151,125 ; 4K用
Return

F14::   
   CoordMode, Mouse,Screen
  
  Click,1131,94 ; 4K用
  sleep, 200
  Click,1151,125 ; 4K用
Return

#IfWinActive

