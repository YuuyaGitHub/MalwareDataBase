Msg = MsgBox("このプログラムは非常に危険です。" & vbCrLf & "実行しますか？",vbExclamation + vbOKCancel, "危険！")
If Msg = vbOK Then
       Msg = MsgBox("実行してPCに損害が出ても一切の責任を負いかねます。" & vbCrLf & "自己責任で実行してください。" & vbCrLf & "" & vbCrLf & "よろしいですか？",vbQuestion + vbYesNo, "本当に実行しますか？")
       If Msg = vbYes Then
              MsgBox "ウイルスをインストールします。",0,""
              MsgBox "0%完了",0,"インストール中"
              MsgBox "3%完了",0,"インストール中"
              MsgBox "7%完了",0,"インストール中"
              MsgBox "11%完了",0,"インストール中"
              MsgBox "23%完了",0,"インストール中"
              MsgBox "31%完了",0,"インストール中"
              MsgBox "45%完了",0,"インストール中"
              MsgBox "66%完了",0,"インストール中"
              MsgBox "76%完了",0,"インストール中"
              MsgBox "98%完了",0,"インストール中"
              MsgBox "100%完了",0,"インストール中"
              MsgBox "ウイルスは正常にインストールされました。",64,"完了"
              MsgBox "システムが破損している恐れがあります。",16,"システムエラー"
              MsgBox "すべてのプログラムが破損している可能性があります。",16,"システムエラー"
              MsgBox "個人用ファイルが破損している可能性があります。",16,"システムエラー"
              MsgBox "ネットワークドライバが破損している可能性があります。",16,"システムエラー"
              MsgBox "ハードディスクが破損している可能性があります。",16,"システムエラー"
              MsgBox "次のファイルが破損している可能性があります:" & vbCrLf & "" & vbCrLf & "C:\Windows\System32\logonui.exe" & vbCrLf & "C:\Windows\System32\winlogon.exe" & vbCrLf & "C:\Windows\System32\wininit.exe",16,"システムエラー"
              MsgBox "ハードドライブを認識できません。",18,"システムエラー"
              MsgBox "ファイルを修復しますか？", vbQuestion + vbYesNo, "確認"
              WScript.Sleep 5000
              MsgBox "修復に必要なファイルが見つかりませんでした。" & vbCrLf & "コンピューターをリカバリーすることをお勧めします。" & vbCrLf & "" & vbCrLf & """OK""をクリックしたあと、すぐにPCをリカバリーしてください。",16,"エラー"
              MsgBox "このプログラムは冗談です。",64,"よくできました。"
              End If
       ElseIf Msg = vbNo Then
              WScript.Quit
         End If
If Msg = vbCancel Then
       WScript.Quit
End If