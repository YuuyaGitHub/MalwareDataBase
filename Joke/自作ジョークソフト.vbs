Msg = MsgBox("���̃v���O�����͔��Ɋ댯�ł��B" & vbCrLf & "���s���܂����H",vbExclamation + vbOKCancel, "�댯�I")
If Msg = vbOK Then
       Msg = MsgBox("���s����PC�ɑ��Q���o�Ă���؂̐ӔC�𕉂����˂܂��B" & vbCrLf & "���ȐӔC�Ŏ��s���Ă��������B" & vbCrLf & "" & vbCrLf & "��낵���ł����H",vbQuestion + vbYesNo, "�{���Ɏ��s���܂����H")
       If Msg = vbYes Then
              MsgBox "�E�C���X���C���X�g�[�����܂��B",0,""
              MsgBox "0%����",0,"�C���X�g�[����"
              MsgBox "3%����",0,"�C���X�g�[����"
              MsgBox "7%����",0,"�C���X�g�[����"
              MsgBox "11%����",0,"�C���X�g�[����"
              MsgBox "23%����",0,"�C���X�g�[����"
              MsgBox "31%����",0,"�C���X�g�[����"
              MsgBox "45%����",0,"�C���X�g�[����"
              MsgBox "66%����",0,"�C���X�g�[����"
              MsgBox "76%����",0,"�C���X�g�[����"
              MsgBox "98%����",0,"�C���X�g�[����"
              MsgBox "100%����",0,"�C���X�g�[����"
              MsgBox "�E�C���X�͐���ɃC���X�g�[������܂����B",64,"����"
              MsgBox "�V�X�e�����j�����Ă��鋰�ꂪ����܂��B",16,"�V�X�e���G���["
              MsgBox "���ׂẴv���O�������j�����Ă���\��������܂��B",16,"�V�X�e���G���["
              MsgBox "�l�p�t�@�C�����j�����Ă���\��������܂��B",16,"�V�X�e���G���["
              MsgBox "�l�b�g���[�N�h���C�o���j�����Ă���\��������܂��B",16,"�V�X�e���G���["
              MsgBox "�n�[�h�f�B�X�N���j�����Ă���\��������܂��B",16,"�V�X�e���G���["
              MsgBox "���̃t�@�C�����j�����Ă���\��������܂�:" & vbCrLf & "" & vbCrLf & "C:\Windows\System32\logonui.exe" & vbCrLf & "C:\Windows\System32\winlogon.exe" & vbCrLf & "C:\Windows\System32\wininit.exe",16,"�V�X�e���G���["
              MsgBox "�n�[�h�h���C�u��F���ł��܂���B",18,"�V�X�e���G���["
              MsgBox "�t�@�C�����C�����܂����H", vbQuestion + vbYesNo, "�m�F"
              WScript.Sleep 5000
              MsgBox "�C���ɕK�v�ȃt�@�C����������܂���ł����B" & vbCrLf & "�R���s���[�^�[�����J�o���[���邱�Ƃ������߂��܂��B" & vbCrLf & "" & vbCrLf & """OK""���N���b�N�������ƁA������PC�����J�o���[���Ă��������B",16,"�G���["
              MsgBox "���̃v���O�����͏�k�ł��B",64,"�悭�ł��܂����B"
              End If
       ElseIf Msg = vbNo Then
              WScript.Quit
         End If
If Msg = vbCancel Then
       WScript.Quit
End If