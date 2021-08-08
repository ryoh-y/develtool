    #MakeGitkeep.vbs
    ' 作業ディレクトリはドラッグ・アンド・ドロップ
    Dim dir
    If WScript.Arguments.Count > 0 Then
        dir = WScript.Arguments.Item(0)
    End If
    If dir = "" Then
        msg = "パスがないです。ドラッグ・アンド・ドロップしてください"
    Else
        ' 一時バッファ用辞書
        Dim fs,tempDic
        Set tempDic = CreateObject("Scripting.Dictionary")
        Set fs = CreateObject("scripting.Filesystemobject")
        Call ChildPathget(dir)
        For Each key In tempDic
            Call Makegitkeep(tempDic.Item(key))
        Next
        Set fs = Nothing
        Set tempDic = Nothing
        msg = "完了"
    End If
     
    msgbox msg
     
    '//==================================================================================
    '// 関数  ：空フォルダのパスを辞書にいれる
    '//==================================================================================
    Private Sub ChildPathget(pdir)
        Dim subF,SubFolder
        Set SubFolder = fs.GetFolder(pdir).SubFolders
        Set wFiles = fs.GetFolder(pdir).Files
        If SubFolder.Count > 0 Then
            For Each subF In SubFolder
                ' 子供がいる時は再帰
                Call ChildPathget(subF.Path)
            Next
        Else
            ' 子供が居なくてファイルもない時は自分
            If wFiles.Count = 0 Then
                Call tempDic.Add(tempDic.Count,pdir)
            End If
        End If
        Set SubFolder = Nothing
        Set wFiles =Nothing
    End Sub
     
    '//==================================================================================
    '// 関数  ：gitkeep作成
    '//==================================================================================
    Private Sub Makegitkeep(wdir)
        Set obj = CreateObject("Scripting.FileSystemObject")
        Set mgit = obj.openTextFile(wdir &"/.gitkeep",8,True)
        Set obj = Nothing
        Set mgit = Nothing
    End Sub