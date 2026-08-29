#Requires AutoHotkey v2.0
#SingleInstance Force

global word := ""
global doc := ""

MyGui := Gui("+Resize", "Word图片批量插入工具v1.2(@天tian)")
MyGui.Add("Text", "x20 y20 w100", "文档路径:")
EditDocPath := MyGui.Add("Edit", "x20 y40 w300")
BtnBrowseDoc := MyGui.Add("Button", "x330 y38 w60 h25", "浏览...")
BtnBrowseDoc.OnEvent("Click", BrowseDoc)
MyGui.Add("Text", "x20 y80 w100", "图片文件夹:")
EditPicFolder := MyGui.Add("Edit", "x20 y100 w300")
BtnBrowsePic := MyGui.Add("Button", "x330 y98 w60 h25", "浏览...")
BtnBrowsePic.OnEvent("Click", BrowsePic)
MyGui.Add("Text", "x20 y140 w100", "排序方式:")
SortDrop := MyGui.Add("DropDownList", "x20 y160 w150", ["升序", "降序"])
SortDrop.Value := 1
BtnStart := MyGui.Add("Button", "x20 y220 w80 h30", "开始处理")
BtnStart.OnEvent("Click", StartProcessing)
MyGui.OnEvent("Close", (*) => ExitApp())
MyGui.Show()

BrowseDoc(*) {
    global EditDocPath
    SelectedFile := FileSelect("3", , "请选择Word文档", "Word文档 (*.doc;*.docx)")
    if (SelectedFile != "") {
        EditDocPath.Text := SelectedFile
    }
}

BrowsePic(*) {
    global EditPicFolder
    SelectedDir := DirSelect(, "3", "请选择图片文件夹")
    if (SelectedDir != "") {
        EditPicFolder.Text := SelectedDir
    }
}

StartProcessing(*) {
    global EditDocPath, EditPicFolder, SortDrop, word, doc
    
    MyGui.Submit("NoHide")
    DocPath := EditDocPath.Text
    PicFolder := EditPicFolder.Text
    SortOrder := SortDrop.Value

    word := ""
    doc := ""

    if (DocPath = "" || PicFolder = "") {
        MsgBox("请先选择Word文档和图片文件夹！", "错误", "IconX")
        return
    }
    if !FileExist(DocPath) || !InStr(FileExist(PicFolder), "D") {
        MsgBox("路径不存在，请检查！", "错误", "IconX")
        return
    }

    ; ---------------- 【修正：使用标准的 ProcessClose】 ----------------
    try {
        ProcessClose("WINWORD.EXE") 
        Sleep 1000
    }
    ; ------------------------------------------------------------------

    fileList := ""
    Loop Files, PicFolder . "\*", "F" {
        ext := StrLower(A_LoopFileExt)
        if (ext != "jpg" && ext != "jpeg" && ext != "png" && ext != "bmp" && ext != "gif") {
            continue
        }
        fileName := A_LoopFileName
        if RegExMatch(fileName, "\d+", &match) {
            num := Integer(match[0])
        } else {
            num := 0
        }
        fileList .= num . "|" . A_LoopFileFullPath . "`n"
    }
    if (fileList = "") {
        MsgBox("图片文件夹中没有找到匹配的图片！", "错误", "IconX")
        return
    }
    
    fileList := Sort(fileList, "N")
    files := []
    Loop Parse, fileList, "`n" {
        if (A_LoopField != "") {
            parts := StrSplit(A_LoopField, "|")
            files.Push({num: parts[1], path: parts[2]})
        }
    }

    ; 打开 Word
    try {
        word := ComObject("Word.Application")
        word.Visible := true
        word.WindowState := 2
        word.DisplayAlerts := 0
        Sleep 1000
        doc := word.Documents.Open(DocPath)
        Sleep 500
    } catch as err {
        MsgBox("打开Word/文档失败：" . err.Message, "错误", "IconX")
        return
    }

    picIndex := 0
    insertedCount := 0
    tableCount := doc.Tables.Count
    
    Loop tableCount {
        table := doc.Tables.Item(A_Index)
        rowCount := table.Rows.Count
        
        Loop rowCount {
            rowIndex := A_Index
            if (Mod(rowIndex, 2) = 0) {
                continue
            }
            
            colCount := table.Columns.Count
            Loop colCount {
                colIndex := A_Index
                picIndex++
                if (picIndex > files.Length) {
                    break
                }
                
                try {
                    cell := table.Cell(rowIndex, colIndex)
                    rng := cell.Range
                    w := cell.Width
                    h := cell.Height
                    
                    ; 核心：先嵌入，再转换，完美居中且浮于上方
                    inlineShape := rng.InlineShapes.AddPicture(files[picIndex].path, false, true)
                    origW := inlineShape.Width
                    origH := inlineShape.Height
                    scale := Min(w / origW, h / origH)
                    inlineShape.Width := origW * scale
                    inlineShape.Height := origH * scale
                    rng.ParagraphFormat.Alignment := 1
                    shape := inlineShape.ConvertToShape()
                    shape.WrapFormat.Type := 6
                    shape.LockAspectRatio := -1
                    
                    insertedCount++
                } catch as err {
                    MsgBox("插入第 " . picIndex . " 张图片失败：" . err.Message, "错误", "IconX")
                    picIndex--
                    continue
                }
            }
            if (picIndex > files.Length) {
                break
            }
        }
        if (picIndex > files.Length) {
            break
        }
    }

    try {
        doc.Save()
        doc.Close(false)
        word.Quit()
    } catch {
        ; 忽略保存后可能的COM报错
    }
    
    MsgBox("处理完成！共插入 " . insertedCount . " 张图片。", "成功", "Iconi")
}