#SingleInstance force
ListLines 0
KeyHistory 0
SendMode "Input" ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir A_ScriptDir ; Ensures a consistent starting directory.

; https://github.com/ayaka14732/TinyPE-on-Win10/tree/main#:~:text=To%20understand%20the%20structure%20of%20the%20PE%20file
fileObj:=FileOpen(A_WinDir "\System32\calc.exe","r")
fileObj.Seek(0x3c)
offset_pe:=fileObj.ReadUInt()
fileObj.Seek(offset_pe + 0x90)
offset_itbl:=fileObj.ReadUInt()
size_itbl:=fileObj.ReadUInt()
pointer_itbl:=offset_itbl
end_itbl:=pointer_itbl+size_itbl
byDllNamesArr:=[]
while (pointer_itbl < end_itbl) {
    fileObj.Seek(pointer_itbl + 0xC)
    offset_dllName:=fileObj.ReadUInt()
    FirstThunk:=fileObj.ReadUInt()
    if (!offset_dllName) {
        break
    }
    fileObj.Seek(pointer_itbl)
    offset_importNames:=fileObj.ReadUInt()
    readStr(fileObj) {
        importName:=""
        while ((chat_importName:=fileObj.ReadUChar())) {
            importName.=Chr(chat_importName)
        }
        return importName
    }
    fileObj.Seek(offset_dllName)
    dllName:=readStr(fileObj)
    pointer_importNames:=offset_importNames
    fileObj.Seek(pointer_importNames)
    importNames:=[]
    while ((offset_importName:=fileObj.ReadInt64())) {
        fileObj.Seek(offset_importName + 0x2)
        importName:=readStr(fileObj)
        importNames.Push(importName)
        pointer_importNames+=8
        fileObj.Seek(pointer_importNames)
    }
    byDllNamesArr.Push([dllName,importNames])
    pointer_itbl+=20
}
MsgBox A_Clipboard:=JSON_stringify(byDllNamesArr)

return

f3::Exitapp

JSON_stringify(obj, maxDepth := 5) {

    stringified := ""

    escape(str) {
        str:=StrReplace(str, "\", "\\", true)
        str:=StrReplace(str, "`t", "\t", true)
        str:=StrReplace(str, "`b", "\b", true)
        str:=StrReplace(str, "`n", "\n", true)
        str:=StrReplace(str, "`r", "\r", true)
        str:=StrReplace(str, "`f", "\f", true)
        str:=StrReplace(str, "`"", "\`"", true)
        return str
    }
    ok(obj, depth) {
        switch (Type(obj)) {
            case 'Map':
                if (depth > maxDepth) {
                    stringified.="`"[DEEP ...Map]`""
                } else {
                    stringified.="{"
                    for k, v in obj {
                        (A_Index > 1 && stringified.=",")
                        ;ESCAPE THIS, using java thingy
                        stringified.="`"" escape(k) "`": "
                        ok(v, depth+1)
                    }
                    stringified.="}"
                }
            case 'Object':
                if (depth > maxDepth) {
                    stringified.="`"[DEEP ...Object]`""
                } else {
                    stringified.="{"
                    for k, v in obj.OwnProps() {
                        (A_Index > 1 && stringified.=",")
                        ;ESCAPE THIS, using java thingy
                        stringified.="`"" escape(k) "`": "
                        ok(v, depth+1)
                    }
                    stringified.="}"
                }
            case 'Array':
                if (depth > maxDepth) {
                    stringified.="`"[DEEP ...Array]`""
                } else {
                    stringified.="["
                    for v in obj {
                        (A_Index > 1 && stringified.=",")
                        ok(v, depth+1)
                    }
                    stringified.="]"
                }
            case 'String':
                ; escape with java
                stringified.="`"" escape(obj) "`"" ;in order to escape \n and etc
            case "Integer", "Float":
                stringified.=obj
        }

    }
    ok(obj, 0)

    return stringified


}