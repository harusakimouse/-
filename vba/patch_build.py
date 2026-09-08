s=open('build.py',encoding='utf-8').read()
if 'def patch_wait' not in s:
    helper = (
        "\ndef patch_wait(name, src):\n"
        "    t = src.decode('cp932')\n"
        "    if name=='Module1':\n"
        "        old = '            ok = ok + 1\\r\\n        Else\\r\\n'\n"
        "        new = '            ok = ok + 1\\r\\n            少し待つ 3\\r\\n        Else\\r\\n'\n"
        "    else:\n"
        "        old = ('                        ws.Cells(r, \"E\").Value = \"起動OK(相手マクロ警告 \" & errNo & \")\"\\r\\n'\n"
        "               '                    End If\\r\\n')\n"
        "        new = ('                        ws.Cells(r, \"E\").Value = \"起動OK(相手マクロ警告 \" & errNo & \")\"\\r\\n'\n"
        "               '                    End If\\r\\n'\n"
        "               '                    少し待つ 3\\r\\n')\n"
        "    assert t.count(old)==1, (name, t.count(old))\n"
        "    return t.replace(old,new).encode('cp932')\n\n"
    )
    s = s.replace("# ---- build stream table ----", helper + "# ---- build stream table ----")
    s = s.replace("    if m['name']=='ThisWorkbook': src=tw_new\n",
                  "    if m['name']=='ThisWorkbook': src=tw_new\n"
                  "    if m['name'] in ('Module1','Mod_Launcher'): src=patch_wait(m['name'],src)\n")
    open('build.py','w',encoding='utf-8').write(s)
print('patched')
