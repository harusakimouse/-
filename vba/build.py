import olefile, struct, cfb
from oletools.olevba import decompress_stream

SRC='x/xl/vbaProject.bin'
o=olefile.OleFileIO(SRC)
dirdec=bytearray(decompress_stream(bytearray(o.openstream('VBA/dir').read())))

# ---- parse module blocks ----
i=482
mods=[]; cur={}
offsets_pos=[]
while i < len(dirdec)-6:
    rid,sz=struct.unpack('<HI',dirdec[i:i+6])
    body=dirdec[i+6:i+6+sz]
    if rid==0x0019: cur={'name':body.decode('cp932')}
    elif rid==0x0031:
        cur['off']=struct.unpack('<I',body)[0]; offsets_pos.append(i+6)
    elif rid==0x002B:
        mods.append(cur); cur={}
    elif rid==0x0010: break
    i+=6+sz
print([ (m['name'],hex(m['off'])) for m in mods])
assert len(mods)==6

# ---- new sources ----
new_mod_name='Mod_RSS_MS2'
new_src=open('src_new.txt','rb').read().decode('utf-8').replace('\r\n','\n').replace('\n','\r\n').encode('cp932')

tw_old=decompress_stream(bytearray(o.openstream('VBA/ThisWorkbook').read()[0x03F8:]))
tw_add=('\r\n'
"Private Sub Workbook_Open()\r\n"
"    ' ブックが開き切ってからRSSを確かめる\r\n"
"    Application.OnTime Now, \"'\" & ThisWorkbook.Name & \"'!\" & \"自動RSS準備\"\r\n"
"End Sub\r\n"
"\r\n"
"Private Sub Workbook_BeforeClose(Cancel As Boolean)\r\n"
"    On Error Resume Next\r\n"
"    Application.Run \"'\" & ThisWorkbook.Name & \"'!\" & \"監視停止\"\r\n"
"End Sub\r\n").encode('cp932')
tw_new=tw_old+tw_add

# ---- build stream table ----
streams={}
for m in mods:
    raw=o.openstream('VBA/'+m['name']).read()
    src=decompress_stream(bytearray(raw[m['off']:]))
    if m['name']=='ThisWorkbook': src=tw_new
    streams[m['name']]=cfb.compress(src)
    assert decompress_stream(bytearray(streams[m['name']]))==src
streams[new_mod_name]=cfb.compress(new_src)
assert decompress_stream(bytearray(streams[new_mod_name]))==new_src

# ---- dir: count 6->7, offsets ->0, append module block ----
struct.pack_into('<H',dirdec,466+6,7)
for p in offsets_pos: struct.pack_into('<I',dirdec,p,0)

def rec(rid,body): return struct.pack('<HI',rid,len(body))+body
na=new_mod_name.encode('cp932'); nw=new_mod_name.encode('utf-16-le')
blk =rec(0x0019,na)+rec(0x0047,nw)+rec(0x001A,na)+rec(0x0032,nw)
blk+=rec(0x001C,b'')+rec(0x0048,b'')
blk+=rec(0x0031,struct.pack('<I',0))
blk+=rec(0x001E,struct.pack('<I',0))
blk+=rec(0x002C,struct.pack('<H',0xFFFF))
blk+=rec(0x0021,b'')+rec(0x002B,b'')
dirdec=dirdec[:-6]+blk+dirdec[-6:]
dir_c=cfb.compress(bytes(dirdec))
assert decompress_stream(bytearray(dir_c))==bytes(dirdec)

# ---- PROJECT ----
proj=o.openstream('PROJECT').read().decode('cp932')
proj=proj.replace('Module=Mod_Launcher\r\n','Module=Mod_Launcher\r\nModule=%s\r\n'%new_mod_name)
proj=proj.rstrip('\r\n')+'\r\n%s=0, 0, 0, 0, C\r\n'%new_mod_name
proj=proj.encode('cp932')

# ---- PROJECTwm ----
wm=o.openstream('PROJECTwm').read()
assert wm.endswith(b'\x00\x00')
wm=wm[:-2]+na+b'\x00'+nw+b'\x00\x00'+b'\x00\x00'

# ---- _VBA_PROJECT: keep header only (performance cache purged) ----
vp=o.openstream('VBA/_VBA_PROJECT').read()[:7]

# ---- assemble CFB ----
root=cfb.Entry('Root Entry',5,mtime=134332976562620000)
vba=cfb.Entry('VBA',1,ctime=134332976562590000,mtime=134332976562620000)
root.children=[vba,
    cfb.Entry('PROJECT',2,data=proj),
    cfb.Entry('PROJECTwm',2,data=wm)]
order=['Module1','ThisWorkbook','Sheet1','Module2','Sheet2','Mod_Launcher',new_mod_name]
vba.children=[cfb.Entry(n,2,data=streams[n]) for n in order]
vba.children.append(cfb.Entry('_VBA_PROJECT',2,data=vp))
vba.children.append(cfb.Entry('dir',2,data=dir_c))
cfb.write_cfb(root,'vbaProject_new.bin')
print('written',len(open('vbaProject_new.bin','rb').read()))
