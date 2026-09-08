import struct

FREESECT=0xFFFFFFFF; ENDOFCHAIN=0xFFFFFFFE; FATSECT=0xFFFFFFFD; NOSTREAM=0xFFFFFFFF

# ---------- MS-OVBA compression ----------
def compress(data:bytes)->bytes:
    out=bytearray(b'\x01')
    pos=0; n=len(data)
    while pos<n:
        chunk_start=pos
        chunk=bytearray()
        while pos<n and (pos-chunk_start)<4096:
            gpos=pos
            group=bytearray(); flags=0; nflag=0
            while nflag<8 and pos<n and (pos-chunk_start)<4096:
                diff=pos-chunk_start
                bc=4
                while (1<<bc)<diff: bc+=1
                if bc<4: bc=4
                if bc>12: bc=12
                maxlen=(0xFFFF>>bc)+3
                best_len=0; best_off=0
                if diff>0:
                    maxoff=min(diff,(1<<bc))
                    limit=min(maxlen, n-pos, 4096-(pos-chunk_start))
                    if limit>=3:
                        for off in range(1,maxoff+1):
                            s0=pos-off; l=0
                            while l<limit and data[s0+l]==data[pos+l]: l+=1
                            if l>best_len:
                                best_len=l; best_off=off
                                if l==limit: break
                if best_len>=3:
                    tok=((best_off-1)<<(16-bc)) | (best_len-3)
                    group+=struct.pack('<H',tok)
                    flags |= (1<<nflag)
                    pos+=best_len
                else:
                    group.append(data[pos]); pos+=1
                nflag+=1
            if len(chunk)+1+len(group) > 4095:
                pos=gpos
                break
            chunk.append(flags); chunk+=group
        hdr=(len(chunk)-1) | 0xB000
        out+=struct.pack('<H',hdr); out+=chunk
    return bytes(out)

def decompressed_len(chunk:bytes)->int:
    # length of decompressed data represented by a (compressed) chunk body
    i=0; dec=0
    while i<len(chunk):
        fb=chunk[i]; i+=1
        for b in range(8):
            if i>=len(chunk): break
            if fb>>b & 1:
                tok=struct.unpack('<H',chunk[i:i+2])[0]; i+=2
                bc=4
                while (1<<bc)<dec: bc+=1
                if bc<4: bc=4
                if bc>12: bc=12
                lm=0xFFFF>>bc
                dec+=(tok&lm)+3
            else:
                i+=1; dec+=1
    return dec

# ---------- CFB writer ----------
class Entry:
    def __init__(self,name,etype,clsid=b'\0'*16,data=b'',state=0,ctime=0,mtime=0):
        self.name=name; self.etype=etype; self.clsid=clsid; self.data=data
        self.state=state; self.ctime=ctime; self.mtime=mtime
        self.children=[]; self.sid=NOSTREAM
        self.left=NOSTREAM; self.right=NOSTREAM; self.child=NOSTREAM
        self.start=ENDOFCHAIN; self.size=0

def _key(e):
    u=e.name.upper()
    return (len(e.name), [ord(c) for c in u])

def _build_tree(nodes):
    if not nodes: return NOSTREAM
    m=len(nodes)//2
    node=nodes[m]
    node.left=_build_tree(nodes[:m])
    node.right=_build_tree(nodes[m+1:])
    return node.sid

def write_cfb(root, path):
    # assign sids depth-first
    allents=[]
    def collect(e):
        e.sid=len(allents); allents.append(e)
        for c in e.children: collect(c)
    collect(root)
    for e in allents:
        if e.children:
            kids=sorted(e.children,key=_key)
            e.child=_build_tree(kids)

    MINI_CUTOFF=4096; SEC=512; MINI=64
    mini=bytearray(); mini_map={}
    for e in allents:
        if e.etype==2 and len(e.data)<MINI_CUTOFF:
            if len(e.data)==0:
                e.start=ENDOFCHAIN; e.size=0; continue
            e.start=len(mini)//MINI
            mini+=e.data
            while len(mini)%MINI: mini.append(0)
            e.size=len(e.data)
    nmini=len(mini)//MINI
    # normal sector allocation
    sectors=[]  # list of bytes(512)
    def alloc(data):
        if not data: return ENDOFCHAIN,0
        start=len(sectors)
        for i in range(0,len(data),SEC):
            blk=data[i:i+SEC]
            blk=blk+b'\0'*(SEC-len(blk))
            sectors.append(blk)
        return start,(len(data)+SEC-1)//SEC
    for e in allents:
        if e.etype==2 and len(e.data)>=MINI_CUTOFF:
            e.start,_=alloc(e.data); e.size=len(e.data)
    mini_start,mini_secs=alloc(bytes(mini))
    root.start=mini_start; root.size=len(mini)
    # minifat
    mf=bytearray()
    for e in allents:
        if e.etype==2 and 0<len(e.data)<MINI_CUTOFF:
            cnt=(len(e.data)+MINI-1)//MINI
            for k in range(cnt):
                mf+=struct.pack('<I', e.start+k+1 if k<cnt-1 else ENDOFCHAIN)
    while len(mf)%SEC: mf+=struct.pack('<I',FREESECT)
    minifat_start,minifat_secs=alloc(bytes(mf))
    if minifat_secs==0: minifat_start=ENDOFCHAIN
    # directory
    dirbuf=bytearray()
    for e in allents:
        nm=e.name.encode('utf-16-le')+b'\0\0'
        assert len(nm)<=64
        d=bytearray(128)
        d[0:len(nm)]=nm
        struct.pack_into('<H',d,64,len(nm))
        d[66]=e.etype; d[67]=1
        struct.pack_into('<III',d,68,e.left,e.right,e.child)
        d[80:96]=e.clsid
        struct.pack_into('<I',d,96,e.state)
        struct.pack_into('<Q',d,100,e.ctime)
        struct.pack_into('<Q',d,108,e.mtime)
        struct.pack_into('<I',d,116,e.start if e.etype in (2,5) else 0)
        struct.pack_into('<Q',d,120,e.size if e.etype in (2,5) else 0)
        dirbuf+=d
    while len(dirbuf)%SEC:
        d=bytearray(128)
        struct.pack_into('<III',d,68,NOSTREAM,NOSTREAM,NOSTREAM)
        dirbuf+=d
    dir_start,dir_secs=alloc(bytes(dirbuf))
    ndata=len(sectors)
    nfat=1
    while True:
        if (ndata+nfat)<=nfat*(SEC//4): break
        nfat+=1
    assert nfat<=109, "DIFAT needed"
    total=ndata+nfat
    fat=[FREESECT]*(nfat*(SEC//4))
    def chain(start,cnt):
        for k in range(cnt):
            fat[start+k]= start+k+1 if k<cnt-1 else ENDOFCHAIN
    # rebuild chains
    idx=0
    for e in allents:
        if e.etype==2 and len(e.data)>=MINI_CUTOFF:
            cnt=(len(e.data)+SEC-1)//SEC; chain(e.start,cnt)
    if mini_secs: chain(mini_start,mini_secs)
    if minifat_secs: chain(minifat_start,minifat_secs)
    chain(dir_start,dir_secs)
    for k in range(nfat): fat[ndata+k]=FATSECT
    hdr=bytearray(512)
    hdr[0:8]=bytes.fromhex('D0CF11E0A1B11AE1')
    struct.pack_into('<HH',hdr,24,0x003E,0x0003)
    struct.pack_into('<H',hdr,28,0xFFFE)
    struct.pack_into('<HH',hdr,30,9,6)
    struct.pack_into('<I',hdr,44,nfat)
    struct.pack_into('<I',hdr,48,dir_start)
    struct.pack_into('<I',hdr,56,0x1000)
    struct.pack_into('<I',hdr,60,minifat_start)
    struct.pack_into('<I',hdr,64,minifat_secs)
    struct.pack_into('<I',hdr,68,ENDOFCHAIN)
    struct.pack_into('<I',hdr,72,0)
    for k in range(109):
        struct.pack_into('<I',hdr,76+4*k, ndata+k if k<nfat else FREESECT)
    with open(path,'wb') as f:
        f.write(hdr)
        for s in sectors: f.write(s)
        fb=b''.join(struct.pack('<I',v) for v in fat)
        f.write(fb)
