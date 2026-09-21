import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import os, json, time, threading, subprocess, shutil, socket, struct, base64, urllib.request, urllib.parse
from openpyxl import load_workbook, Workbook

APP = Path(os.environ.get("APPDATA", str(Path.home()))) / "KakaoFavoriteMacro"
PROFILE = APP / "chrome_profile"
RESULT = APP / "result"
ERROR = APP / "error-excel"
STATE = APP / "state.json"
PORT = 9222
MAP = "https://map.kakao.com/"
for d in (APP, PROFILE, RESULT, ERROR): d.mkdir(parents=True, exist_ok=True)

def chrome():
    ps = [
        os.path.join(os.environ.get("PROGRAMFILES",""), "Google","Chrome","Application","chrome.exe"),
        os.path.join(os.environ.get("PROGRAMFILES(X86)",""), "Google","Chrome","Application","chrome.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA",""), "Google","Chrome","Application","chrome.exe")
    ]
    return next((p for p in ps if os.path.exists(p)), shutil.which("chrome.exe") or shutil.which("chrome"))

def get_json(path):
    with urllib.request.urlopen("http://127.0.0.1:{}{}".format(PORT,path),timeout=5) as r:
        return json.loads(r.read().decode("utf-8"))

class WS:
    def __init__(self,url):
        u=urllib.parse.urlparse(url)
        self.s=socket.create_connection((u.hostname,u.port),timeout=10)
        k=base64.b64encode(os.urandom(16)).decode()
        req="GET {} HTTP/1.1\r\nHost: {}:{}\r\nUpgrade: websocket\r\nConnection: Upgrade\r\nSec-WebSocket-Key: {}\r\nSec-WebSocket-Version: 13\r\n\r\n".format(u.path or "/",u.hostname,u.port,k)
        self.s.sendall(req.encode())
        b=b""
        while b"\r\n\r\n" not in b: b+=self.s.recv(4096)
    def read(self,n):
        b=b""
        while len(b)<n:
            x=self.s.recv(n-len(b))
            if not x: raise RuntimeError("Chrome 연결이 종료되었습니다.")
            b+=x
        return b
    def send(self,o):
        raw=json.dumps(o,separators=(",",":")).encode(); L=len(raw)
        h=bytearray([129])
        if L<126: h.append(128|L)
        elif L<65536: h.extend([254]); h.extend(struct.pack("!H",L))
        else: h.extend([255]); h.extend(struct.pack("!Q",L))
        m=os.urandom(4); h.extend(m); h.extend(bytes(x^m[i%4] for i,x in enumerate(raw))); self.s.sendall(h)
    def recv(self):
        h=self.read(2); b2=h[1]; L=b2&127
        if L==126: L=struct.unpack("!H",self.read(2))[0]
        elif L==127: L=struct.unpack("!Q",self.read(8))[0]
        m=self.read(4) if b2&128 else b""; b=self.read(L)
        if m: b=bytes(x^m[i%4] for i,x in enumerate(b))
        if h[0]&15==1: return json.loads(b.decode())
        return self.recv()

class CDP:
    def __init__(self,url): self.w=WS(url); self.i=0
    def call(self,method,params=None,timeout=20):
        self.i+=1; i=self.i; self.w.send({"id":i,"method":method,"params":params or {}})
        end=time.time()+timeout
        while time.time()<end:
            r=self.w.recv()
            if r.get("id")==i:
                if "error" in r: raise RuntimeError(str(r["error"]))
                return r.get("result",{})
        raise TimeoutError(method)
    def js(self,x,timeout=25):
        r=self.call("Runtime.evaluate",{"expression":x,"awaitPromise":True,"returnByValue":True},timeout)
        return r.get("result",{}).get("value")
    def nav(self,url):
        self.call("Page.navigate",{"url":url}); time.sleep(1)

def connect():
    for t in get_json("/json/list"):
        if t.get("type")=="page" and t.get("webSocketDebuggerUrl"):
            return CDP(t["webSocketDebuggerUrl"])

def start():
    p=chrome()
    if not p: raise RuntimeError("Google Chrome을 찾지 못했습니다.")
    subprocess.Popen([p,"--remote-debugging-port={}".format(PORT),"--user-data-dir={}".format(PROFILE),"--no-first-run","--no-default-browser-check",MAP],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    for _ in range(40):
        time.sleep(.5)
        try:
            c=connect()
            if c: return c
        except Exception: pass
    raise RuntimeError("Chrome 연결에 실패했습니다. Chrome을 모두 종료한 후 다시 시도하세요.")

SEARCH = r"""(async function(q){
const i=document.querySelector('#search\.keyword\.query');
if(!i)return {ok:false,error:'검색창을 찾지 못했습니다.'};
i.focus();i.value=q;i.dispatchEvent(new Event('input',{bubbles:true}));
const b=document.querySelector('#search\.keyword\.submit');if(b)b.click();else i.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',code:'Enter',bubbles:true}));
await new Promise(r=>setTimeout(r,1800));
return {ok:true,items:[...document.querySelectorAll('#info\.search\.place\.list .PlaceItem,.placelist .PlaceItem')].map((e,n)=>({i:n,addr:[...e.querySelectorAll('.addr,.jibun,.road')].map(x=>x.innerText.trim()).filter(Boolean).join(' | '),lat:e.getAttribute('data-lat')||e.dataset.lat||'',lng:e.getAttribute('data-lng')||e.dataset.lng||'',href:e.querySelector('.moreview')?.href||'',txt:(e.innerText||'').slice(0,1000)}))};
})(__Q__)"""

FOLDERS = "fetch('/folder/list.json?sort=CREATE_AT').then(r=>r.json()).then(x=>x.result||[]).catch(e=>[])"
ADD = r"""(async function(o){
const body=JSON.stringify([{type:'address',key:'N3'+Math.floor(1000000+Math.random()*9000000),display1:o.display1,display2:'',x:Math.round(Number(o.x)),y:Math.round(Number(o.y)),color:'01',folderId:String(o.folderId),memo:o.memo}]);
const r=await fetch('/favorite/add.json',{method:'POST',headers:{'Content-Type':'application/json'},body:body});
return {http:r.status,text:await r.text()};
})(__O__)"""

class App:
    def __init__(self,root):
        self.r=root; self.c=None; self.rows=[]; self.folders=[]; self.stopflag=False; self.running=False
        root.title("카카오맵 주소 즐겨찾기 자동 등록"); root.geometry("780x620")
        self.ui(); self.load()
    def ui(self):
        f=ttk.Frame(self.r,padding=16); f.pack(fill="both",expand=True)
        ttk.Label(f,text="카카오맵 주소 즐겨찾기 자동 등록",font=("Malgun Gothic",18,"bold")).pack(anchor="w")
        ttk.Label(f,text="Chrome 로그인 후 엑셀 주소를 검색하여 선택한 그룹에 등록합니다.").pack(anchor="w",pady=5)
        bar=ttk.Frame(f); bar.pack(fill="x")
        ttk.Button(bar,text="① Chrome 시작 / 로그인",command=self.connect).pack(side="left")
        ttk.Button(bar,text="② 엑셀 선택",command=self.pick).pack(side="left",padx=8)
        self.file=ttk.Label(bar,text="파일: 선택 안 됨"); self.file.pack(side="left",fill="x",expand=True)
        g=ttk.LabelFrame(f,text="즐겨찾기 그룹",padding=10); g.pack(fill="x",pady=10)
        self.combo=ttk.Combobox(g,state="readonly",width=55); self.combo.pack(side="left")
        ttk.Button(g,text="그룹 새로고침",command=self.groups).pack(side="left",padx=8)
        o=ttk.LabelFrame(f,text="즐겨찾기 이름",padding=10); o.pack(fill="x")
        self.mode=tk.StringVar(value="space")
        ttk.Radiobutton(o,text="1~4열을 공백으로 연결",variable=self.mode,value="space").pack(anchor="w")
        ttk.Radiobutton(o,text="1~4열을 줄바꿈으로 연결",variable=self.mode,value="line").pack(anchor="w")
        self.pb=ttk.Progressbar(f,maximum=100); self.pb.pack(fill="x",pady=10)
        self.status=tk.StringVar(value="준비됨"); ttk.Label(f,textvariable=self.status).pack(anchor="w")
        self.log=tk.Text(f,height=15,font=("Consolas",9)); self.log.pack(fill="both",expand=True,pady=5)
        b=ttk.Frame(f); b.pack(fill="x")
        ttk.Button(b,text="▶ 등록 시작",command=self.begin).pack(side="left")
        ttk.Button(b,text="■ 중지",command=self.stop).pack(side="left",padx=8)
    def msg(self,s):
        self.r.after(0,lambda:(self.log.insert("end",s+"\n"),self.log.see("end")))
    def connect(self):
        def w():
            try:
                self.c=self.c or start(); self.msg("Chrome 연결 완료. Chrome에서 카카오 로그인/2차 인증을 완료하세요."); self.status.set("Chrome 연결됨")
            except Exception as e: self.r.after(0,lambda:messagebox.showerror("오류",str(e)))
        threading.Thread(target=w,daemon=True).start()
    def pick(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx")])
        if not p:return
        try:
            ws=load_workbook(p,data_only=True).active
            if ws.max_column<5: raise ValueError("엑셀은 최소 5개 열이 필요합니다.")
            self.rows=[]
            for n,row in enumerate(ws.iter_rows(min_row=2,values_only=True),2):
                if not any(str(x or "").strip() for x in row[:5]): continue
                self.rows.append([str(row[i] or "").strip() for i in range(5)])
            self.file.config(text="파일: {} / {}건".format(Path(p).name,len(self.rows))); self.msg("엑셀 {}건 불러옴".format(len(self.rows)))
        except Exception as e: messagebox.showerror("엑셀 오류",str(e))
    def groups(self):
        if not self.c: messagebox.showwarning("안내","먼저 Chrome을 시작하세요."); return
        def w():
            try:
                self.c.nav(MAP); self.folders=self.c.js(FOLDERS) or []
                names=["{} [ID {}]".format(x.get("title",""),x.get("folderId","")) for x in self.folders]
                self.r.after(0,lambda:self.combo.config(values=names))
                if names:self.r.after(0,lambda:self.combo.current(0))
                self.msg("그룹 {}개 확인".format(len(names)))
            except Exception as e:self.msg("그룹 조회 실패: "+str(e))
        threading.Thread(target=w,daemon=True).start()
    def load(self):
        try:
            self.last=json.loads(STATE.read_text(encoding="utf-8"))
        except Exception:self.last={}
    def stop(self): self.stopflag=True; self.status.set("중지 요청됨")
    def begin(self):
        if self.running:return
        if not self.c or not self.rows or not self.folders or self.combo.current()<0:
            messagebox.showwarning("안내","Chrome 시작 → 로그인 → 엑셀 선택 → 그룹 새로고침 순서로 진행하세요."); return
        if len(self.rows)>500 and not messagebox.askyesno("500건 초과","현재 {}건입니다. 한 그룹의 제한을 넘을 수 있습니다. 계속할까요?".format(len(self.rows))): return
        fid=str(self.folders[self.combo.current()].get("folderId"))
        threading.Thread(target=self.run,args=(fid,),daemon=True).start()
    def run(self,fid):
        self.running=True; self.stopflag=False; ok=fail=0; total=len(self.rows)
        out=Workbook(); ow=out.active; ow.append(["순번","선로명","선로번호","전산화번호","주소","상태","비고"])
        er=Workbook(); ew=er.active; ew.append(["순번","선로명","선로번호","전산화번호","주소","오류"])
        for n,row in enumerate(self.rows,1):
            if self.stopflag: break
            self.pb.config(value=n/total*100); self.status.set("{}/{} 처리 중: {}".format(n,total,row[4]))
            name=("\n" if self.mode.get()=="line" else " ").join(row[:4])
            try:
                res=self.c.js(SEARCH.replace("__Q__",json.dumps(row[4],ensure_ascii=False)),25) or {}
                items=res.get("items",[])
                if not items: raise RuntimeError("검색 결과가 없습니다.")
                t=items[0]
                for x in items:
                    if row[4].replace(" ","") in x.get("txt","").replace(" ",""): t=x; break
                if not t.get("lat") or not t.get("lng"):
                    if t.get("href"):
                        self.c.nav(t["href"]); time.sleep(1)
                        q=self.c.js("(()=>{let s=document.body.innerText||'';let m=s.match(/(-?\\d+\\.\\d+)\\s*,\\s*(-?\\d+\\.\\d+)/);return m?{x:m[1],y:m[2]}:null})()")
                        if q:t["lng"]=q["x"];t["lat"]=q["y"]
                        self.c.nav(MAP);time.sleep(.5)
                if not t.get("lat") or not t.get("lng"): raise RuntimeError("검색 결과 좌표를 확인하지 못했습니다.")
                obj=json.dumps({"display1":t.get("addr") or row[4],"x":t["lng"],"y":t["lat"],"folderId":fid,"memo":name},ensure_ascii=False)
                a=self.c.js(ADD.replace("__O__",obj),20) or {}
                if int(a.get("http",0)) not in (200,201): raise RuntimeError("즐겨찾기 저장 HTTP {}".format(a.get("http")))
                ok+=1; ow.append(row+["성공",""]); self.msg("[{}/{}] 성공: {}".format(n,total,row[4]))
            except Exception as e:
                fail+=1; m=str(e); ew.append(row+[m]); ow.append(row+["실패",m]); self.msg("[{}/{}] 실패: {} / {}".format(n,total,row[4],m))
                try:self.c.nav(MAP)
                except Exception:pass
        ts=time.strftime("%Y%m%d_%H%M%S"); rp=RESULT/"result_{}.xlsx".format(ts); ep=ERROR/"error_{}.xlsx".format(ts)
        out.save(rp); er.save(ep); self.running=False; self.status.set("완료: 성공 {} / 실패 {}".format(ok,fail))
        self.r.after(0,lambda:messagebox.showinfo("작업 완료","성공 {}건 / 실패 {}건\n\n결과: {}\n실패목록: {}".format(ok,fail,rp,ep)))
if __name__=="__main__":
    root=tk.Tk(); App(root); root.mainloop()
