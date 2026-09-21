import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import os, json, time, threading, subprocess, shutil, socket, struct, base64, re, html, urllib.request, urllib.parse
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

KAKAO_SEARCH_URL = "https://m.map.kakao.com/actions/searchView"
KAKAO_PANEL_URL = "https://place-api.map.kakao.com/places/panel3/"

KAKAO_BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "ko,en-US;q=0.9,en;q=0.8",
}
KAKAO_PANEL_HEADERS = {
    **KAKAO_BROWSER_HEADERS,
    "Accept": "application/json, text/plain, */*",
    "Origin": "https://place.map.kakao.com",
    "Referer": "https://place.map.kakao.com/",
    "appVersion": "6.6.0",
    "pf": "PC",
}

def _norm(v):
    return re.sub(r"[^0-9A-Za-z가-힣]+", "", str(v or "")).lower()

def _strip_bl(v):
    return re.sub(r"\s+(?:BL|BLK)$", "", str(v or "").strip(), flags=re.I)

def _decode_html(v):
    return html.unescape(re.sub(r"<[^>]+>", " ", str(v or ""))).replace("&nbsp;", " ").strip()

def _search_candidates(text):
    items = []
    pattern = re.compile(r'<li\s+class="search_item\s+base"([\s\S]*?)</li>', re.I)
    for m in pattern.finditer(text or ""):
        frag = m.group(1)
        mid = re.search(r'data-id="([^"]+)"', frag, re.I)
        title = re.search(r'data-title="([^"]*)"', frag, re.I)
        if not mid:
            continue
        cid = html.unescape(mid.group(1)).strip()
        name = html.unescape(title.group(1)).strip() if title else ""
        if not name:
            mt = re.search(r'<[^>]+class="[^"]*tit_g[^"]*"[^>]*>([\s\S]*?)</', frag, re.I)
            name = _decode_html(mt.group(1)) if mt else ""
        spans = [_decode_html(x) for x in re.findall(r'<span class="txt_g">([\s\S]*?)</span>', frag, re.I)]
        address = spans[-1] if spans else ""
        if cid:
            items.append({"id": cid, "name": name, "address": address})
    return items

def _panel_point(obj):
    if isinstance(obj, dict):
        summary = obj.get("summary")
        if isinstance(summary, dict):
            point = summary.get("point")
            if isinstance(point, dict):
                lat, lon = point.get("lat"), point.get("lon")
                if lat not in (None, "") and lon not in (None, ""):
                    return str(lat), str(lon), summary
        for v in obj.values():
            found = _panel_point(v)
            if found:
                return found
    elif isinstance(obj, list):
        for v in obj:
            found = _panel_point(v)
            if found:
                return found
    return None

def search_address(address):
    query = str(address or "").strip()
    variants = []
    for q in (query, _strip_bl(query), re.sub(r"\s+", " ", _strip_bl(query))):
        q = q.strip()
        if q and q not in variants:
            variants.append(q)

    last_error = None
    for q in variants:
        try:
            url = KAKAO_SEARCH_URL + "?" + urllib.parse.urlencode({"q": q})
            req = urllib.request.Request(url, headers=KAKAO_BROWSER_HEADERS)
            with urllib.request.urlopen(req, timeout=12) as r:
                text = r.read().decode("utf-8", "ignore")
            candidates = _search_candidates(text)
            if not candidates:
                continue

            nq = _norm(_strip_bl(q))
            def score(item):
                a = _norm(_strip_bl(item.get("address", "")))
                n = _norm(item.get("name", ""))
                score = 0
                if a == nq: score += 2000
                if nq and (nq in a or a in nq): score += 1200
                if nq and nq in n: score += 300
                return score

            candidates.sort(key=score, reverse=True)

            for item in candidates[:8]:
                pid = item.get("id")
                if not pid:
                    continue
                try:
                    preq = urllib.request.Request(
                        KAKAO_PANEL_URL + urllib.parse.quote(str(pid), safe=""),
                        headers=KAKAO_PANEL_HEADERS
                    )
                    with urllib.request.urlopen(preq, timeout=12) as r:
                        panel = json.loads(r.read().decode("utf-8", "ignore"))
                    point = _panel_point(panel)
                    if point:
                        lat, lon, summary = point
                        address_name = ""
                        road_name = ""
                        if isinstance(summary, dict):
                            sa = summary.get("address")
                            if isinstance(sa, dict):
                                address_name = str(sa.get("disp") or sa.get("address_name") or "")
                            ra = summary.get("road_address")
                            if isinstance(ra, dict):
                                road_name = str(ra.get("address_name") or "")
                        return {
                            "ok": True,
                            "addr": address_name or item.get("address") or query,
                            "road": road_name,
                            "lat": lat,
                            "lng": lon,
                            "id": str(pid),
                        }
                except Exception as e:
                    last_error = e
                    continue
        except Exception as e:
            last_error = e
            continue

    if last_error:
        return {"ok": False, "error": "카카오맵 주소검색 통신 오류: {}".format(last_error)}
    return {"ok": True, "items": []}

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
                res=search_address(row[4])
                if res.get("error"): raise RuntimeError(res.get("error"))
                if not res.get("ok"): raise RuntimeError("주소 검색에 실패했습니다.")
                if not res.get("lat") or not res.get("lng"): raise RuntimeError("카카오맵에서 주소 결과를 찾지 못했습니다: {}".format(row[4]))
                t=res
                if not t.get("lat") or not t.get("lng"): raise RuntimeError("카카오맵에서 주소 좌표를 확인하지 못했습니다.")
                self.msg("[{}/{}] 검색 성공: {} / 좌표 X={}, Y={}".format(n,total,row[4],t["lng"],t["lat"]))
                obj=json.dumps({"display1":t.get("addr") or row[4],"x":t["lng"],"y":t["lat"],"folderId":fid,"memo":name},ensure_ascii=False)
                a=self.c.js(ADD.replace("__O__",obj),20) or {}
                if int(a.get("http",0)) not in (200,201): raise RuntimeError("즐겨찾기 저장 HTTP {} / {}".format(a.get("http"),str(a.get("text",""))[:300]))
                self.msg("[{}/{}] 즐겨찾기 등록 확인: HTTP {}".format(n,total,a.get("http")))
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
