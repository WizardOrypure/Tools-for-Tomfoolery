import os, sys, time, json, datetime as dt, traceback
from typing import Dict, List, Optional
import msal, requests
TENANT_ID     = os.getenv("TENANT_ID", "INSERT HERE")
CLIENT_ID     = os.getenv("CLIENT_ID", "INSERT HERE")
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "INSERT HERE")
GRAPH_BASE    = "https://graph.microsoft.com/v1.0"
SCOPES        = ["https://graph.microsoft.com/.default"]
TIMEOUT=2
MAX_RETRIES=5
RETRY_STATUS={429,503,504}
DEBUG=True

MENU_FIELDS:Dict[str,tuple[str,str,str]]={
 "1":("accountEnabled","Account Enabled","Whether the account is enabled"),
 "2":("lastPasswordChangeDateTime","Last Password Change","Password last reset/change time"),
 "3":("officeLocation","Office Location","Office/desk location"),
 "4":("businessPhones","Business Phones","List of business phone numbers"),
 "5":("mobilePhone","Mobile Phone","User's mobile phone number"),
 "6":("signInActivity","Sign-in Activity","Recent sign-in activity timestamps"),
 "7":("exit","Nothing, let me go","Exit immediately"),
}
CORE_FIELDS=["id","userPrincipalName","displayName"]

def debug(msg):
    if DEBUG: print(f"[DEBUG] {msg}",file=sys.stderr)

def acquire_token():
    debug("Acquiring token...")
    app=msal.ConfidentialClientApplication(CLIENT_ID,authority=f"https://login.microsoftonline.com/{TENANT_ID}",client_credential=CLIENT_SECRET)
    r=app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in r: raise RuntimeError(f"Token acquisition failed: {r}")
    debug("Token acquired."); return r["access_token"]

def graph_get(url,token,params=None,extra_headers=None):
    a=0; h={"Authorization":f"Bearer {token}"}; h.update(extra_headers or {})
    while True:
        a+=1; resp=requests.get(url,headers=h,params=params,timeout=TIMEOUT)
        if resp.status_code in RETRY_STATUS and a<MAX_RETRIES:
            w=int(resp.headers.get("Retry-After","2")); debug(f"Retryable {resp.status_code} on {url}; sleeping {w}s ({a}/{MAX_RETRIES})"); time.sleep(w); continue
        if not resp.ok: raise requests.HTTPError(f"GET {url} failed [{resp.status_code}]: {resp.text}",response=resp)
        return resp.json()

def deep_merge_defaults(d,x):
    o={}
    for k,v in d.items():
        if isinstance(v,dict):
            dv=x.get(k,{})
            o[k]=deep_merge_defaults(v,dv if isinstance(dv,dict) else {})
        else:o[k]=x.get(k,v)
    for k,v in x.items():
        if k not in o:o[k]=v
    return o

def build_defaults(sel):
    z={"id":"","displayName":"","userPrincipalName":""}
    if "accountEnabled" in sel:z["accountEnabled"]=True
    if "lastPasswordChangeDateTime" in sel:z["lastPasswordChangeDateTime"]=None
    if "officeLocation" in sel:z["officeLocation"]=""
    if "businessPhones" in sel:z["businessPhones"]=[] 
    if "mobilePhone" in sel:z["mobilePhone"]=""
    if "signInActivity" in sel:z["signInActivity"]={"lastSignInDateTime":None,"lastSignInRequestId":"","lastNonInteractiveSignInDateTime":None,"lastNonInteractiveSignInRequestId":"","lastSuccessfulSignInDateTime":None,"lastSuccessfulSignInRequestId":""}
    return z

def normalize_user(raw,defs): return deep_merge_defaults(defs,raw or {})

def ask_user_selection():
    print("\n==============================")
    print("    Coded by WizardOrypure    ")
    print("==============================")
    print("What do you want to parse? (multi-select)")
    for k,(_,label,desc) in MENU_FIELDS.items(): print(f"  {k}. {label} — {desc}")
    print('Type numbers like 2,4,6 or "all". Press Enter for none.')
    print("==============================")
    print("    Coded by WizardOrypure    ")
    print("==============================")
    while True:
        c=input("Selection: ").strip().lower()
        if c in ("all","a","*"): return [v[0] for v in MENU_FIELDS.values() if v[0]!="exit"]
        if c in ("7","exit","quit"): print("Bro why even open this then."); sys.exit(0)
        if c=="": return []
        p=[q.strip() for q in c.split(",") if q.strip()]
        if all(q in MENU_FIELDS for q in p):
            if "7" in p: print("Bro why even open this then."); sys.exit(0)
            fs=[MENU_FIELDS[q][0] for q in p]; s=set(); out=[]
            for f in fs:
                if f not in s:s.add(f); out.append(f)
            return out
        print("Invalid input.")

def list_all_users(token,sel):
    debug("Fetching users...")
    url=f"{GRAPH_BASE}/users"; select=",".join(CORE_FIELDS+sel)
    top="500" if "signInActivity" in sel else "999"
    params={"$select":select,"$top":top}
    res=[]; page=graph_get(url,token,params=params); res+=page.get("value",[])
    while "@odata.nextLink" in page: debug("Paging..."); page=graph_get(page["@odata.nextLink"],token); res+=page.get("value",[])
    debug(f"Fetched {len(res)} users"); return res

def main():
    if not TENANT_ID or not CLIENT_ID or not CLIENT_SECRET:
        print("Set TENANT_ID, CLIENT_ID, CLIENT_SECRET.",file=sys.stderr); sys.exit(2)
    sel=ask_user_selection(); debug(f"Selected: {', '.join(sel) if sel else '(core only)'}")
    defs=build_defaults(sel)
    try:
        token=acquire_token(); users=list_all_users(token,sel); results=[]
        from datetime import datetime,timezone
        try: from zoneinfo import ZoneInfo; local_tz=ZoneInfo("Australia/Melbourne")
        except: local_tz=None
        def parse_iso_z(s):
            if not s:return None
            try:return datetime.fromisoformat(s.replace("Z","+00:00")).astimezone(timezone.utc)
            except:return None
        for i,u in enumerate(users,1):
            upn=u.get("userPrincipalName") or u.get("id"); debug(f"Processing {i}/{len(users)}: {upn}")
            try:
                n=normalize_user(u,defs)
                f={k:n.get(k) for k in CORE_FIELDS}
                for k in sel:f[k]=n.get(k)
                results.append(f)
            except Exception as ex:
                debug(f"Error {u.get('id')}: {ex}\n{traceback.format_exc()}"); fb={k:u.get(k) for k in CORE_FIELDS}
                for k in sel: fb[k]=u.get(k); results.append({"user":fb,"error":str(ex)})
        ts=dt.datetime.now().strftime("%Y%m%d_%H%M%S"); out=f"users_output_{ts}.json"
        with open(out,"w",encoding="utf-8") as f: json.dump(results,f,indent=2,ensure_ascii=False)
        debug(f"Saved to {out}")
        print(f"\n Saved {len(results)} users to {out}")
        print("==============================")
        print("    Coded by WizardOrypure    ")
        print("==============================")
    except Exception as e:
        debug(f"Fatal: {e}\n{traceback.format_exc()}"); sys.exit(1)

if __name__=="__main__": main()
