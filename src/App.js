import { useState, useEffect, useCallback } from "react";

const CLIENT_ID = "637cb567-14f5-41ce-a1af-1ff2c6276418";
const TENANT_ID = "common";
const REDIRECT_URI = "https://myday-app-pi.vercel.app";
const SCOPES = "Tasks.ReadWrite User.Read";
const AUTH_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`;
const TOKEN_URL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const GRAPH = "https://graph.microsoft.com/v1.0";

function generateCodeVerifier() {
  const array = new Uint8Array(32);
  window.crypto.getRandomValues(array);
  return btoa(String.fromCharCode(...array)).replace(/\+/g,"-").replace(/\//g,"_").replace(/=/g,"");
}
async function generateCodeChallenge(verifier) {
  const data = new TextEncoder().encode(verifier);
  const digest = await window.crypto.subtle.digest("SHA-256", data);
  return btoa(String.fromCharCode(...new Uint8Array(digest))).replace(/\+/g,"-").replace(/\//g,"_").replace(/=/g,"");
}
async function buildAuthUrl() {
  const verifier = generateCodeVerifier();
  const challenge = await generateCodeChallenge(verifier);
  localStorage.setItem("pkce_verifier", verifier);
  const params = new URLSearchParams({
    client_id: CLIENT_ID, response_type: "code", redirect_uri: REDIRECT_URI,
    scope: SCOPES, response_mode: "query", code_challenge: challenge,
    code_challenge_method: "S256", prompt: "select_account",
  });
  return `${AUTH_URL}?${params}`;
}
async function exchangeCode(code) {
  const verifier = localStorage.getItem("pkce_verifier");
  const body = new URLSearchParams({
    client_id: CLIENT_ID, grant_type: "authorization_code", code,
    redirect_uri: REDIRECT_URI, code_verifier: verifier, scope: SCOPES,
  });
  const res = await fetch(TOKEN_URL, { method:"POST", headers:{"Content-Type":"application/x-www-form-urlencoded"}, body });
  return res.json();
}
async function graphFetch(token, path, method="GET", body=null) {
  const res = await fetch(`${GRAPH}${path}`, {
    method, headers:{ Authorization:`Bearer ${token}`, "Content-Type":"application/json" },
    body: body ? JSON.stringify(body) : null,
  });
  if (res.status === 204) return null;
  return res.json();
}

const importanceMap = { high:"High", normal:"Medium", low:"Low" };
const reverseImportance = { High:"high", Medium:"normal", Low:"low" };
const priorityColors = {
  High:   { bg:"#FEE2E2", text:"#DC2626", dot:"#EF4444" },
  Medium: { bg:"#FEF3C7", text:"#D97706", dot:"#F59E0B" },
  Low:    { bg:"#DCFCE7", text:"#16A34A", dot:"#22C55E" },
};

function ConfettiPop({ onDone }) {
  useEffect(() => { const t = setTimeout(onDone,1200); return ()=>clearTimeout(t); }, [onDone]);
  const pieces = Array.from({length:20},(_,i)=>({
    id:i, x:(Math.random()-.5)*220, y:-(Math.random()*130+40),
    color:["#3B82F6","#10B981","#F59E0B","#8B5CF6","#EC4899","#F97316"][i%6],
    size:Math.random()*9+4, round:Math.random()>.5,
  }));
  return (
    <div style={{position:"absolute",top:"50%",left:"50%",pointerEvents:"none",zIndex:100}}>
      {pieces.map(p=>(
        <div key={p.id} style={{
          position:"absolute",width:p.size,height:p.size,background:p.color,
          borderRadius:p.round?"50%":"2px",left:p.x,top:p.y,
          animation:"cfetti 1.1s ease-out forwards",opacity:0,
          animationDelay:`${Math.random()*.15}s`
        }}/>
      ))}
      <style>{`@keyframes cfetti{0%{opacity:1;transform:translate(0,0) rotate(0deg)}100%{opacity:0;transform:translate(0,30px) rotate(480deg)}}`}</style>
    </div>
  );
}

function LoginScreen({ onLogin }) {
  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0F172A 0%,#1E1B4B 50%,#0F172A 100%)",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif",padding:24}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=DM+Sans:wght@300;400;500;600&display=swap');
      .signin-btn:hover{transform:translateY(-2px);box-shadow:0 16px 48px rgba(59,130,246,0.4)!important;}
      @keyframes floatIn{from{opacity:0;transform:translateY(24px)}to{opacity:1;transform:none}}`}</style>
      <div style={{textAlign:"center",animation:"floatIn 0.6s cubic-bezier(.16,1,.3,1) both"}}>
        <div style={{fontSize:56,marginBottom:20}}>☀️</div>
        <h1 style={{fontFamily:"'Playfair Display',serif",fontSize:42,color:"#F8FAFC",fontWeight:700,letterSpacing:"-1px",marginBottom:12}}>My Day</h1>
        <p style={{color:"#94A3B8",fontSize:16,marginBottom:48,maxWidth:320,margin:"0 auto 48px"}}>Your focused daily task companion, powered by Microsoft To Do</p>
        <button className="signin-btn" onClick={onLogin} style={{display:"inline-flex",alignItems:"center",gap:12,background:"linear-gradient(135deg,#3B82F6,#6366F1)",color:"white",border:"none",borderRadius:16,padding:"16px 32px",fontSize:16,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600,boxShadow:"0 8px 32px rgba(59,130,246,0.25)",transition:"all 0.25s"}}>
          <svg width="20" height="20" viewBox="0 0 21 21" fill="none">
            <rect x="1" y="1" width="9" height="9" fill="#F25022"/>
            <rect x="11" y="1" width="9" height="9" fill="#7FBA00"/>
            <rect x="1" y="11" width="9" height="9" fill="#00A4EF"/>
            <rect x="11" y="11" width="9" height="9" fill="#FFB900"/>
          </svg>
          Sign in with Microsoft
        </button>
        <p style={{color:"#334155",fontSize:12,marginTop:24}}>Syncs with your Microsoft To Do account</p>
      </div>
    </div>
  );
}

function TaskDetailPanel({ task, lists, onClose, onSave, onDelete, getToken }) {
  const [title, setTitle] = useState(task.title);
  const [priority, setPriority] = useState(task.priority);
  const [listId, setListId] = useState(task.listId);
  const [notes, setNotes] = useState(task.msTask?.body?.content || "");
  const [dueDate, setDueDate] = useState(task.msTask?.dueDateTime?.dateTime?.split("T")[0] || "");
  const [addedToDay, setAddedToDay] = useState(task.addedToDay);
  const [subtasks, setSubtasks] = useState([]);
  const [newSubtask, setNewSubtask] = useState("");
  const [saving, setSaving] = useState(false);
  const [loadingSubtasks, setLoadingSubtasks] = useState(true);

  useEffect(() => {
    async function fetchSubtasks() {
      try {
        const token = await getToken();
        const data = await graphFetch(token, `/me/todo/lists/${task.listId}/tasks/${task.id}/checklistItems`);
        setSubtasks(data?.value || []);
      } catch { setSubtasks([]); }
      setLoadingSubtasks(false);
    }
    fetchSubtasks();
  }, [task.id, task.listId, getToken]);

  async function handleSave() {
    setSaving(true);
    try {
      const token = await getToken();
      const body = {
        title,
        importance: reverseImportance[priority] || "normal",
        isReminderOn: addedToDay,
        categories: addedToDay ? ["MyDay"] : [],
        body: { content: notes, contentType: "text" },
      };
      if (dueDate) body.dueDateTime = { dateTime: `${dueDate}T00:00:00`, timeZone: "UTC" };
      else body.dueDateTime = null;
      const updated = await graphFetch(token, `/me/todo/lists/${task.listId}/tasks/${task.id}`, "PATCH", body);
      onSave({ ...task, title, priority, listId, addedToDay, msTask: updated });
    } catch(e) { console.error(e); }
    setSaving(false);
  }

  async function addSubtask() {
    if (!newSubtask.trim()) return;
    const token = await getToken();
    const created = await graphFetch(token, `/me/todo/lists/${task.listId}/tasks/${task.id}/checklistItems`, "POST", {
      displayName: newSubtask.trim(), isChecked: false,
    });
    setSubtasks(prev => [...prev, created]);
    setNewSubtask("");
  }

  async function toggleSubtask(subtask) {
    const token = await getToken();
    await graphFetch(token, `/me/todo/lists/${task.listId}/tasks/${task.id}/checklistItems/${subtask.id}`, "PATCH", { isChecked: !subtask.isChecked });
    setSubtasks(prev => prev.map(s => s.id === subtask.id ? { ...s, isChecked: !s.isChecked } : s));
  }

  async function deleteSubtask(subtask) {
    const token = await getToken();
    await graphFetch(token, `/me/todo/lists/${task.listId}/tasks/${task.id}/checklistItems/${subtask.id}`, "DELETE");
    setSubtasks(prev => prev.filter(s => s.id !== subtask.id));
  }
 
  return (
    <>
      <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.2)",zIndex:40,animation:"fi 0.2s ease"}}/>
      <div style={{position:"fixed",top:0,right:0,bottom:0,width:420,background:"white",zIndex:50,boxShadow:"-8px 0 48px rgba(0,0,0,0.12)",display:"flex",flexDirection:"column",animation:"slideIn 0.3s cubic-bezier(.16,1,.3,1)"}}>
        <style>{`
          @keyframes slideIn{from{transform:translateX(100%)}to{transform:translateX(0)}}
          @keyframes fi{from{opacity:0}to{opacity:1}}
          .subtask-row:hover .subtask-del{opacity:1!important}
        `}</style>
        <div style={{padding:"20px 24px 16px",borderBottom:"1px solid #F1F5F9",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <h2 style={{fontFamily:"'Playfair Display',serif",fontSize:18,color:"#0F172A",fontWeight:700}}>Task Details</h2>
          <button onClick={onClose} style={{background:"#F1F5F9",border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:16,color:"#64748B",display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
        </div>

        <div style={{flex:1,overflowY:"auto",padding:"20px 24px"}}>
          <div style={{marginBottom:18}}>
            <label style={{fontSize:11,fontWeight:600,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em",display:"block",marginBottom:6}}>Title</label>
            <input value={title} onChange={e=>setTitle(e.target.value)}
              style={{width:"100%",fontSize:16,fontWeight:500,padding:"12px 14px",borderRadius:10,border:"1.5px solid #E2E8F0",fontFamily:"'DM Sans',sans-serif",color:"#0F172A",boxSizing:"border-box"}}/>
          </div>

          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:18}}>
            <div>
              <label style={{fontSize:11,fontWeight:600,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em",display:"block",marginBottom:6}}>Priority</label>
              <div style={{display:"flex",gap:6}}>
                {["High","Medium","Low"].map(pv=>(
                  <button key={pv} onClick={()=>setPriority(pv)} style={{flex:1,padding:"7px 4px",borderRadius:8,border:"none",cursor:"pointer",fontSize:11,fontWeight:600,transition:"all 0.15s",background:priority===pv?priorityColors[pv].bg:"#F8FAFC",color:priority===pv?priorityColors[pv].text:"#94A3B8"}}>{pv}</button>
                ))}
              </div>
            </div>
            <div>
              <label style={{fontSize:11,fontWeight:600,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em",display:"block",marginBottom:6}}>My Day</label>
              <button onClick={()=>setAddedToDay(!addedToDay)} style={{width:"100%",padding:"8px",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:500,border:`1.5px solid ${addedToDay?"#BFDBFE":"#E2E8F0"}`,background:addedToDay?"#EFF6FF":"#F8FAFC",color:addedToDay?"#2563EB":"#94A3B8",transition:"all 0.15s"}}>{addedToDay?"☀️ In My Day":"+ My Day"}</button>
            </div>
          </div>

          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:18}}>
            <div>
              <label style={{fontSize:11,fontWeight:600,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em",display:"block",marginBottom:6}}>List</label>
              <select value={listId} onChange={e=>setListId(e.target.value)} style={{width:"100%",padding:"10px 12px",borderRadius:10,fontSize:13,border:"1.5px solid #E2E8F0",fontFamily:"'DM Sans',sans-serif",color:"#1E293B",background:"white",cursor:"pointer",boxSizing:"border-box"}}>
                {lists.map(l=><option key={l.id} value={l.id}>{l.displayName}</option>)}
              </select>
            </div>
            <div>
              <label style={{fontSize:11,fontWeight:600,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em",display:"block",marginBottom:6}}>Due Date</label>
              <input type="date" value={dueDate} onChange={e=>setDueDate(e.target.value)}
                style={{width:"100%",padding:"10px 12px",borderRadius:10,fontSize:13,border:"1.5px solid #E2E8F0",fontFamily:"'DM Sans',sans-serif",color:"#1E293B",background:"white",cursor:"pointer",boxSizing:"border-box"}}/>
            </div>
          </div>

          <div style={{marginBottom:18}}>
            <label style={{fontSize:11,fontWeight:600,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em",display:"block",marginBottom:6}}>Notes</label>
            <textarea value={notes} onChange={e=>setNotes(e.target.value)} placeholder="Add notes…" rows={4}
              style={{width:"100%",padding:"10px 14px",borderRadius:10,fontSize:14,border:"1.5px solid #E2E8F0",fontFamily:"'DM Sans',sans-serif",color:"#1E293B",resize:"vertical",boxSizing:"border-box",background:"white"}}/>
          </div>

          <div>
            <label style={{fontSize:11,fontWeight:600,color:"#94A3B8",textTransform:"uppercase",letterSpacing:"0.06em",display:"block",marginBottom:10}}>Subtasks</label>
            {loadingSubtasks ? <p style={{color:"#94A3B8",fontSize:13}}>Loading…</p> : (
              <>
                {subtasks.map(s=>(
                  <div key={s.id} className="subtask-row" style={{display:"flex",alignItems:"center",gap:10,padding:"8px 10px",borderRadius:8,marginBottom:4,background:"#F8FAFC",border:"1px solid #F1F5F9"}}>
                    <div onClick={()=>toggleSubtask(s)} style={{width:18,height:18,borderRadius:"50%",flexShrink:0,cursor:"pointer",border:`2px solid ${s.isChecked?"#10B981":"#CBD5E1"}`,background:s.isChecked?"#10B981":"transparent",display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,color:"white"}}>{s.isChecked&&"✓"}</div>
                    <span style={{flex:1,fontSize:13,color:s.isChecked?"#94A3B8":"#334155",textDecoration:s.isChecked?"line-through":"none"}}>{s.displayName}</span>
                    <button className="subtask-del" onClick={()=>deleteSubtask(s)} style={{opacity:0,transition:"opacity 0.15s",background:"transparent",border:"none",color:"#FDA4AF",cursor:"pointer",fontSize:13,padding:"2px 6px"}}>✕</button>
                  </div>
                ))}
                <div style={{display:"flex",gap:8,marginTop:8}}>
                  <input value={newSubtask} onChange={e=>setNewSubtask(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addSubtask()} placeholder="Add a subtask…"
                    style={{flex:1,padding:"8px 12px",borderRadius:8,fontSize:13,border:"1.5px solid #E2E8F0",fontFamily:"'DM Sans',sans-serif",color:"#1E293B",background:"white"}}/>
                  <button onClick={addSubtask} style={{padding:"8px 14px",borderRadius:8,border:"none",background:"#0F172A",color:"white",fontSize:13,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}>Add</button>
                </div>
              </>
            )}
          </div>
        </div>

        <div style={{padding:"16px 24px",borderTop:"1px solid #F1F5F9",display:"flex",gap:10}}>
          <button onClick={()=>onDelete(task)} style={{padding:"10px 16px",borderRadius:10,border:"1px solid #FEE2E2",background:"white",color:"#EF4444",fontSize:13,cursor:"pointer",fontFamily:"'DM Sans',sans-serif"}}>🗑 Delete</button>
          <button onClick={handleSave} disabled={saving} style={{flex:1,padding:"10px",borderRadius:10,border:"none",background:"#0F172A",color:"white",fontSize:14,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600,opacity:saving?0.6:1}}>{saving?"Saving…":"Save Changes"}</button>
        </div>
      </div>
    </>
  );
}

export default function App() {
  const [authState, setAuthState] = useState(()=>{ const s=localStorage.getItem("ms_tokens"); return s?JSON.parse(s):null; });
  const [user, setUser] = useState(null);
  const [lists, setLists] = useState([]);
  const [tasks, setTasks] = useState([]);
  const [loading, setLoading] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [error, setError] = useState(null);
  const [activeList, setActiveList] = useState("myday");
  const [view, setView] = useState("planner");
  const [focusTasks, setFocusTasks] = useState([]);
  const [focusIndex, setFocusIndex] = useState(0);
  const [allDone, setAllDone] = useState(false);
  const [confetti, setConfetti] = useState(false);
  const [showAddTask, setShowAddTask] = useState(false);
  const [newTask, setNewTask] = useState({ title:"", listId:"", priority:"Medium" });
  const [selectedTask, setSelectedTask] = useState(null);
  const [dragOverId, setDragOverId] = useState(null);

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const code = params.get("code");
    const err = params.get("error");
    if (err) { setError("Microsoft error: " + params.get("error_description")); return; }
    if (code) {
      window.history.replaceState({}, "", window.location.pathname);
      exchangeCode(code).then(tokens => {
        if (tokens.access_token) { localStorage.setItem("ms_tokens", JSON.stringify(tokens)); setAuthState(tokens); }
        else setError("Token error: " + JSON.stringify(tokens));
      });
    }
  }, []);

  const getToken = useCallback(async () => {
    if (!authState) return null;
    return authState.access_token;
  }, [authState]);

  useEffect(() => {
    if (!authState?.access_token) return;
    loadAll();
  }, [authState]); // eslint-disable-line react-hooks/exhaustive-deps

  async function loadAll() {
    setLoading(true); setError(null);
    try {
      const token = await getToken();
      const me = await graphFetch(token, "/me");
      setUser(me);
      const listsData = await graphFetch(token, "/me/todo/lists");
      const fetchedLists = listsData?.value || [];
      setLists(fetchedLists);
      const allTasks = [];
      for (const list of fetchedLists) {
        const tasksData = await graphFetch(token, `/me/todo/lists/${list.id}/tasks?$top=100`);
        const listTasks = (tasksData?.value || [])
          .filter(t => t.status !== "completed")
          .map(t => ({
            id: t.id, listId: list.id, listName: list.displayName, title: t.title,
            priority: importanceMap[t.importance] || "Medium",
            addedToDay: t.categories?.includes("MyDay") || t.isReminderOn || false,
            done: false, msTask: t,
          }));
        allTasks.push(...listTasks);
      }
      setTasks(allTasks);
      if (fetchedLists.length > 0) setNewTask(prev => ({ ...prev, listId: fetchedLists[0].id }));
    } catch { setError("Failed to load tasks."); }
    setLoading(false);
  }

  async function addTask() {
    if (!newTask.title.trim() || !newTask.listId) return;
    setSyncing(true);
    try {
      const token = await getToken();
      const created = await graphFetch(token, `/me/todo/lists/${newTask.listId}/tasks`, "POST", {
        title: newTask.title.trim(), importance: reverseImportance[newTask.priority] || "normal",
      });
      const list = lists.find(l => l.id === newTask.listId);
      setTasks(prev => [...prev, { id:created.id, listId:newTask.listId, listName:list?.displayName||"", title:created.title, priority:importanceMap[created.importance]||"Medium", addedToDay:false, done:false, msTask:created }]);
      setNewTask(prev => ({ ...prev, title:"" }));
      setShowAddTask(false);
    } catch { setError("Failed to add task."); }
    setSyncing(false);
  }

  async function completeTaskInMs(task) {
    const token = await getToken();
    await graphFetch(token, `/me/todo/lists/${task.listId}/tasks/${task.id}`, "PATCH", { status:"completed" });
    setTasks(prev => prev.filter(t => t.id !== task.id));
  }

  async function toggleMyDay(task) {
    setSyncing(true);
    try {
      const token = await getToken();
      const newVal = !task.addedToDay;
      await graphFetch(token, `/me/todo/lists/${task.listId}/tasks/${task.id}`, "PATCH", {
        isReminderOn: newVal, categories: newVal ? ["MyDay"] : [],
      });
      setTasks(prev => prev.map(t => t.id === task.id ? { ...t, addedToDay: newVal } : t));
    } catch { setError("Sync failed."); }
    setSyncing(false);
  }

  async function deleteTask(task) {
    setSyncing(true);
    try {
      const token = await getToken();
      await graphFetch(token, `/me/todo/lists/${task.listId}/tasks/${task.id}`, "DELETE");
      setTasks(prev => prev.filter(t => t.id !== task.id));
      setSelectedTask(null);
    } catch { setError("Failed to delete task."); }
    setSyncing(false);
  }

  function handleSaveTask(updated) {
    setTasks(prev => prev.map(t => t.id === updated.id ? updated : t));
    setSelectedTask(null);
  }

  const myDayTasks = tasks.filter(t => t.addedToDay && !t.done);
  const sortedMyDay = myDayOrder.length
  ? [...myDayTasks].sort((a,b) => myDayOrder.indexOf(a.id) - myDayOrder.indexOf(b.id))
  : [...myDayTasks].sort((a,b) => ["High","Medium","Low"].indexOf(a.priority) - ["High","Medium","Low"].indexOf(b.priority));

  function startFocus() { setFocusTasks(sortedMyDay); setFocusIndex(0); setAllDone(false); setView("focus"); }

  async function completeFocusTask() {
    const task = focusTasks[focusIndex];
    setConfetti(true);
    await completeTaskInMs(task);
    const next = focusIndex + 1;
    if (next >= focusTasks.length) setTimeout(() => setAllDone(true), 500);
    else setTimeout(() => setFocusIndex(next), 400);
  }

  function skipFocusTask() {
    const next = focusIndex + 1;
    if (next >= focusTasks.length) setAllDone(true); else setFocusIndex(next);
  }

  const displayedTasks = activeList === "myday" ? tasks.filter(t=>t.addedToDay)
    : activeList === "all" ? tasks
    : tasks.filter(t=>t.listId===activeList);

  if (!authState?.access_token && !new URLSearchParams(window.location.search).get("code")) {
    return <LoginScreen onLogin={async()=>{ window.location.href = await buildAuthUrl(); }}/>;
  }

  if (loading) return (
    <div style={{minHeight:"100vh",background:"#0F172A",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif",color:"#94A3B8"}}>
      <div style={{fontSize:48,marginBottom:16,animation:"spin 1.5s linear infinite"}}>☀️</div>
      <p style={{fontSize:16}}>Loading your tasks from Microsoft To Do…</p>
      <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
    </div>
  );

  if (view === "focus") {
    const cur = focusTasks[focusIndex];
    const focusPriority = cur ? priorityColors[cur.priority] : null;
    return (
      <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0F172A 0%,#1E293B 60%,#0F172A 100%)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif",padding:24}}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=DM+Sans:wght@300;400;500;600&display=swap');
          .fc{animation:su 0.5s cubic-bezier(.16,1,.3,1) both}
          @keyframes su{from{opacity:0;transform:translateY(28px) scale(0.97)}to{opacity:1;transform:none}}
          .cbtn:hover{transform:scale(1.05);box-shadow:0 0 48px rgba(16,185,129,0.5)!important}
          .sbtn:hover{opacity:0.7}`}</style>
        <div style={{width:"100%",maxWidth:520,marginBottom:44}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:8,color:"#475569",fontSize:13}}>
            <span style={{color:"#64748B"}}>Task {Math.min(focusIndex+1,focusTasks.length)} of {focusTasks.length}</span>
            <span style={{cursor:"pointer",color:"#475569"}} onClick={()=>setView("planner")}>← Back to Planner</span>
          </div>
          <div style={{height:3,background:"rgba(255,255,255,0.06)",borderRadius:4,overflow:"hidden"}}>
            <div style={{height:"100%",borderRadius:4,background:"linear-gradient(90deg,#3B82F6,#10B981)",width:`${(focusIndex/focusTasks.length)*100}%`,transition:"width 0.5s ease"}}/>
          </div>
        </div>
        {allDone ? (
          <div style={{textAlign:"center",animation:"su 0.5s both"}}>
            <div style={{fontSize:72,marginBottom:16}}>🎉</div>
            <h1 style={{fontFamily:"'Playfair Display',serif",color:"#F8FAFC",fontSize:40,marginBottom:12,fontWeight:700}}>All done!</h1>
            <p style={{color:"#94A3B8",fontSize:18,marginBottom:40}}>Your tasks are synced back to Microsoft To Do.</p>
            <button onClick={()=>{setView("planner");loadAll();}} style={{background:"linear-gradient(135deg,#3B82F6,#6366F1)",color:"white",border:"none",borderRadius:14,padding:"14px 36px",fontSize:16,fontFamily:"'DM Sans',sans-serif",cursor:"pointer",fontWeight:600}}>Back to Planner</button>
          </div>
        ) : (
          <div className="fc" key={focusIndex} style={{background:"rgba(255,255,255,0.04)",backdropFilter:"blur(12px)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:28,padding:"52px 48px",maxWidth:520,width:"100%",textAlign:"center",position:"relative"}}>
            {confetti && <ConfettiPop onDone={()=>setConfetti(false)}/>}
            <div style={{display:"inline-flex",alignItems:"center",gap:8,marginBottom:32,background:"rgba(255,255,255,0.06)",borderRadius:100,padding:"6px 16px"}}>
              <span style={{color:"#94A3B8",fontSize:13}}>{cur?.listName}</span>
              <span style={{background:focusPriority?.bg,color:focusPriority?.text,borderRadius:100,padding:"2px 10px",fontSize:11,fontWeight:600}}>{cur?.priority}</span>
            </div>
            <h1 style={{fontFamily:"'Playfair Display',serif",color:"#F8FAFC",fontSize:32,lineHeight:1.35,marginBottom:52,fontWeight:700,letterSpacing:"-0.5px"}}>{cur?.title}</h1>
            <button className="cbtn" onClick={completeFocusTask} style={{width:72,height:72,borderRadius:"50%",background:"linear-gradient(135deg,#10B981,#059669)",border:"none",fontSize:30,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 20px",transition:"all 0.2s",boxShadow:"0 8px 32px rgba(16,185,129,0.2)"}}>✓</button>
            <p style={{color:"#475569",fontSize:13,marginBottom:24}}>Tap to complete & sync</p>
            {focusIndex < focusTasks.length-1 && (
              <button className="sbtn" onClick={skipFocusTask} style={{background:"transparent",border:"1px solid rgba(255,255,255,0.1)",color:"#475569",borderRadius:10,padding:"8px 20px",fontSize:13,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",transition:"opacity 0.2s"}}>Skip for now →</button>
            )}
          </div>
        )}
      </div>
    );
  }

  return (
    <div style={{minHeight:"100vh",background:"#F8FAFC",display:"flex",fontFamily:"'DM Sans',sans-serif"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=DM+Sans:wght@300;400;500;600&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        .scat:hover{background:#F1F5F9!important}
        .trow:hover .tactions{opacity:1!important}
        .trow:hover{background:#F8FAFC!important}
        .trow{cursor:pointer}
        .addbtn:hover{background:#E0F2FE!important}
        .stbtn:hover{transform:translateY(-1px);box-shadow:0 12px 40px rgba(59,130,246,0.35)!important}
        .mo{animation:fi 0.15s ease}
        @keyframes fi{from{opacity:0}to{opacity:1}}
        input:focus,select:focus,textarea:focus{outline:2px solid #BFDBFE!important}`}</style>

      <div style={{width:248,background:"white",borderRight:"1px solid #E2E8F0",display:"flex",flexDirection:"column",padding:"0 0 24px",position:"sticky",top:0,height:"100vh",flexShrink:0}}>
        <div style={{padding:"28px 20px 16px"}}>
          <div style={{fontFamily:"'Playfair Display',serif",fontSize:22,color:"#0F172A",fontWeight:700,letterSpacing:"-0.5px"}}>My Day</div>
          {user && (
            <div style={{fontSize:12,color:"#94A3B8",marginTop:4,display:"flex",alignItems:"center",gap:6}}>
              <div style={{width:20,height:20,borderRadius:"50%",background:"linear-gradient(135deg,#3B82F6,#6366F1)",display:"flex",alignItems:"center",justifyContent:"center",color:"white",fontSize:10,fontWeight:700,flexShrink:0}}>{user.displayName?.[0]?.toUpperCase()||"U"}</div>
              <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{user.displayName||user.mail}</span>
            </div>
          )}
        </div>
        <div style={{padding:"0 12px 4px"}}>
          {[{id:"myday",label:"☀️ My Day",count:myDayTasks.length},{id:"all",label:"📋 All Tasks",count:tasks.length}].map(item=>(
            <button key={item.id} className="scat" onClick={()=>setActiveList(item.id)} style={{width:"100%",textAlign:"left",padding:"10px 12px",borderRadius:10,border:"none",cursor:"pointer",fontSize:14,fontWeight:500,background:activeList===item.id?"#EFF6FF":"transparent",color:activeList===item.id?"#2563EB":"#475569",fontFamily:"'DM Sans',sans-serif",transition:"background 0.15s",display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:2}}>
              <span>{item.label}</span>
              {item.count>0&&<span style={{background:activeList===item.id?"#BFDBFE":"#F1F5F9",color:activeList===item.id?"#2563EB":"#94A3B8",borderRadius:100,padding:"1px 8px",fontSize:11,fontWeight:600}}>{item.count}</span>}
            </button>
          ))}
        </div>
        <div style={{padding:"12px 20px 8px",fontSize:11,color:"#94A3B8",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.08em"}}>Lists</div>
        <div style={{padding:"0 12px",flex:1,overflowY:"auto"}}>
          {lists.map(list=>{
            const count=tasks.filter(t=>t.listId===list.id).length;
            return (
              <button key={list.id} className="scat" onClick={()=>setActiveList(list.id)} style={{width:"100%",textAlign:"left",padding:"10px 12px",borderRadius:10,border:"none",cursor:"pointer",fontSize:14,background:activeList===list.id?"#EFF6FF":"transparent",color:activeList===list.id?"#2563EB":"#475569",fontFamily:"'DM Sans',sans-serif",transition:"background 0.15s",display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:2}}>
                <span>📝 {list.displayName}</span>
                {count>0&&<span style={{background:activeList===list.id?"#BFDBFE":"#F1F5F9",color:activeList===list.id?"#2563EB":"#94A3B8",borderRadius:100,padding:"1px 8px",fontSize:11,fontWeight:600}}>{count}</span>}
              </button>
            );
          })}
        </div>
        <div style={{padding:"12px 12px 0"}}>
          <button onClick={()=>{localStorage.removeItem("ms_tokens");setAuthState(null);}} style={{width:"100%",padding:"9px 12px",borderRadius:10,border:"1px solid #E2E8F0",background:"transparent",color:"#94A3B8",fontSize:13,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",textAlign:"left"}}>↩ Sign out</button>
        </div>
      </div>

      <div style={{flex:1,display:"flex",flexDirection:"column"}}>
        <div style={{padding:"28px 36px 0",display:"flex",alignItems:"flex-start",justifyContent:"space-between"}}>
          <div>
            <h1 style={{fontFamily:"'Playfair Display',serif",fontSize:30,color:"#0F172A",fontWeight:700,letterSpacing:"-0.5px"}}>
              {activeList==="myday"?"My Day":activeList==="all"?"All Tasks":lists.find(l=>l.id===activeList)?.displayName}
            </h1>
            <p style={{color:"#94A3B8",fontSize:14,marginTop:4}}>
              {new Date().toLocaleDateString("en-US",{weekday:"long",month:"long",day:"numeric"})}
              {syncing&&<span style={{marginLeft:10,color:"#3B82F6"}}>↻ Syncing…</span>}
            </p>
          </div>
          <div style={{display:"flex",gap:10}}>
            <button onClick={loadAll} style={{background:"white",border:"1px solid #E2E8F0",borderRadius:12,padding:"10px 16px",fontSize:13,cursor:"pointer",color:"#64748B",fontFamily:"'DM Sans',sans-serif"}}>↻ Refresh</button>
            <button onClick={()=>setShowAddTask(true)} style={{background:"#0F172A",color:"white",border:"none",borderRadius:12,padding:"10px 20px",fontSize:14,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600}}>+ Add Task</button>
          </div>
        </div>

        {error&&<div style={{margin:"16px 36px 0",padding:"12px 16px",background:"#FEE2E2",borderRadius:10,color:"#DC2626",fontSize:13}}>⚠️ {error} <span style={{cursor:"pointer",marginLeft:8,textDecoration:"underline"}} onClick={()=>setError(null)}>Dismiss</span></div>}

        {myDayTasks.length>0&&(
          <div style={{padding:"20px 36px 0"}}>
            <div style={{background:"white",borderRadius:20,border:"1px solid #E2E8F0",padding:"20px 24px",boxShadow:"0 2px 12px rgba(0,0,0,0.04)"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:18}}>☀️</span>
                  <span style={{fontSize:15,fontWeight:600,color:"#0F172A"}}>Ready to focus</span>
                  <span style={{background:"#EFF6FF",color:"#3B82F6",borderRadius:100,padding:"1px 10px",fontSize:12,fontWeight:600}}>{myDayTasks.length}</span>
                </div>
                <button className="stbtn" onClick={startFocus} style={{background:"linear-gradient(135deg,#3B82F6,#6366F1)",color:"white",border:"none",borderRadius:10,padding:"9px 22px",fontSize:14,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600,transition:"all 0.2s",boxShadow:"0 4px 20px rgba(59,130,246,0.2)"}}>▶ Start</button>
              </div>
              <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                {sortedMyDay.map((task,i)=>{
                  const p=priorityColors[task.priority];
                  return (
                    <div key={task.id}
                      draggable
                      onDragStart={()=>setDragOverId(task.id)}
                      onDragOver={e=>{
                        e.preventDefault();
                        if(dragOverId && dragOverId!==task.id){
                          const ids=sortedMyDay.map(t=>t.id);
                          const from=ids.indexOf(dragOverId);
                          const to=ids.indexOf(task.id);
                          const reordered=[...ids];
                          reordered.splice(from,1);
                          reordered.splice(to,0,dragOverId);
                          setMyDayOrder(reordered);
                        }
                      }}
                      onDragEnd={()=>setDragOverId(null)}
                      onClick={()=>setSelectedTask(task)}
                      style={{display:"flex",alignItems:"center",gap:6,background:dragOverId===task.id?"#EFF6FF":"#F8FAFC",borderRadius:8,padding:"6px 12px",fontSize:13,border:`1px solid ${dragOverId===task.id?"#BFDBFE":"#E2E8F0"}`,cursor:"grab",userSelect:"none",transition:"all 0.15s"}}>
                      <span style={{color:"#94A3B8",fontSize:11,fontWeight:600}}>{i+1}</span>
                      <span style={{fontSize:10,color:"#CBD5E1"}}>⠿</span>
                      <span style={{width:6,height:6,borderRadius:"50%",background:p.dot,flexShrink:0}}/>
                      <span style={{color:"#334155"}}>{task.title}</span>
                    </div>
                  );
                })}
              </div>

        <div style={{padding:"20px 36px 36px"}}>
          <div style={{background:"white",borderRadius:20,border:"1px solid #E2E8F0",overflow:"hidden",boxShadow:"0 2px 12px rgba(0,0,0,0.04)"}}>
            {displayedTasks.length===0 ? (
              <div style={{padding:"52px 24px",textAlign:"center",color:"#94A3B8"}}>
                <div style={{fontSize:36,marginBottom:12}}>✨</div>
                <p style={{fontSize:15}}>{activeList==="myday"?"No tasks in My Day — hover a task and click '+ My Day'":"No tasks here yet."}</p>
              </div>
            ) : displayedTasks.map((task,i)=>{
              const p=priorityColors[task.priority];
              return (
                <div key={task.id} className="trow" onClick={()=>setSelectedTask(task)} style={{display:"flex",alignItems:"center",gap:14,padding:"14px 20px",borderBottom:i<displayedTasks.length-1?"1px solid #F1F5F9":"none",transition:"background 0.15s",background:"white"}}>
                  <div onClick={e=>{e.stopPropagation();completeTaskInMs(task);}} style={{width:20,height:20,borderRadius:"50%",border:"2px solid #CBD5E1",background:"transparent",flexShrink:0,cursor:"pointer",transition:"all 0.15s"}}/>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:14,color:"#1E293B",fontWeight:450}}>{task.title}</div>
                    <div style={{fontSize:12,color:"#94A3B8",marginTop:2,display:"flex",gap:8}}>
                      <span>{task.listName}</span>
                      {task.msTask?.dueDateTime&&<span style={{color:"#F59E0B"}}>📅 {new Date(task.msTask.dueDateTime.dateTime).toLocaleDateString()}</span>}
                    </div>
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
                    <span style={{fontSize:11,fontWeight:600,background:p.bg,color:p.text,borderRadius:100,padding:"2px 10px"}}>{task.priority}</span>
                    <div className="tactions" style={{display:"flex",gap:4,opacity:0,transition:"opacity 0.15s"}}>
                      <button className="addbtn" onClick={e=>{e.stopPropagation();toggleMyDay(task);}} style={{background:task.addedToDay?"#EFF6FF":"transparent",color:task.addedToDay?"#2563EB":"#94A3B8",border:`1px solid ${task.addedToDay?"#BFDBFE":"#E2E8F0"}`,borderRadius:7,padding:"4px 10px",fontSize:11,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",transition:"background 0.15s"}}>
                        {task.addedToDay?"☀️ In My Day":"+ My Day"}
                      </button>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {selectedTask&&(
        <TaskDetailPanel task={selectedTask} lists={lists} onClose={()=>setSelectedTask(null)} onSave={handleSaveTask} onDelete={deleteTask} getToken={getToken}/>
      )}

      {showAddTask&&(
        <div className="mo" onClick={()=>setShowAddTask(false)} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.3)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:50}}>
          <div onClick={e=>e.stopPropagation()} style={{background:"white",borderRadius:20,padding:32,width:400,boxShadow:"0 24px 80px rgba(0,0,0,0.15)"}}>
            <h2 style={{fontFamily:"'Playfair Display',serif",fontSize:22,color:"#0F172A",marginBottom:24}}>New Task</h2>
            <div style={{marginBottom:16}}>
              <label style={{fontSize:12,fontWeight:600,color:"#64748B",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:6}}>Task name</label>
              <input value={newTask.title} onChange={e=>setNewTask({...newTask,title:e.target.value})} onKeyDown={e=>e.key==="Enter"&&addTask()} placeholder="What needs to be done?" autoFocus style={{width:"100%",padding:"10px 14px",borderRadius:10,fontSize:14,border:"1.5px solid #E2E8F0",fontFamily:"'DM Sans',sans-serif",color:"#1E293B"}}/>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:24}}>
              <div>
                <label style={{fontSize:12,fontWeight:600,color:"#64748B",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:6}}>List</label>
                <select value={newTask.listId} onChange={e=>setNewTask({...newTask,listId:e.target.value})} style={{width:"100%",padding:"10px 14px",borderRadius:10,fontSize:14,border:"1.5px solid #E2E8F0",fontFamily:"'DM Sans',sans-serif",color:"#1E293B",background:"white",cursor:"pointer"}}>
                  {lists.map(l=><option key={l.id} value={l.id}>{l.displayName}</option>)}
                </select>
              </div>
              <div>
                <label style={{fontSize:12,fontWeight:600,color:"#64748B",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:6}}>Priority</label>
                <select value={newTask.priority} onChange={e=>setNewTask({...newTask,priority:e.target.value})} style={{width:"100%",padding:"10px 14px",borderRadius:10,fontSize:14,border:"1.5px solid #E2E8F0",fontFamily:"'DM Sans',sans-serif",color:"#1E293B",background:"white",cursor:"pointer"}}>
                  {["High","Medium","Low"].map(p=><option key={p} value={p}>{p}</option>)}
                </select>
              </div>
            </div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setShowAddTask(false)} style={{flex:1,padding:"11px",borderRadius:10,border:"1.5px solid #E2E8F0",background:"white",color:"#64748B",fontSize:14,cursor:"pointer",fontFamily:"'DM Sans',sans-serif"}}>Cancel</button>
              <button onClick={addTask} disabled={syncing} style={{flex:2,padding:"11px",borderRadius:10,border:"none",background:"#0F172A",color:"white",fontSize:14,cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600,opacity:syncing?0.6:1}}>{syncing?"Adding…":"Add Task"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
