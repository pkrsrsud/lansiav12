"use client";

import { FormEvent, useEffect, useMemo, useRef, useState } from "react";
import { api } from "@/lib/api";
import type { AdminData, Attendance, EventRow, Participant } from "@/lib/types";

const TZ = "Asia/Jakarta";

function formatDate(value: string) {
  const raw = String(value || "").trim();
  if (!raw) return "-";
  const datePart = raw.match(/\d{4}-\d{2}-\d{2}/)?.[0] || raw.slice(0, 10);
  const d = new Date(`${datePart}T00:00:00+07:00`);
  if (Number.isNaN(d.getTime())) return "-";
  return new Intl.DateTimeFormat("id-ID", { day: "numeric", month: "long", year: "numeric", timeZone: TZ }).format(d);
}

function toDateTimeLocal(value: string) {
  if (!value) return "";
  const d = new Date(value);
  return new Intl.DateTimeFormat("en-CA", { timeZone: TZ, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false }).format(d).replace(", ", "T");
}

function eventStatus(event: EventRow): [string, "yellow" | "green" | "gray"] {
  const now = Date.now();
  const open = new Date(event.open_at).getTime();
  const close = new Date(event.close_at).getTime();
  if (now < open) return ["Belum dibuka", "yellow"];
  if (now >= close) return ["Selesai", "gray"];
  return ["Sedang dibuka", "green"];
}

const blankEvent = { id: "", name: "", event_date: "", event_time: "", location: "", open_at: "", close_at: "", is_active: true };
const blankParticipant = { id: "", full_name: "", address: "" };

type DashboardProps = { token: string; onLogout: () => void };

function Stat({ label, value }: { label: string; value: string | number }) {
  return <div className="stat"><div className="stat-label">{label}</div><div className="stat-value">{value}</div></div>;
}

export default function AdminPage() {
  const [token, setToken] = useState("");
  const [checking, setChecking] = useState(true);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [tab, setTab] = useState("overview");

  useEffect(() => {
    const saved = window.localStorage.getItem("lansia_admin_token") || "";
    setToken(saved);
    setChecking(false);
  }, []);

  async function login(e: FormEvent) {
    e.preventDefault();
    setLoginError("");
    try {
      const data = await api<{ token: string }>("adminLogin", { email, password });
      window.localStorage.setItem("lansia_admin_token", data.token);
      setToken(data.token);
    } catch (err) {
      setLoginError(err instanceof Error ? err.message : "Login gagal.");
    }
  }

  function logout() {
    window.localStorage.removeItem("lansia_admin_token");
    setToken("");
  }

  if (checking) return <div className="login-screen"><div className="login-box"><h2>Memuat…</h2></div></div>;
  if (!token) return <div className="login-screen"><form className="login-box" onSubmit={login}><div className="eyebrow">LANSIA SMART</div><h1 className="admin-title">Login Admin</h1><p className="muted">Kelola kegiatan, peserta, kehadiran, dan laporan.</p><div className="field"><label>Email</label><input value={email} onChange={(e) => setEmail(e.target.value)} type="email" required /></div><div className="field" style={{marginTop:12}}><label>Password</label><input value={password} onChange={(e) => setPassword(e.target.value)} type="password" required /></div>{loginError && <div className="error">{loginError}</div>}<button className="btn btn-primary" style={{width:"100%",marginTop:16}}>Masuk</button><a className="back-link" href="/">← Kembali ke halaman peserta</a></form></div>;

  return <Dashboard token={token} onLogout={logout} />;
}

function Dashboard({ token, onLogout }: DashboardProps) {
  const [data, setData] = useState<AdminData>({ events: [], participants: [], attendance: [] });
  const [selectedEventId, setSelectedEventId] = useState("");
  const [tab, setTab] = useState("overview");
  const [eventForm, setEventForm] = useState(blankEvent);
  const [participantForm, setParticipantForm] = useState(blankParticipant);
  const [participantQuery, setParticipantQuery] = useState("");
  const [toast, setToast] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [importUrl, setImportUrl] = useState("");
  const [importSourceType, setImportSourceType] = useState<"auto" | "spreadsheet" | "docs">("auto");
  const [importUpdateAddress, setImportUpdateAddress] = useState(true);
  const [importReactivate, setImportReactivate] = useState(true);
  const [importBusy, setImportBusy] = useState(false);

  function normalizeAdminData(value: unknown): AdminData {
    const candidate = (value && typeof value === "object" && "data" in value && typeof (value as { data?: unknown }).data === "object")
      ? (value as { data: unknown }).data
      : value;
    const src = (candidate && typeof candidate === "object") ? candidate as Partial<AdminData> : {};
    if (!Array.isArray(src.events) || !Array.isArray(src.participants) || !Array.isArray(src.attendance)) {
      throw new Error("Respons dashboard admin tidak lengkap. Pastikan Apps Script yang dipakai sudah menggunakan Code.gs versi terbaru dan deployment Web App sudah diperbarui.");
    }
    return { events: src.events, participants: src.participants, attendance: src.attendance };
  }

  const loadData = async (eventId = selectedEventId) => {
    setError("");
    const raw = await api<unknown>("adminLoad", { event_id: eventId }, token);
    const next = normalizeAdminData(raw);
    setData(next);
    if (!eventId && next.events[0]) setSelectedEventId(next.events[0].id);
  };

  const initialLoadDone = useRef(false);

  useEffect(() => {
    loadData().then(() => { initialLoadDone.current = true; }).catch((e) => setError(e instanceof Error ? e.message : "Gagal memuat dashboard."));
  }, []);

  useEffect(() => {
    if (!selectedEventId || !initialLoadDone.current) return;
    loadData(selectedEventId).catch((e) => setError(e instanceof Error ? e.message : "Gagal memuat data."));
  }, [selectedEventId]);

  const selectedEvent = data.events.find((e) => e.id === selectedEventId) || data.events[0] || null;
  const activeParticipants = data.participants.filter((p) => p.is_active);
  const attendanceCount = selectedEvent ? data.attendance.length : 0;
  const attendanceRate = activeParticipants.length ? Math.round(attendanceCount / activeParticipants.length * 100) : 0;
  const filteredParticipants = useMemo(() => {
    const q = participantQuery.trim().toLowerCase();
    if (!q) return data.participants;
    return data.participants.filter((p) => p.full_name.toLowerCase().includes(q) || p.address.toLowerCase().includes(q));
  }, [data.participants, participantQuery]);

  function showToast(message: string) { setToast(message); window.setTimeout(() => setToast(""), 2600); }

  async function saveParticipant(e: FormEvent) {
    e.preventDefault();
    setBusy(true);
    try {
      await api("saveParticipant", participantForm, token);
      setParticipantForm(blankParticipant);
      await loadData(selectedEventId);
      showToast("Peserta disimpan.");
    } catch (err) { showToast(err instanceof Error ? err.message : "Gagal menyimpan peserta."); }
    finally { setBusy(false); }
  }

  async function toggleParticipant(p: Participant) {
    try { await api("toggleParticipant", { id: p.id, is_active: !p.is_active }, token); await loadData(selectedEventId); showToast(p.is_active ? "Peserta dinonaktifkan." : "Peserta diaktifkan."); }
    catch (err) { showToast(err instanceof Error ? err.message : "Gagal mengubah peserta."); }
  }

  async function importParticipants(e: FormEvent) {
    e.preventDefault();
    if (!importUrl.trim()) { showToast("Tempel URL Google Spreadsheet atau Google Docs."); return; }
    setImportBusy(true);
    try {
      const result = await api<{ total:number; inserted:number; updated:number; unchanged:number; skipped:number; source_type:string; source_name:string }>(
        "importParticipants",
        { url: importUrl.trim(), source_type: importSourceType, update_address: importUpdateAddress, reactivate_existing: importReactivate },
        token
      );
      await loadData(selectedEventId);
      const sourceLabel = result.source_type === "spreadsheet" ? "Spreadsheet" : "Google Docs";
      showToast(`${sourceLabel}: ${result.inserted} baru, ${result.updated} diperbarui, ${result.unchanged} tetap, ${result.skipped} dilewati.`);
      setImportUrl("");
    } catch (err) {
      showToast(err instanceof Error ? err.message : "Gagal mengimport peserta.");
    } finally {
      setImportBusy(false);
    }
  }

  async function saveEvent(e: FormEvent) {
    e.preventDefault();
    if (!eventForm.name || !eventForm.event_date || !eventForm.location || !eventForm.open_at || !eventForm.close_at) { showToast("Lengkapi data kegiatan."); return; }
    if (new Date(eventForm.close_at).getTime() <= new Date(eventForm.open_at).getTime()) { showToast("Waktu tutup pendaftaran harus setelah waktu buka."); return; }
    setBusy(true);
    try {
      const saved = await api<{ event: EventRow }>("saveEvent", eventForm, token);
      setEventForm(blankEvent);
      await loadData(saved.event.id);
      setSelectedEventId(saved.event.id);
      setTab("events");
      showToast("Kegiatan disimpan.");
    } catch (err) { showToast(err instanceof Error ? err.message : "Gagal menyimpan kegiatan."); }
    finally { setBusy(false); }
  }

  function editEvent(e: EventRow) {
    setEventForm({ id: e.id, name: e.name, event_date: e.event_date.slice(0,10), event_time: e.event_time, location: e.location, open_at: toDateTimeLocal(e.open_at), close_at: toDateTimeLocal(e.close_at), is_active: e.is_active });
    setTab("events");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  async function toggleEvent(e: EventRow) {
    try { await api("toggleEvent", { id: e.id, is_active: !e.is_active }, token); await loadData(selectedEventId); showToast(e.is_active ? "Kegiatan ditutup." : "Kegiatan diaktifkan."); }
    catch (err) { showToast(err instanceof Error ? err.message : "Gagal mengubah kegiatan."); }
  }

  async function deleteEvent(e: EventRow) {
    if (!window.confirm(`Hapus kegiatan "${e.name}"?\n\nSemua data kehadiran untuk kegiatan ini juga akan dihapus.`)) return;
    try { await api("deleteEvent", { id: e.id }, token); const remaining = data.events.filter(x => x.id !== e.id); setSelectedEventId(remaining[0]?.id || ""); await loadData(remaining[0]?.id || ""); showToast("Kegiatan dihapus."); }
    catch (err) { showToast(err instanceof Error ? err.message : "Gagal menghapus kegiatan."); }
  }

  async function removeAttendance(a: Attendance) {
    if (!window.confirm(`Hapus ${a.participant?.full_name || "peserta"} dari daftar hadir?`)) return;
    try { await api("removeAttendance", { id: a.id }, token); await loadData(selectedEventId); showToast("Data kehadiran dihapus."); }
    catch (err) { showToast(err instanceof Error ? err.message : "Gagal menghapus kehadiran."); }
  }

  async function exportReport(format: "doc" | "pdf") {
    if (!selectedEvent) return;

    // Buka tab baru SEBELUM request async supaya browser tidak memblokir popup.
    const newTab = window.open("about:blank", "_blank");

    try {
      const result = await api<{ url: string; docUrl?: string; pdfUrl?: string }>(
        "exportReport",
        { event_id: selectedEvent.id, format },
        token
      );

      const targetUrl = format === "doc" ? (result.docUrl || result.url) : (result.pdfUrl || result.url);

      if (newTab && !newTab.closed) {
        newTab.location.href = targetUrl;
        newTab.focus();
      } else {
        // Fallback kalau browser tetap memblokir pembukaan tab.
        window.location.href = targetUrl;
      }

      showToast(format === "doc" ? "Google Docs berhasil dibuat." : "PDF berhasil dibuat.");
    } catch (err) {
      if (newTab && !newTab.closed) newTab.close();
      showToast(err instanceof Error ? err.message : "Gagal membuat laporan.");
    }
  }

  const [statusText, statusColor] = selectedEvent ? eventStatus(selectedEvent) : ["-", "gray"];

  return <>
    <header className="topbar"><div className="topbar-inner"><div><div className="brand">Lansia SMART RSUD Wonosari</div><div className="subbrand">Dashboard Admin</div></div><div className="top-actions"><a className="admin-link" href="/">Halaman Peserta</a><button className="btn btn-secondary" onClick={onLogout}>Keluar</button></div></div></header>
    <main className="admin-wrap">
      <div className="admin-head"><div><h1 className="admin-title">Dashboard</h1><div className="admin-sub">Kelola kegiatan, peserta, kehadiran, dan laporan.</div></div></div>
      {error && <div className="error">{error}</div>}
      <div className="tabs">
        {[['overview','Ringkasan'],['events','Kegiatan'],['participants','Peserta'],['attendance','Kehadiran'],['reports','Laporan']].map(([key,label]) => <button key={key} className={`tab ${tab===key?'active':''}`} onClick={()=>setTab(key)}>{label}</button>)}
      </div>

      {tab === "overview" && <section>
        <div className="panel"><div className="toolbar"><div><h3>Kegiatan yang dipantau</h3><div className="muted">{selectedEvent ? `${selectedEvent.name} • ${formatDate(selectedEvent.event_date)}` : "Belum ada kegiatan"}</div></div><select value={selectedEvent?.id || ""} onChange={(e)=>setSelectedEventId(e.target.value)}><option value="">Pilih kegiatan</option>{data.events.map(e=><option key={e.id} value={e.id}>{e.name} • {formatDate(e.event_date)}</option>)}</select></div></div>
        <div className="grid4"><Stat label="Total peserta aktif" value={activeParticipants.length}/><Stat label="Hadir" value={attendanceCount}/><Stat label="Belum hadir" value={Math.max(activeParticipants.length-attendanceCount,0)}/><Stat label="Persentase" value={`${attendanceRate}%`}/></div>
        <div className="panel"><div className="toolbar"><div><h3>Status kegiatan</h3><div className="muted">{selectedEvent ? `${statusText}${selectedEvent.is_active ? " • Publik aktif" : " • Publik nonaktif"}` : "-"}</div></div>{selectedEvent && <button className="btn btn-secondary" onClick={()=>toggleEvent(selectedEvent)}>{selectedEvent.is_active?"Tutup Publik":"Aktifkan"}</button>}</div></div>
      </section>}

      {tab === "events" && <section>
        <div className="panel"><div className="toolbar"><div><h3>{eventForm.id ? "Edit kegiatan" : "Tambah kegiatan"}</h3><div className="muted">Waktu kegiatan dan waktu pendaftaran dipisahkan.</div></div></div><form onSubmit={saveEvent} className="form-grid" style={{marginTop:16}}>
          <div className="field"><label>Nama kegiatan</label><input value={eventForm.name} onChange={e=>setEventForm({...eventForm,name:e.target.value})} placeholder="Pertemuan Lansia"/></div>
          <div className="field"><label>Tanggal kegiatan</label><input type="date" value={eventForm.event_date} onChange={e=>setEventForm({...eventForm,event_date:e.target.value})}/></div>
          <div className="field"><label>Waktu kegiatan</label><input value={eventForm.event_time} onChange={e=>setEventForm({...eventForm,event_time:e.target.value})} placeholder="07.00 WIB - Selesai"/></div>
          <div className="field"><label>Tempat</label><input value={eventForm.location} onChange={e=>setEventForm({...eventForm,location:e.target.value})} placeholder="Ruang ..."/></div>
          <div className="field"><label>Buka pendaftaran</label><input type="datetime-local" value={eventForm.open_at} onChange={e=>setEventForm({...eventForm,open_at:e.target.value})}/></div>
          <div className="field"><label>Tutup pendaftaran</label><input type="datetime-local" value={eventForm.close_at} onChange={e=>setEventForm({...eventForm,close_at:e.target.value})}/></div>
          <label style={{display:"flex",alignItems:"center",gap:8,fontWeight:800}}><input type="checkbox" checked={eventForm.is_active} onChange={e=>setEventForm({...eventForm,is_active:e.target.checked})}/> Jadikan kegiatan aktif</label>
          <div className="actions" style={{alignItems:"flex-end"}}><button className="btn btn-primary" disabled={busy}>Simpan</button><button type="button" className="btn btn-secondary" onClick={()=>setEventForm(blankEvent)}>Bersihkan</button></div>
        </form></div>
        <div className="panel"><h3>Daftar kegiatan</h3><div className="table-wrap"><table className="table"><thead><tr><th>Nama</th><th>Tanggal</th><th>Waktu</th><th>Pendaftaran</th><th>Status</th><th>Aksi</th></tr></thead><tbody>{data.events.map(e=>{const [s,c]=eventStatus(e);return <tr key={e.id}><td><b>{e.name}</b><div className="muted">{e.location}</div></td><td>{formatDate(e.event_date)}</td><td>{e.event_time||"-"}</td><td>{new Date(e.open_at).toLocaleString("id-ID",{timeZone:TZ})}<br/>s/d<br/>{new Date(e.close_at).toLocaleString("id-ID",{timeZone:TZ})}</td><td><span className={`pill pill-${c}`}>{s}</span>{e.is_active&&<span className="pill pill-green" style={{marginLeft:6}}>Aktif</span>}</td><td><div className="actions"><button className="btn btn-secondary btn-small" onClick={()=>editEvent(e)}>Edit</button><button className="btn btn-secondary btn-small" onClick={()=>toggleEvent(e)}>{e.is_active?"Tutup":"Aktifkan"}</button><button className="btn btn-danger btn-small" onClick={()=>deleteEvent(e)}>Hapus</button></div></td></tr>})}</tbody></table></div></div>
      </section>}

      {tab === "participants" && <section>
        <div className="panel">
          <div className="toolbar"><div><h3>Import peserta</h3><div className="muted">Masukkan URL Google Spreadsheet atau Google Docs. Data yang cocok berdasarkan nama akan diperbarui, data baru akan ditambahkan. Data lama yang tidak ada di sumber tidak dihapus.</div></div></div>
          <form onSubmit={importParticipants} style={{marginTop:16}}>
            <div className="form-grid">
              <div className="field" style={{gridColumn:"1 / -1"}}><label>URL sumber data</label><input required value={importUrl} onChange={e=>setImportUrl(e.target.value)} placeholder="https://docs.google.com/spreadsheets/d/... atau https://docs.google.com/document/d/..." /></div>
              <div className="field"><label>Jenis sumber</label><select value={importSourceType} onChange={e=>setImportSourceType(e.target.value as "auto"|"spreadsheet"|"docs")}><option value="auto">Otomatis</option><option value="spreadsheet">Google Spreadsheet</option><option value="docs">Google Docs</option></select></div>
              <div className="field"><label>Format data</label><div className="muted" style={{paddingTop:12,lineHeight:1.55}}>Spreadsheet: baris pertama berisi <b>Nama</b> dan <b>Alamat</b>. Docs: gunakan tabel dengan kolom <b>NAMA</b> dan <b>ALAMAT</b>.</div></div>
            </div>
            <div className="actions" style={{marginTop:14}}>
              <label style={{display:"flex",alignItems:"center",gap:8,fontWeight:800}}><input type="checkbox" checked={importUpdateAddress} onChange={e=>setImportUpdateAddress(e.target.checked)} /> Perbarui alamat yang sudah ada</label>
              <label style={{display:"flex",alignItems:"center",gap:8,fontWeight:800}}><input type="checkbox" checked={importReactivate} onChange={e=>setImportReactivate(e.target.checked)} /> Aktifkan kembali peserta lama</label>
            </div>
            <div className="actions" style={{marginTop:14}}><button className="btn btn-primary" disabled={importBusy}>{importBusy ? "Mengimport…" : "Import Peserta"}</button></div>
          </form>
        </div>
        <div className="panel"><h3>{participantForm.id?"Edit peserta":"Tambah peserta"}</h3><form onSubmit={saveParticipant} className="form-grid" style={{marginTop:16}}><div className="field"><label>Nama lengkap</label><input required value={participantForm.full_name} onChange={e=>setParticipantForm({...participantForm,full_name:e.target.value})}/></div><div className="field"><label>Alamat</label><input required value={participantForm.address} onChange={e=>setParticipantForm({...participantForm,address:e.target.value})}/></div><div className="actions"><button className="btn btn-primary" disabled={busy}>Simpan</button><button type="button" className="btn btn-secondary" onClick={()=>setParticipantForm(blankParticipant)}>Bersihkan</button></div></form></div>
        <div className="panel"><div className="toolbar"><h3>Data peserta</h3><input value={participantQuery} onChange={e=>setParticipantQuery(e.target.value)} placeholder="Cari nama/alamat..."/></div><div className="table-wrap"><table className="table"><thead><tr><th>Nama</th><th>Alamat</th><th>Status</th><th>Aksi</th></tr></thead><tbody>{filteredParticipants.map(p=><tr key={p.id}><td><b>{p.full_name}</b></td><td>{p.address}</td><td><span className={`pill ${p.is_active?'pill-green':'pill-gray'}`}>{p.is_active?'Aktif':'Nonaktif'}</span></td><td><div className="actions"><button className="btn btn-secondary btn-small" onClick={()=>setParticipantForm({id:p.id,full_name:p.full_name,address:p.address})}>Edit</button><button className="btn btn-secondary btn-small" onClick={()=>toggleParticipant(p)}>{p.is_active?'Nonaktifkan':'Aktifkan'}</button></div></td></tr>)}</tbody></table></div></div>
      </section>}

      {tab === "attendance" && <section>
        <div className="panel"><div className="toolbar"><div><h3>Daftar hadir</h3><div className="muted">{selectedEvent ? `${selectedEvent.name} • ${formatDate(selectedEvent.event_date)}` : "Pilih kegiatan"}</div></div><select value={selectedEvent?.id || ""} onChange={e=>setSelectedEventId(e.target.value)}><option value="">Pilih kegiatan</option>{data.events.map(e=><option key={e.id} value={e.id}>{e.name} • {formatDate(e.event_date)}</option>)}</select></div><div className="table-wrap"><table className="table"><thead><tr><th>No</th><th>Nama</th><th>Alamat</th><th>Konfirmasi</th><th>Aksi</th></tr></thead><tbody>{data.attendance.map((a,i)=><tr key={a.id}><td>{i+1}</td><td><b>{a.participant?.full_name || "Peserta"}</b></td><td>{a.participant?.address || ""}</td><td>{new Date(a.created_at).toLocaleString("id-ID",{timeZone:TZ})}</td><td><button className="btn btn-danger btn-small" onClick={()=>removeAttendance(a)}>Hapus</button></td></tr>)}{!data.attendance.length&&<tr><td colSpan={5} className="empty-cell">Belum ada data kehadiran.</td></tr>}</tbody></table></div></div>
      </section>}

      {tab === "reports" && <section>
        <div className="panel"><h3>Laporan</h3><p className="muted" style={{lineHeight:1.6}}>Export akan membuat salinan Google Docs dari template resmi dan, bila memilih PDF, membuat file PDF di Google Drive admin.</p><div className="report-event"><label>Kegiatan</label><select value={selectedEvent?.id || ""} onChange={e=>setSelectedEventId(e.target.value)}><option value="">Pilih kegiatan</option>{data.events.map(e=><option key={e.id} value={e.id}>{e.name} • {formatDate(e.event_date)}</option>)}</select></div><div className="actions" style={{marginTop:16}}><button className="btn btn-primary" disabled={!selectedEvent || !data.attendance.length} onClick={()=>exportReport("doc")}>Buat Google Docs</button><button className="btn btn-secondary" disabled={!selectedEvent || !data.attendance.length} onClick={()=>exportReport("pdf")}>Buat PDF</button></div></div>
        <div className="panel"><h3>Ringkasan laporan</h3>{selectedEvent?<div className="summary-grid"><div><span>Peserta aktif</span><b>{activeParticipants.length}</b></div><div><span>Hadir</span><b>{attendanceCount}</b></div><div><span>Belum hadir</span><b>{Math.max(activeParticipants.length-attendanceCount,0)}</b></div><div><span>Persentase</span><b>{attendanceRate}%</b></div></div>:<div className="empty-cell">Pilih kegiatan untuk melihat ringkasan.</div>}</div>
      </section>}

      {toast && <div className="toast">{toast}</div>}
    </main>
  </>;
}
