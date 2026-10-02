"use client";

import { useEffect, useMemo, useState } from "react";
import { api } from "@/lib/api";
import type { EventRow, Participant } from "@/lib/types";

const TZ = "Asia/Jakarta";

function normalizeName(value: string) {
  return value.trim().replace(/\s+/g, " ").toLowerCase();
}

function formatDate(value: string) {
  return new Intl.DateTimeFormat("id-ID", { day: "numeric", month: "long", year: "numeric", timeZone: TZ }).format(new Date(`${value.slice(0, 10)}T00:00:00+07:00`));
}

function countdownText(seconds: number) {
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  const s = seconds % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
}

export default function HomePage() {
  const [event, setEvent] = useState<EventRow | null>(null);
  const [search, setSearch] = useState("");
  const [matched, setMatched] = useState<Participant | null>(null);
  const [alreadyPresent, setAlreadyPresent] = useState(false);
  const [attendanceNumber, setAttendanceNumber] = useState<number | null>(null);
  const [attendanceName, setAttendanceName] = useState("");
  const [status, setStatus] = useState<"loading" | "not-started" | "open" | "closed" | "none">("loading");
  const [secondsLeft, setSecondsLeft] = useState(0);
  const [confirming, setConfirming] = useState(false);
  const [toast, setToast] = useState("");
  const [error, setError] = useState("");
  const [searching, setSearching] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const data = await api<{ event: EventRow | null }>("getActiveEvent");
        if (!cancelled) setEvent(data.event);
      } catch (e) {
        if (!cancelled) setError(e instanceof Error ? e.message : "Gagal memuat kegiatan.");
      }
    })();
    return () => { cancelled = true; };
  }, []);

  useEffect(() => {
    if (!event) { setStatus("none"); return; }
    const tick = () => {
      const now = Date.now();
      const open = new Date(event.open_at).getTime();
      const close = new Date(event.close_at).getTime();
      if (now < open) {
        setStatus("not-started");
        setSecondsLeft(Math.max(0, Math.floor((open - now) / 1000)));
      } else if (now >= close) {
        setStatus("closed");
        setSecondsLeft(0);
      } else {
        setStatus("open");
        setSecondsLeft(Math.max(0, Math.floor((close - now) / 1000)));
      }
    };
    tick();
    const timer = window.setInterval(tick, 1000);
    return () => window.clearInterval(timer);
  }, [event]);

  useEffect(() => {
    setMatched(null);
    setAlreadyPresent(false);
    setAttendanceNumber(null);
    setAttendanceName("");
    if (status !== "open" || !event) return;

    const q = normalizeName(search);
    if (!q || q.split(" ").length < 2) return;

    const timer = window.setTimeout(async () => {
      setSearching(true);
      try {
        const data = await api<{ participant: Participant | null; already_present: boolean }>("searchParticipant", {
          full_name: search.trim(),
          event_id: event.id
        });
        setMatched(data.participant || null);
        setAlreadyPresent(Boolean(data.already_present));
      } catch (e) {
        setMatched(null);
        setAlreadyPresent(false);
        setError(e instanceof Error ? e.message : "Gagal mencari peserta.");
      } finally {
        setSearching(false);
      }
    }, 700);

    return () => window.clearTimeout(timer);
  }, [search, status, event]);

  const searchHint = useMemo(() => {
    const q = normalizeName(search);
    if (!q) return "Contoh: Siti Aminah";
    if (q.split(" ").length < 2) return "Silakan ketik nama lengkap Anda, misalnya: Siti Aminah";
    return "";
  }, [search]);

  async function confirmAttendance() {
    if (!event || !matched || status !== "open" || alreadyPresent) return;
    try {
      const data = await api<{ success: boolean; duplicate?: boolean; message: string; display_order?: number }>("submitAttendance", {
        event_id: event.id,
        participant_id: matched.id
      });
      setConfirming(false);
      if (!data.success) {
        setAlreadyPresent(Boolean(data.duplicate));
        showToast(data.message || "Kehadiran gagal dicatat.");
        return;
      }
      setAttendanceNumber(Number(data.display_order || 0) || null);
      setAttendanceName(matched.full_name);
      setSearch("");
      setMatched(null);
      setAlreadyPresent(false);
      showToast("✓ Kehadiran berhasil dicatat.");
    } catch (e) {
      setConfirming(false);
      showToast(e instanceof Error ? e.message : "Kehadiran gagal dicatat.");
    }
  }

  function showToast(message: string) {
    setToast(message);
    window.setTimeout(() => setToast(""), 2800);
  }

  return (
    <>
      <header className="topbar">
        <div className="topbar-inner">
          <div>
            <div className="brand">Lansia SMART RSUD Wonosari</div>
            <div className="subbrand">Sehat • Mandiri • Aktif • Religius • Tidak Pikun</div>
          </div>
        </div>
      </header>

      <main className="container public-container">
        {error && <div className="error">{error}</div>}

        {!event || status === "none" ? (
          <section className="status-card">
            <div className="status-icon">📋</div>
            <h2>Belum ada kegiatan</h2>
            <p>Admin belum menyiapkan kegiatan yang dapat digunakan untuk konfirmasi kehadiran.</p>
          </section>
        ) : status === "not-started" ? (
          <section className="status-card">
            <div className="status-icon">🕐</div>
            <h2>Belum dibuka</h2>
            <p><b>{event.name}</b></p>
            <p>Akan dibuka pada <b>{formatDate(event.event_date)}</b>.</p>
          </section>
        ) : status === "closed" ? (
          <section className="status-card">
            <div className="status-icon">🔒</div>
            <h2>Konfirmasi Kehadiran Telah Ditutup</h2>
            <p>Pendaftaran untuk <b>{event.name}</b> sudah berakhir.</p>
            <p>Terima kasih.</p>
          </section>
        ) : (
          <section className="hero">
            <div className="center event-heading">
              <div className="eyebrow">KONFIRMASI KEHADIRAN</div>
              <h1>{event.name}</h1>
              <div className="location">{event.location}</div>
              <div className="date-big">{formatDate(event.event_date)}</div>
            </div>

            <label className="search-label" htmlFor="search">Cari nama lengkap Anda</label>
            <input
              id="search"
              className="search"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Ketik nama lengkap..."
              autoComplete="off"
              inputMode="text"
            />
            <div className="hint">{searchHint}</div>

            {secondsLeft <= 300 && (
              <div className="countdown">Pendaftaran ditutup dalam {countdownText(secondsLeft)}</div>
            )}

            <div className="results">
              {attendanceNumber !== null && attendanceName && (
                <div className="attendance-success-card" role="status" aria-live="polite">
                  <div className="attendance-success-icon">✓</div>
                  <div className="attendance-success-label">Kehadiran berhasil dicatat</div>
                  <div className="attendance-number">{attendanceNumber}.</div>
                  <div className="attendance-name">{attendanceName}</div>
                  <div className="attendance-success-note">Nomor kehadiran Anda</div>
                </div>
              )}

              {attendanceNumber === null && normalizeName(search).split(" ").length >= 2 && searching && (
                <div className="empty">Mencari nama…</div>
              )}

              {attendanceNumber === null && normalizeName(search).split(" ").length >= 2 && !searching && !matched && (
                <div className="empty">Nama lengkap tidak ditemukan. Periksa kembali penulisan nama Anda.</div>
              )}

              {attendanceNumber === null && matched && (
                <div className={`result-card ${alreadyPresent ? "done" : ""}`}>
                  <div className="result-main">
                    <div className="result-name">{matched.full_name}</div>
                    <div className="result-address">{matched.address || "Alamat belum diisi"}</div>
                  </div>
                  <button type="button" className="result-action" disabled={alreadyPresent} onClick={() => setConfirming(true)}>
                    {alreadyPresent ? "✓ Sudah Hadir" : "HADIR"}
                  </button>
                </div>
              )}
            </div>
          </section>
        )}
      </main>

      {confirming && matched && (
        <div className="confirm-overlay" role="dialog" aria-modal="true">
          <div className="confirm-box">
            <div className="confirm-title">Pastikan data Anda benar</div>
            <div className="confirm-person">{matched.full_name}</div>
            <div className="confirm-address">{matched.address || "Alamat belum diisi"}</div>
            <p>Apakah benar ini data Anda?</p>
            <div className="actions">
              <button className="btn btn-secondary" onClick={() => setConfirming(false)}>Batal</button>
              <button className="btn btn-primary" onClick={confirmAttendance}>Ya, Saya Hadir</button>
            </div>
          </div>
        </div>
      )}

      {toast && <div className="toast">{toast}</div>}
    </>
  );
}
