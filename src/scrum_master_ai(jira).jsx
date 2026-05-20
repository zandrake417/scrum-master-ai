import { useState, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Soniox WebSocket STT ───────────────────────────────────────────────────
const SONIOX_API_KEY = ""; // User must supply their key

// ─── Scrum Master System Prompt ─────────────────────────────────────────────
const SCRUM_MASTER_SYSTEM = `You are an elite AI Scrum Master and Project Manager with 15+ years of experience leading high-performing software teams. You do NOT simply transcribe or summarize — you THINK, ANALYZE, and PRODUCE structured artifacts like a seasoned Scrum Master would after a real meeting.

Your role:
1. **Extract & Prioritize**: Identify all action items, decisions, blockers, and risks from the discussion — even when they are implied or vaguely stated.
2. **Assign Ownership**: Deduce responsible persons from context (who spoke about it, who was mentioned, domain expertise implied).
3. **Estimate & Categorize**: Assign realistic story points (1,2,3,5,8,13) and Kanban columns (Backlog / To Do / In Progress / Review / Done) based on complexity signals in the discussion.
4. **Risk Radar**: Proactively flag risks, dependencies, and potential blockers the team may not have explicitly mentioned.
5. **Sprint Health**: Assess the sprint's health and team sentiment from the discussion tone and content.

CRITICAL INSTRUCTION: You MUST write all generated text, summaries, titles, descriptions, action items, and risks in Indonesian (Bahasa Indonesia). Only the JSON keys should remain in English.

OUTPUT FORMAT — respond ONLY with valid JSON matching this exact schema:

{
  "meeting": {
    "title": "string — concise meeting title derived from content",
    "date": "string — ISO date or 'Not specified'",
    "sprint": "string — sprint name/number or 'Not specified'",
    "duration_estimate": "string — estimated meeting duration",
    "attendees": ["array of names extracted from context"],
    "meeting_type": "string — one of: Sprint Planning / Daily Standup / Sprint Review / Sprint Retrospective / Backlog Refinement / Ad-hoc / Unknown"
  },
  "executive_summary": "string — 2-3 sentences written as a Scrum Master reporting to stakeholders, highlighting what was decided and what's at stake",
  "decisions": [
    {
      "id": "D001",
      "decision": "string — clear decision made",
      "rationale": "string — why this decision was made",
      "impact": "High | Medium | Low"
    }
  ],
  "risks": [
    {
      "id": "R001",
      "risk": "string — risk description",
      "likelihood": "High | Medium | Low",
      "impact": "High | Medium | Low",
      "mitigation": "string — recommended mitigation"
    }
  ],
  "sprint_health": {
    "velocity_signal": "On Track | At Risk | Off Track",
    "team_morale": "High | Medium | Low | Unknown",
    "blocker_count": number,
    "notes": "string — Scrum Master's honest assessment"
  },
  "kanban_tickets": [
    {
      "id": "TICKET-001",
      "title": "string — clear, actionable ticket title (imperative verb)",
      "description": "string — acceptance criteria written in Gherkin-style or clear bullet format",
      "type": "Story | Task | Bug | Spike | Epic",
      "priority": "Critical | High | Medium | Low",
      "column": "Backlog | To Do | In Progress | Review | Done",
      "story_points": number,
      "ai_reasoning": "string — JELASKAN secara singkat KENAPA tiket ini dibuat dan KENAPA diberi poin sekian berdasarkan transkrip",
      "assignee": "string — name or 'Unassigned'",
      "labels": ["array", "of", "tags"],
      "dependencies": ["array of ticket IDs this depends on"],
      "acceptance_criteria": ["array of acceptance criteria strings"]
    }
  ],
  "meeting_notes_md": "string — full professional meeting notes in Markdown format, written as a Scrum Master's official record. Include sections: ## Meeting Overview, ## Key Decisions, ## Action Items (as checklist), ## Risks & Blockers, ## Sprint Health, ## Next Steps. Use proper Markdown formatting."
}

Be decisive. Be specific. Think like a Scrum Master, not a secretary.`;

// ─── Helpers ────────────────────────────────────────────────────────────────
function generateXLSX(tickets, meeting) {
  const wb = XLSX.utils.book_new();

  // Kanban Board sheet
  const columns = ["Backlog", "To Do", "In Progress", "Review", "Done"];
  const maxRows = Math.max(...columns.map(col => tickets.filter(t => t.column === col).length));
  const boardData = [columns];
  for (let i = 0; i < maxRows; i++) {
    const row = columns.map(col => {
      const t = tickets.filter(t2 => t2.column === col)[i];
      return t ? `[${t.id}] ${t.title}\n👤 ${t.assignee} | ${t.story_points}pts | ${t.priority}` : "";
    });
    boardData.push(row);
  }
  const boardWs = XLSX.utils.aoa_to_sheet(boardData);
  boardWs["!cols"] = columns.map(() => ({ wch: 30 }));
  XLSX.utils.book_append_sheet(wb, boardWs, "Kanban Board");

  // Tickets Detail sheet
  const headers = ["ID", "Title", "Type", "Priority", "Column", "Story Points", "Assignee", "Labels", "Dependencies", "Description"];
  const rows = tickets.map(t => [
    t.id, t.title, t.type, t.priority, t.column, t.story_points,
    t.assignee, (t.labels || []).join(", "), (t.dependencies || []).join(", "), t.description
  ]);
  const detailWs = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  detailWs["!cols"] = [8, 35, 10, 10, 14, 8, 15, 20, 20, 40].map(wch => ({ wch }));
  XLSX.utils.book_append_sheet(wb, detailWs, "Ticket Details");

  // Summary sheet
  const summaryData = [
    ["Meeting Summary", ""],
    ["Title", meeting.title],
    ["Date", meeting.date],
    ["Sprint", meeting.sprint],
    ["Type", meeting.meeting_type],
    ["Attendees", (meeting.attendees || []).join(", ")],
    [""],
    ["Ticket Statistics", ""],
    ["Total Tickets", tickets.length],
    ["Total Story Points", tickets.reduce((s, t) => s + (t.story_points || 0), 0)],
    ["By Priority", ""],
    ["Critical", tickets.filter(t => t.priority === "Critical").length],
    ["High", tickets.filter(t => t.priority === "High").length],
    ["Medium", tickets.filter(t => t.priority === "Medium").length],
    ["Low", tickets.filter(t => t.priority === "Low").length],
  ];
  const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
  summaryWs["!cols"] = [{ wch: 22 }, { wch: 40 }];
  XLSX.utils.book_append_sheet(wb, summaryWs, "Summary");

  return wb;
}

function generateTrelloJSON(data) {
  const columns = ["Backlog", "To Do", "In Progress", "Review", "Done"];
  const lists = columns.map((name, i) => ({ id: `list-${i}`, name, pos: (i + 1) * 1000 }));
  const cards = data.kanban_tickets.map((t, i) => {
    const listIdx = columns.indexOf(t.column);
    return {
      id: t.id,
      name: `[${t.id}] ${t.title}`,
      desc: `**Type**: ${t.type}\n**Story Points**: ${t.story_points}\n**Priority**: ${t.priority}\n\n**Description**:\n${t.description}\n\n**Acceptance Criteria**:\n${(t.acceptance_criteria || []).map(a => `- ${a}`).join("\n")}`,
      idList: `list-${listIdx >= 0 ? listIdx : 0}`,
      pos: (i + 1) * 1000,
      labels: (t.labels || []).map((l, j) => ({ name: l, color: ["blue", "green", "orange", "red", "purple"][j % 5] })),
      members: t.assignee !== "Unassigned" ? [{ fullName: t.assignee }] : [],
      storyPoints: t.story_points,
      priority: t.priority,
      type: t.type,
      dependencies: t.dependencies || [],
    };
  });
  return {
    name: data.meeting.title,
    desc: data.executive_summary,
    lists,
    cards,
    meta: {
      sprint: data.meeting.sprint,
      date: data.meeting.date,
      meeting_type: data.meeting.meeting_type,
      attendees: data.meeting.attendees,
      generated_by: "AI Scrum Master",
    },
  };
}

function generateJiraJSON(data) {
  return {
    projects: [{
      key: "SCRUM",
      name: data.meeting.title,
      issues: data.kanban_tickets.map(t => ({
        key: t.id,
        summary: t.title,
        description: t.description,
        issuetype: { name: t.type },
        priority: { name: t.priority },
        status: { name: t.column },
        story_points: t.story_points,
        assignee: t.assignee !== "Unassigned" ? { displayName: t.assignee } : null,
        labels: t.labels || [],
        customfield_10014: t.story_points,
        acceptance_criteria: t.acceptance_criteria || [],
        issuelinks: (t.dependencies || []).map(dep => ({
          type: { name: "Depends" },
          inwardIssue: { key: dep }
        })),
      })),
      sprint: {
        name: data.meeting.sprint,
        goal: data.executive_summary,
      }
    }],
  };
}

// ─── Priority colors ─────────────────────────────────────────────────────────
const priorityStyle = {
  Critical: { bg: "#FCEBEB", color: "#A32D2D", border: "#F09595" },
  High: { bg: "#FAEEDA", color: "#854F0B", border: "#FAC775" },
  Medium: { bg: "#E6F1FB", color: "#185FA5", border: "#85B7EB" },
  Low: { bg: "#EAF3DE", color: "#3B6D11", border: "#C0DD97" },
};

const columnStyle = {
  "Backlog":     { bg: "#F1EFE8", color: "#444441" },
  "To Do":       { bg: "#E6F1FB", color: "#185FA5" },
  "In Progress": { bg: "#FAEEDA", color: "#854F0B" },
  "Review":      { bg: "#EEEDFE", color: "#3C3489" },
  "Done":        { bg: "#EAF3DE", color: "#3B6D11" },
};

// ─── Main Component ─────────────────────────────────────────────────────────
export default function ScrumMasterAI() {
  const [soniox_key, setSoniox_key] = useState("");
  
  // 1. UBAH STATE TRANSCRIPT (Baca dari Local Storage)
  const [transcript, setTranscript] = useState(() => {
    return localStorage.getItem("scrum_ai_transcript") || "";
  });
  
  const [isRecording, setIsRecording] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isUploadingAudio, setIsUploadingAudio] = useState(false);
  
  // 2. UBAH STATE RESULT (Baca dari Local Storage)
  const [result, setResult] = useState(() => {
    const saved = localStorage.getItem("scrum_ai_result");
    return saved ? JSON.parse(saved) : null;
  });
  
  const [error, setError] = useState(null);
  const [activeTab, setActiveTab] = useState("kanban");
  const [selectedTicket, setSelectedTicket] = useState(null);
  const [filterCol, setFilterCol] = useState("All");
  const [filterPriority, setFilterPriority] = useState("All");
  const [recordingTime, setRecordingTime] = useState(0);
  const [inputMode, setInputMode] = useState("text"); // "text" | "mic" | "audio"
  const [streamStatus, setStreamStatus] = useState("");

  const wsRef = useRef(null);
  const mediaRef = useRef(null);
  const timerRef = useRef(null);
  const processorRef = useRef(null);
  const audioCtxRef = useRef(null);

  // ==========================================
  // 3. TAMBAHKAN USE-EFFECT AUTO SAVE DI SINI
  // ==========================================
  
  // Auto-save transcript setiap kali mengetik/berubah
  useEffect(() => {
    localStorage.setItem("scrum_ai_transcript", transcript);
  }, [transcript]);

  // Auto-save result (tiket & notulensi) setiap kali AI selesai generate
  useEffect(() => {
    if (result) {
      localStorage.setItem("scrum_ai_result", JSON.stringify(result));
    } else {
      localStorage.removeItem("scrum_ai_result");
    }
  }, [result]);
  
  // (Opsional) Tambahkan fungsi ini jika Anda ingin membuat tombol "Clear Data" di UI
  const handleClearData = () => {
    if(confirm("Apakah Anda yakin ingin menghapus semua data lokal?")) {
      setTranscript("");
      setResult(null);
    }
  };

// ─── Fungsi untuk Memindahkan Status / Kolom Tiket ──────────────────────
  const handleUpdateTicketStatus = (ticketId, newStatus) => {
  // 1. Update status tiket di dalam state utama 'result'
  setResult(prevResult => {
    if (!prevResult || !prevResult.kanban_tickets) return prevResult;
    return {
      ...prevResult,
      kanban_tickets: prevResult.kanban_tickets.map(ticket =>
        ticket.id === ticketId ? { ...ticket, column: newStatus } : ticket
      )
    };
  });

  // 2. Update juga data tiket yang sedang aktif dibuka di modal detail
  setSelectedTicket(prevTicket => {
    if (prevTicket && prevTicket.id === ticketId) {
      return { ...prevTicket, column: newStatus };
    }
    return prevTicket;
  });
};
  
  // Timer
  useEffect(() => {
    if (isRecording) {
      timerRef.current = setInterval(() => setRecordingTime(t => t + 1), 1000);
    } else {
      clearInterval(timerRef.current);
      setRecordingTime(0);
    }
    return () => clearInterval(timerRef.current);
  }, [isRecording]);

  const formatTime = s => `${Math.floor(s / 60).toString().padStart(2, "0")}:${(s % 60).toString().padStart(2, "0")}`;

  // ── Soniox STT ────────────────────────────────────────────────────────────
  const startRecording = useCallback(async () => {
    if (!soniox_key.trim()) {
      setError("Masukkan Soniox API Key terlebih dahulu.");
      return;
    }
    setError(null);
    setStreamStatus("Menghubungkan ke Soniox...");

    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      mediaRef.current = stream;

      audioCtxRef.current = new AudioContext({ sampleRate: 16000 });
      const source = audioCtxRef.current.createMediaStreamSource(stream);

      await audioCtxRef.current.audioWorklet.addModule(
        URL.createObjectURL(new Blob([`
          class PCMProcessor extends AudioWorkletProcessor {
            process(inputs) {
              const ch = inputs[0][0];
              if (ch) {
                const buf = new Int16Array(ch.length);
                for (let i = 0; i < ch.length; i++) {
                  buf[i] = Math.max(-32768, Math.min(32767, Math.round(ch[i] * 32767)));
                }
                this.port.postMessage(buf.buffer, [buf.buffer]);
              }
              return true;
            }
          }
          registerProcessor("pcm-processor", PCMProcessor);
        `], { type: "application/javascript" }))
      );

      processorRef.current = new AudioWorkletNode(audioCtxRef.current, "pcm-processor");
      source.connect(processorRef.current);

      const ws = new WebSocket(`wss://api.soniox.com/transcribe-websocket`);
      wsRef.current = ws;

      ws.onopen = () => {
        setStreamStatus("Terhubung — merekam...");
        setIsRecording(true);
        ws.send(JSON.stringify({
          api_key: soniox_key,
          sample_rate: 16000,
          num_audio_channels: 1,
          include_nonfinal: true,
          enable_endpoint_detection: false,
        }));
      };

      ws.onmessage = (e) => {
        const msg = JSON.parse(e.data);
        if (msg.words) {
          const text = msg.words.map(w => w.text).join(" ");
          if (msg.final_proc_time_ms) {
            setTranscript(prev => prev ? prev + " " + text : text);
          }
        }
      };

      ws.onerror = () => setError("Koneksi Soniox gagal. Periksa API key & koneksi internet.");
      ws.onclose = () => setStreamStatus("");

      processorRef.current.port.onmessage = (e) => {
        if (ws.readyState === WebSocket.OPEN) {
          ws.send(e.data);
        }
      };
    } catch (err) {
      setError("Tidak dapat mengakses mikrofon: " + err.message);
      setStreamStatus("");
    }
  }, [soniox_key]);

  const stopRecording = useCallback(() => {
    setIsRecording(false);
    if (wsRef.current) { wsRef.current.close(); wsRef.current = null; }
    if (processorRef.current) { processorRef.current.disconnect(); processorRef.current = null; }
    if (audioCtxRef.current) { audioCtxRef.current.close(); audioCtxRef.current = null; }
    if (mediaRef.current) { mediaRef.current.getTracks().forEach(t => t.stop()); mediaRef.current = null; }
    setStreamStatus("");
  }, []);

  // ── Groq Whisper Audio Upload ─────────────────────────────────────────────
  const handleAudioUpload = async (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const MAX_FILE_SIZE = 25 * 1024 * 1024; // 25 MB
    if (file.size > MAX_FILE_SIZE) {
      setError("❌ Ukuran file terlalu besar (Maksimal 25MB).\nTips: Ubah format ke .m4a atau kompres ke bitrate lebih rendah.");
      return;
    }

    setIsUploadingAudio(true);
    setError(null);
    setStreamStatus("Mengunggah & mentranskrip audio...");
    
    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("model", "whisper-large-v3");
      formData.append("response_format", "json");
      formData.append("language", "id");

      const response = await fetch("https://api.groq.com/openai/v1/audio/transcriptions", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${import.meta.env.VITE_GROQ_API_KEY}`
        },
        body: formData
      });

      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.error?.message || "Gagal memproses audio di Groq.");
      }

      const data = await response.json();
      setTranscript(prev => prev ? `${prev}\n\n${data.text}` : data.text);
      setInputMode("text"); // Otomatis kembali ke mode teks untuk melihat hasilnya
      setStreamStatus("");
    } catch (err) {
      console.error("Groq Upload Error:", err);
      setError(`Gagal memproses audio: ${err.message}`);
      setStreamStatus("");
    } finally {
      setIsUploadingAudio(false);
      event.target.value = null; // Reset input agar bisa re-upload file yang sama jika perlu
    }
  };

  // ── Snifox AI (OpenAI-compatible) ─────────────────────────────────────────
// ── Otak Scrum Master Pakai GROQ (Llama 3) ─────────────────────────────────
const analyzeWithAI = useCallback(async () => {
  if (!transcript.trim()) {
    setError("Belum ada transkrip. Rekam atau ketik terlebih dahulu.");
    return;
  }
  setIsProcessing(true);
  setError(null);
  setResult(null);

  try {
    // Menembak langsung ke API Groq, bukan ke /api/chat/completions (Vite proxy)
    const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: { 
        "Content-Type": "application/json",
        // Menyelipkan Authorization Header langsung di sini
        "Authorization": `Bearer ${import.meta.env.VITE_GROQ_API_KEY}`
      },
      body: JSON.stringify({
        model: "llama-3.3-70b-versatile", // Menggunakan model teks paling pintar di Groq
        messages: [
          { role: "system", content: SCRUM_MASTER_SYSTEM },
          {
            role: "user",
            content: `Analyze this meeting transcript as a Scrum Master. Extract all actionable insights and produce the full JSON artifact:\n\n---\n${transcript}\n---`,
          },
        ],
        response_format: { type: "json_object" }, // Paksa Groq membalas format JSON
        temperature: 0.2,
      }),
    });

    if (!res.ok) {
      const errorData = await res.json();
      throw new Error(`API request gagal: ${res.status} — ${errorData.error?.message || res.statusText}`);
    }

    const data = await res.json();
    const rawText = data?.choices?.[0]?.message?.content || "";

    if (!rawText) throw new Error("AI tidak menghasilkan respons yang valid.");
    
    // Parse hasil JSON dari Llama 3
    const parsed = JSON.parse(rawText);
    setResult(parsed);
    setActiveTab("kanban");
  } catch (err) {
    console.error("Analysis Error:", err);
    setError("Error: " + err.message);
  } finally {
    setIsProcessing(false);
  }
}, [transcript]);

  // ── Downloads ─────────────────────────────────────────────────────────────
  const downloadMD = () => {
    const blob = new Blob([result.meeting_notes_md], { type: "text/markdown" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${(result.meeting.title || "meeting").replace(/\s+/g, "_")}_notulensi.md`;
    a.click();
  };

  const downloadXLSX = () => {
    const wb = generateXLSX(result.kanban_tickets, result.meeting);
    XLSX.writeFile(wb, `${(result.meeting.title || "kanban").replace(/\s+/g, "_")}_kanban.xlsx`);
  };

  const downloadJSON = (type) => {
    const data = type === "trello" ? generateTrelloJSON(result) : generateJiraJSON(result);
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `kanban_${type}.json`;
    a.click();
  };

  // ── Filtered tickets ──────────────────────────────────────────────────────
  const filteredTickets = result?.kanban_tickets?.filter(t => {
    if (filterCol !== "All" && t.column !== filterCol) return false;
    if (filterPriority !== "All" && t.priority !== filterPriority) return false;
    return true;
  }) || [];

  const KANBAN_COLS = ["Backlog", "To Do", "In Progress", "Review", "Done"];

  const handleExportToJira = async () => {
    if (!result || !result.kanban_tickets || result.kanban_tickets.length === 0) return;

    // Kita tidak perlu lagi prompt Domain Jira karena sudah di-hardcode di Vite Proxy
    const email = prompt("Masukkan Email akun Jira Anda:");
    if (!email) return;
    
    const apiToken = prompt("Masukkan Jira API Token Anda:\n(Bukan password email!)");
    if (!apiToken) return;
    
    const projectKey = prompt("Masukkan Project Key Jira Anda\n(contoh: SCRUM):");
    if (!projectKey) return;

    const auth = btoa(`${email}:${apiToken}`);
    let successCount = 0;

    for (const ticket of result.kanban_tickets) {
      try {
        // PERHATIKAN: URL-nya diubah jadi memanggil /jira-api/ lokal
        const response = await fetch(`/jira-api/rest/api/3/issue`, {
          method: "POST",
          headers: {
            "Authorization": `Basic ${auth}`,
            "Accept": "application/json",
            "Content-Type": "application/json",
            "X-Atlassian-Token": "no-check" // Header wajib Jira
          },
          body: JSON.stringify({
            fields: {
              project: { key: projectKey },
              summary: ticket.title,
              description: {
                type: "doc",
                version: 1,
                content: [
                  {
                    type: "paragraph",
                    content: [{ type: "text", text: ticket.description || "Tidak ada deskripsi." }]
                  }
                ]
              },
              issuetype: { name: ticket.type.toLowerCase() === "bug" ? "Bug" : "Task" } 
              // Jika masih gagal, ganti jadi "Story"
            }
          })
        });

        if (response.ok) {
          successCount++;
          console.log(`✅ SUCCESS: Tiket '${ticket.title}' berhasil dibuat.`);
        } else {
          const errData = await response.json();
          alert(`❌ Gagal export tiket '${ticket.title}'.\nAlasan Jira: ${JSON.stringify(errData.errors || errData)}`);
        }
      } catch (error) {
        alert(`🚨 ERROR KONEKSI: ${error.message}`);
      }
    }

    alert(`Proses selesai! ${successCount} tiket berhasil dibuat di Jira.`);
  };

  // ─────────────────────────────────────────────────────────────────────────
  return (
    <div style={{ fontFamily: "'IBM Plex Sans', system-ui, sans-serif", maxWidth: 900, margin: "0 auto", padding: "1.5rem 1rem", color: "var(--color-text-primary)" }}>
      <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet" />

      {/* Header */}
      <div style={{ marginBottom: "2rem" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
          <div style={{ width: 32, height: 32, borderRadius: 8, background: "#185FA5", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <i className="ti ti-robot" style={{ color: "#fff", fontSize: 18 }} />
          </div>
          <h1 style={{ margin: 0, fontSize: 20, fontWeight: 600, letterSpacing: "-0.3px" }}>AI Scrum Master</h1>
          <span style={{ fontSize: 11, fontWeight: 500, padding: "2px 8px", background: "#E6F1FB", color: "#185FA5", borderRadius: 20, border: "0.5px solid #85B7EB" }}>BETA</span>
        </div>
        <p style={{ margin: 0, fontSize: 13, color: "var(--color-text-secondary)" }}>
          Transformasi rapat menjadi notulensi, tiket Kanban, dan laporan sprint — otomatis, oleh AI.
        </p>
      </div>

      {/* Step 1: Config */}
      <section style={{ marginBottom: "1.25rem", background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 12, padding: "1rem 1.25rem" }}>
        <p style={{ margin: "0 0 10px", fontSize: 12, fontWeight: 500, color: "var(--color-text-secondary)", textTransform: "uppercase", letterSpacing: "0.06em" }}>① Soniox Speech-to-Text</p>
        <div style={{ display: "flex", gap: 8 }}>
          <input
            type="password"
            placeholder="Soniox API Key (opsional — bisa input teks manual atau upload file)"
            value={soniox_key}
            onChange={e => setSoniox_key(e.target.value)}
            style={{ flex: 1, fontFamily: "inherit", fontSize: 13, padding: "7px 10px", border: "0.5px solid var(--color-border-secondary)", borderRadius: 8, background: "var(--color-background-secondary)", color: "var(--color-text-primary)" }}
          />
          <a href="https://soniox.com" target="_blank" rel="noreferrer" style={{ fontSize: 12, color: "#185FA5", display: "flex", alignItems: "center", gap: 4, textDecoration: "none", whiteSpace: "nowrap" }}>
            Dapatkan key <i className="ti ti-external-link" style={{ fontSize: 13 }} />
          </a>
        </div>
      </section>

      {/* Step 2: Input */}
      <section style={{ marginBottom: "1.25rem", background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 12, padding: "1rem 1.25rem" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10, flexWrap: "wrap", gap: 8 }}>
          <p style={{ margin: 0, fontSize: 12, fontWeight: 500, color: "var(--color-text-secondary)", textTransform: "uppercase", letterSpacing: "0.06em" }}>② Transkrip Rapat</p>
          <div style={{ display: "flex", gap: 4, background: "var(--color-background-secondary)", borderRadius: 8, padding: 3, border: "0.5px solid var(--color-border-tertiary)" }}>
            {["text", "mic", "audio"].map(m => (
              <button key={m} onClick={() => setInputMode(m)} style={{ padding: "4px 12px", borderRadius: 6, border: "none", cursor: "pointer", fontSize: 12, fontWeight: 500, background: inputMode === m ? "var(--color-background-primary)" : "transparent", color: inputMode === m ? "var(--color-text-primary)" : "var(--color-text-secondary)", boxShadow: inputMode === m ? "0 1px 3px rgba(0,0,0,0.08)" : "none", display: "flex", alignItems: "center" }}>
                <i className={`ti ti-${m === "text" ? "keyboard" : m === "mic" ? "microphone" : "upload"}`} style={{ marginRight: 4 }} />
                {m === "text" ? "Teks" : m === "mic" ? "Mikrofon" : "Upload Audio"}
              </button>
            ))}
          </div>
        </div>

        {inputMode === "mic" && (
          <div style={{ marginBottom: 12, display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: isRecording ? "#FCEBEB" : "var(--color-background-secondary)", borderRadius: 8, border: `0.5px solid ${isRecording ? "#F09595" : "var(--color-border-tertiary)"}` }}>
            <button
              onClick={isRecording ? stopRecording : startRecording}
              style={{ padding: "7px 16px", borderRadius: 8, border: "none", cursor: "pointer", fontWeight: 500, fontSize: 13, fontFamily: "inherit", background: isRecording ? "#A32D2D" : "#185FA5", color: "#fff", display: "flex", alignItems: "center", gap: 6 }}
            >
              <i className={`ti ti-${isRecording ? "player-stop" : "microphone"}`} />
              {isRecording ? "Hentikan" : "Mulai Rekam"}
            </button>
            {isRecording && (
              <>
                <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 15, fontWeight: 500, color: "#A32D2D" }}>{formatTime(recordingTime)}</span>
                <span style={{ width: 8, height: 8, borderRadius: "50%", background: "#E24B4A", animation: "pulse 1s infinite" }} />
              </>
            )}
            {streamStatus && <span style={{ fontSize: 12, color: "var(--color-text-secondary)" }}>{streamStatus}</span>}
          </div>
        )}

        {inputMode === "audio" && (
          <div style={{ marginBottom: 12, display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: "var(--color-background-secondary)", borderRadius: 8, border: "0.5px solid var(--color-border-tertiary)" }}>
            <label style={{ padding: "7px 16px", borderRadius: 8, border: "none", cursor: isUploadingAudio ? "not-allowed" : "pointer", fontWeight: 500, fontSize: 13, fontFamily: "inherit", background: isUploadingAudio ? "#B5D4F4" : "#185FA5", color: "#fff", display: "flex", alignItems: "center", gap: 6 }}>
              {isUploadingAudio ? (
                <><i className="ti ti-loader" style={{ animation: "spin 1s linear infinite" }} /> Memproses File...</>
              ) : (
                <><i className="ti ti-upload" /> Pilih File MP3 / M4A</>
              )}
              <input
                type="file"
                accept="audio/mp3, audio/wav, audio/m4a, audio/mpeg"
                onChange={handleAudioUpload}
                disabled={isUploadingAudio}
                style={{ display: "none" }}
              />
            </label>
            <span style={{ fontSize: 12, color: "var(--color-text-secondary)" }}>
              {isUploadingAudio ? streamStatus : "Batas ukuran maksimal: 25MB"}
            </span>
          </div>
        )}

        <textarea
          value={transcript}
          onChange={e => setTranscript(e.target.value)}
          placeholder={`Ketik atau paste transkrip rapat di sini...\n\nContoh:\n"Hari ini kita sprint planning. Tim ada Budi, Sari, dan Andi.\nSari bilang ada bug kritis di login. Budi akan kerjakan fitur search minggu ini.\nAda blocker di API payment gateway yang belum selesai dari tim external..."`}
          rows={8}
          style={{ width: "100%", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5, lineHeight: 1.6, border: "0.5px solid var(--color-border-secondary)", borderRadius: 8, padding: "10px 12px", resize: "vertical", background: "var(--color-background-secondary)", color: "var(--color-text-primary)", boxSizing: "border-box" }}
        />
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8 }}>
          <span style={{ fontSize: 12, color: "var(--color-text-secondary)" }}>{transcript.length} karakter · {transcript.split(/\s+/).filter(Boolean).length} kata</span>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={() => setTranscript("")} style={{ padding: "6px 12px", borderRadius: 8, border: "0.5px solid var(--color-border-secondary)", cursor: "pointer", fontSize: 12, background: "transparent", color: "var(--color-text-secondary)", fontFamily: "inherit" }}>
              <i className="ti ti-trash" style={{ marginRight: 4 }} />Hapus
            </button>
            <button
              onClick={analyzeWithAI}
              disabled={isProcessing || !transcript.trim()}
              style={{ padding: "6px 16px", borderRadius: 8, border: "none", cursor: isProcessing ? "not-allowed" : "pointer", fontSize: 13, fontWeight: 500, background: isProcessing ? "#B5D4F4" : "#185FA5", color: "#fff", fontFamily: "inherit", display: "flex", alignItems: "center", gap: 6 }}
            >
              {isProcessing ? (
                <><i className="ti ti-loader" style={{ animation: "spin 1s linear infinite" }} />Menganalisis...</>
              ) : (
                <><i className="ti ti-brain" />Analisis sebagai Scrum Master</>
              )}
            </button>
          </div>
        </div>
      </section>

      {/* Error */}
      {error && (
        <div style={{ marginBottom: "1.25rem", padding: "10px 14px", background: "#FCEBEB", border: "0.5px solid #F09595", borderRadius: 8, color: "#A32D2D", fontSize: 13, display: "flex", alignItems: "flex-start", gap: 8 }}>
          <i className="ti ti-alert-circle" style={{ fontSize: 16, marginTop: 1, flexShrink: 0 }} />
          {error}
        </div>
      )}

      {/* Loading */}
      {isProcessing && (
        <div style={{ marginBottom: "1.25rem", padding: "16px", background: "var(--color-background-secondary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 12, textAlign: "center" }}>
          <div style={{ fontSize: 13, color: "var(--color-text-secondary)", marginBottom: 8 }}>
            <i className="ti ti-brain" style={{ marginRight: 6, fontSize: 16 }} />
            AI Scrum Master sedang menganalisis rapat...
          </div>
          <div style={{ fontSize: 12, color: "var(--color-text-secondary)" }}>
            Mengekstrak keputusan · Membuat tiket · Menilai risiko · Menghasilkan notulensi
          </div>
          <div style={{ marginTop: 12, height: 3, background: "var(--color-border-tertiary)", borderRadius: 2, overflow: "hidden" }}>
            <div style={{ height: "100%", width: "60%", background: "#185FA5", borderRadius: 2, animation: "loading 1.5s ease-in-out infinite" }} />
          </div>
        </div>
      )}

      {/* Results */}
      {result && (
        <div>
          {/* Meeting Header */}
          <div style={{ marginBottom: "1.25rem", background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 12, padding: "1rem 1.25rem" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 8 }}>
              <div>
                <h2 style={{ margin: "0 0 4px", fontSize: 17, fontWeight: 600 }}>{result.meeting.title}</h2>
                <p style={{ margin: 0, fontSize: 12.5, color: "var(--color-text-secondary)", lineHeight: 1.6 }}>
                  <span style={{ marginRight: 12 }}><i className="ti ti-calendar" style={{ marginRight: 4 }} />{result.meeting.date}</span>
                  <span style={{ marginRight: 12 }}><i className="ti ti-refresh" style={{ marginRight: 4 }} />{result.meeting.sprint}</span>
                  <span><i className="ti ti-tag" style={{ marginRight: 4 }} />{result.meeting.meeting_type}</span>
                </p>
              </div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {/* Download buttons */}
                {[
                  { label: "Notulensi .md", icon: "ti-file-text", action: downloadMD, color: "#185FA5" },
                  { label: "Kanban .xlsx", icon: "ti-table", action: downloadXLSX, color: "#3B6D11" },
                  { label: "Trello .json", icon: "ti-brand-trello", action: () => downloadJSON("trello"), color: "#3C3489" },
                  { label: "Jira .json", icon: "ti-bug", action: () => downloadJSON("jira"), color: "#854F0B" },
                ].map(btn => (
                  <button key={btn.label} onClick={btn.action} style={{ padding: "5px 10px", borderRadius: 8, border: `0.5px solid ${btn.color}40`, cursor: "pointer", fontSize: 11.5, fontWeight: 500, fontFamily: "inherit", background: `${btn.color}12`, color: btn.color, display: "flex", alignItems: "center", gap: 5 }}>
                    <i className={`ti ${btn.icon}`} style={{ fontSize: 13 }} />
                    {btn.label}
                  </button>
                ))}

                {/* ─── TAMBAHKAN TOMBOL EKSPOR JIRA DI SINI ─── */}
                <button 
                    onClick={handleExportToJira} 
                    style={{ 
                      padding: "5px 10px", 
                      borderRadius: 8, 
                      border: "none", 
                      cursor: "pointer", 
                      fontSize: 11.5, 
                      fontWeight: 600, 
                      fontFamily: "inherit", 
                      background: "#0052CC", // Warna biru Jira
                      color: "#ffffff", 
                      display: "flex", 
                      alignItems: "center", 
                      gap: 5,
                      boxShadow: "0 2px 4px rgba(0,82,204,0.2)"
                    }}
                  >
                    <i className="ti ti-brand-jira" style={{ fontSize: 13 }} />
                    Push to Jira API
                  </button>
                  {/* ────────────────────────────────────────────── */}

              </div>
            </div>

            {/* Executive Summary */}
            <div style={{ marginTop: 12, padding: "10px 12px", background: "var(--color-background-secondary)", borderRadius: 8, borderLeft: "3px solid #185FA5" }}>
              <p style={{ margin: 0, fontSize: 13, lineHeight: 1.65, color: "var(--color-text-primary)" }}>{result.executive_summary}</p>
            </div>

            {/* Stats row */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(0, 1fr))", gap: 8, marginTop: 12 }}>
              {[
                { label: "Total Tiket", value: result.kanban_tickets?.length || 0, color: "#185FA5" },
                { label: "Story Points", value: result.kanban_tickets?.reduce((s, t) => s + (t.story_points || 0), 0) || 0, color: "#3C3489" },
                { label: "Risiko", value: result.risks?.length || 0, color: "#A32D2D" },
                { label: "Keputusan", value: result.decisions?.length || 0, color: "#3B6D11" },
                { label: "Sprint Health", value: result.sprint_health?.velocity_signal || "—", color: result.sprint_health?.velocity_signal === "On Track" ? "#3B6D11" : result.sprint_health?.velocity_signal === "At Risk" ? "#854F0B" : "#A32D2D", small: true },
              ].map(s => (
                <div key={s.label} style={{ background: "var(--color-background-secondary)", borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
                  <div style={{ fontSize: s.small ? 11 : 18, fontWeight: 600, color: s.color }}>{s.value}</div>
                  <div style={{ fontSize: 11, color: "var(--color-text-secondary)", marginTop: 2 }}>{s.label}</div>
                </div>
              ))}
            </div>
          </div>

          {/* Tabs */}
          <div style={{ display: "flex", gap: 2, marginBottom: "1rem", background: "var(--color-background-secondary)", borderRadius: 10, padding: 4, border: "0.5px solid var(--color-border-tertiary)" }}>
            {[
              { id: "kanban", label: "Kanban Board", icon: "ti-layout-columns" },
              { id: "risks", label: "Risiko & Keputusan", icon: "ti-alert-triangle" },
              { id: "attendees", label: "Peserta & Sprint", icon: "ti-users" },
            ].map(tab => (
              <button key={tab.id} onClick={() => setActiveTab(tab.id)} style={{ flex: 1, padding: "7px 10px", borderRadius: 8, border: "none", cursor: "pointer", fontSize: 12.5, fontWeight: activeTab === tab.id ? 500 : 400, fontFamily: "inherit", background: activeTab === tab.id ? "var(--color-background-primary)" : "transparent", color: activeTab === tab.id ? "var(--color-text-primary)" : "var(--color-text-secondary)", boxShadow: activeTab === tab.id ? "0 1px 3px rgba(0,0,0,0.07)" : "none", display: "flex", alignItems: "center", justifyContent: "center", gap: 5 }}>
                <i className={`ti ${tab.icon}`} style={{ fontSize: 13 }} />
                {tab.label}
              </button>
            ))}
          </div>

          {/* Kanban Tab */}
          {activeTab === "kanban" && (
            <div>
              {/* Filters */}
              <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
                <select value={filterCol} onChange={e => setFilterCol(e.target.value)} style={{ padding: "5px 10px", borderRadius: 8, border: "0.5px solid var(--color-border-secondary)", fontSize: 12, background: "var(--color-background-secondary)", color: "var(--color-text-primary)", fontFamily: "inherit", cursor: "pointer" }}>
                  <option value="All">Semua Kolom</option>
                  {KANBAN_COLS.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
                <select value={filterPriority} onChange={e => setFilterPriority(e.target.value)} style={{ padding: "5px 10px", borderRadius: 8, border: "0.5px solid var(--color-border-secondary)", fontSize: 12, background: "var(--color-background-secondary)", color: "var(--color-text-primary)", fontFamily: "inherit", cursor: "pointer" }}>
                  <option value="All">Semua Prioritas</option>
                  {["Critical", "High", "Medium", "Low"].map(p => <option key={p} value={p}>{p}</option>)}
                </select>
                <span style={{ fontSize: 12, color: "var(--color-text-secondary)", display: "flex", alignItems: "center" }}>
                  {filteredTickets.length} tiket · {filteredTickets.reduce((s, t) => s + (t.story_points || 0), 0)} pts
                </span>
              </div>

              {/* Board */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(155px, 1fr))", gap: 8 }}>
                {KANBAN_COLS.map(col => {
                  const colTickets = filteredTickets.filter(t => t.column === col);
                  const cs = columnStyle[col] || { bg: "#F1EFE8", color: "#444441" };
                  return (
                    <div key={col}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, padding: "5px 8px", borderRadius: 7, background: cs.bg }}>
                        <span style={{ fontSize: 11, fontWeight: 600, color: cs.color, letterSpacing: "0.02em" }}>{col.toUpperCase()}</span>
                        <span style={{ fontSize: 11, fontWeight: 600, color: cs.color, background: `${cs.color}22`, borderRadius: 10, padding: "1px 6px" }}>{colTickets.length}</span>
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                        {colTickets.map(ticket => {
                          const ps = priorityStyle[ticket.priority] || priorityStyle.Medium;
                          return (
                            <button
                              key={ticket.id}
                              onClick={() => setSelectedTicket(ticket)}
                              style={{ textAlign: "left", padding: "10px", borderRadius: 8, border: "0.5px solid var(--color-border-tertiary)", background: "var(--color-background-primary)", cursor: "pointer", width: "100%", transition: "border-color 0.15s" }}
                            >
                              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 5, gap: 4 }}>
                                <span style={{ fontSize: 10, fontFamily: "'IBM Plex Mono', monospace", color: "var(--color-text-secondary)", fontWeight: 500 }}>{ticket.id}</span>
                                <span style={{ fontSize: 10, fontWeight: 600, padding: "1px 5px", borderRadius: 5, background: ps.bg, color: ps.color, border: `0.5px solid ${ps.border}`, whiteSpace: "nowrap" }}>{ticket.priority}</span>
                              </div>
                              <p style={{ margin: "0 0 6px", fontSize: 12, fontWeight: 500, lineHeight: 1.35, color: "var(--color-text-primary)" }}>{ticket.title}</p>
                              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                <span style={{ fontSize: 10, color: "var(--color-text-secondary)" }}>
                                  <i className="ti ti-user" style={{ fontSize: 10, marginRight: 2 }} />
                                  {ticket.assignee === "Unassigned" ? "—" : ticket.assignee?.split(" ")[0]}
                                </span>
                                <span style={{ fontSize: 10, fontWeight: 600, color: "#185FA5", background: "#E6F1FB", padding: "1px 6px", borderRadius: 10 }}>{ticket.story_points}pt</span>
                              </div>
                              {ticket.type && (
                                <div style={{ marginTop: 5 }}>
                                  <span style={{ fontSize: 10, padding: "1px 6px", borderRadius: 5, background: "var(--color-background-secondary)", color: "var(--color-text-secondary)", border: "0.5px solid var(--color-border-tertiary)" }}>{ticket.type}</span>
                                </div>
                              )}
                            </button>
                          );
                        })}
                        {colTickets.length === 0 && (
                          <div style={{ padding: "16px 8px", textAlign: "center", color: "var(--color-text-secondary)", fontSize: 11, border: "0.5px dashed var(--color-border-tertiary)", borderRadius: 8 }}>—</div>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* Risks Tab */}
          {activeTab === "risks" && (
            <div style={{ display: "grid", gap: "1.25rem", gridTemplateColumns: "1fr 1fr" }}>
              <div>
                <h3 style={{ margin: "0 0 10px", fontSize: 14, fontWeight: 500 }}>
                  <i className="ti ti-alert-triangle" style={{ marginRight: 6, color: "#854F0B" }} />
                  Risiko ({result.risks?.length || 0})
                </h3>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {(result.risks || []).map(r => (
                    <div key={r.id} style={{ padding: "10px 12px", borderRadius: 8, border: "0.5px solid var(--color-border-tertiary)", background: "var(--color-background-primary)" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 5 }}>
                        <span style={{ fontSize: 10, fontFamily: "'IBM Plex Mono', monospace", color: "var(--color-text-secondary)" }}>{r.id}</span>
                        <div style={{ display: "flex", gap: 4 }}>
                          {[{ label: r.likelihood, prefix: "L:" }, { label: r.impact, prefix: "I:" }].map(({ label, prefix }) => {
                            const c = label === "High" ? priorityStyle.Critical : label === "Medium" ? priorityStyle.Medium : priorityStyle.Low;
                            return <span key={prefix} style={{ fontSize: 10, padding: "1px 5px", borderRadius: 5, background: c.bg, color: c.color }}>{prefix}{label}</span>;
                          })}
                        </div>
                      </div>
                      <p style={{ margin: "0 0 5px", fontSize: 12.5, fontWeight: 500, lineHeight: 1.4 }}>{r.risk}</p>
                      <p style={{ margin: 0, fontSize: 11.5, color: "var(--color-text-secondary)", lineHeight: 1.4 }}>
                        <i className="ti ti-shield" style={{ marginRight: 3 }} />{r.mitigation}
                      </p>
                    </div>
                  ))}
                  {!result.risks?.length && <div style={{ fontSize: 13, color: "var(--color-text-secondary)", padding: 12 }}>Tidak ada risiko teridentifikasi.</div>}
                </div>
              </div>
              <div>
                <h3 style={{ margin: "0 0 10px", fontSize: 14, fontWeight: 500 }}>
                  <i className="ti ti-check-circle" style={{ marginRight: 6, color: "#3B6D11" }} />
                  Keputusan ({result.decisions?.length || 0})
                </h3>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  {(result.decisions || []).map(d => {
                    const c = d.impact === "High" ? priorityStyle.Critical : d.impact === "Medium" ? priorityStyle.High : priorityStyle.Low;
                    return (
                      <div key={d.id} style={{ padding: "10px 12px", borderRadius: 8, border: "0.5px solid var(--color-border-tertiary)", background: "var(--color-background-primary)" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 5 }}>
                          <span style={{ fontSize: 10, fontFamily: "'IBM Plex Mono', monospace", color: "var(--color-text-secondary)" }}>{d.id}</span>
                          <span style={{ fontSize: 10, padding: "1px 5px", borderRadius: 5, background: c.bg, color: c.color }}>Impact: {d.impact}</span>
                        </div>
                        <p style={{ margin: "0 0 5px", fontSize: 12.5, fontWeight: 500, lineHeight: 1.4 }}>{d.decision}</p>
                        <p style={{ margin: 0, fontSize: 11.5, color: "var(--color-text-secondary)", lineHeight: 1.4 }}>{d.rationale}</p>
                      </div>
                    );
                  })}
                  {!result.decisions?.length && <div style={{ fontSize: 13, color: "var(--color-text-secondary)", padding: 12 }}>Tidak ada keputusan teridentifikasi.</div>}
                </div>
              </div>
            </div>
          )}

          {/* Attendees Tab */}
          {activeTab === "attendees" && (
            <div style={{ display: "grid", gap: "1.25rem", gridTemplateColumns: "1fr 1fr" }}>
              <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 12, padding: "1rem 1.25rem" }}>
                <h3 style={{ margin: "0 0 12px", fontSize: 14, fontWeight: 500 }}>
                  <i className="ti ti-users" style={{ marginRight: 6 }} />Peserta Rapat
                </h3>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  {(result.meeting.attendees || []).map((a, i) => (
                    <div key={i} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <div style={{ width: 30, height: 30, borderRadius: "50%", background: ["#E6F1FB", "#EAF3DE", "#EEEDFE", "#FAEEDA"][i % 4], display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 600, fontSize: 12, color: ["#185FA5", "#3B6D11", "#3C3489", "#854F0B"][i % 4], flexShrink: 0 }}>
                        {a.split(" ").map(n => n[0]).join("").slice(0, 2)}
                      </div>
                      <span style={{ fontSize: 13 }}>{a}</span>
                    </div>
                  ))}
                  {!result.meeting.attendees?.length && <div style={{ fontSize: 13, color: "var(--color-text-secondary)" }}>Tidak teridentifikasi</div>}
                </div>
              </div>
              <div style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-tertiary)", borderRadius: 12, padding: "1rem 1.25rem" }}>
                <h3 style={{ margin: "0 0 12px", fontSize: 14, fontWeight: 500 }}>
                  <i className="ti ti-activity" style={{ marginRight: 6 }} />Sprint Health
                </h3>
                {result.sprint_health && (() => {
                  const sh = result.sprint_health;
                  const vc = sh.velocity_signal === "On Track" ? "#3B6D11" : sh.velocity_signal === "At Risk" ? "#854F0B" : "#A32D2D";
                  return (
                    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                      {[
                        { label: "Velocity", value: sh.velocity_signal, color: vc },
                        { label: "Morale", value: sh.team_morale, color: sh.team_morale === "High" ? "#3B6D11" : sh.team_morale === "Medium" ? "#854F0B" : "#A32D2D" },
                        { label: "Blockers", value: `${sh.blocker_count} aktif`, color: sh.blocker_count > 2 ? "#A32D2D" : sh.blocker_count > 0 ? "#854F0B" : "#3B6D11" },
                      ].map(s => (
                        <div key={s.label} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: "0.5px solid var(--color-border-tertiary)" }}>
                          <span style={{ fontSize: 12.5, color: "var(--color-text-secondary)" }}>{s.label}</span>
                          <span style={{ fontSize: 12.5, fontWeight: 600, color: s.color }}>{s.value}</span>
                        </div>
                      ))}
                      {sh.notes && <p style={{ margin: "6px 0 0", fontSize: 12, color: "var(--color-text-secondary)", lineHeight: 1.5, fontStyle: "italic" }}>"{sh.notes}"</p>}
                    </div>
                  );
                })()}
              </div>
            </div>
          )}
        </div>
      )}

      {/* Ticket Detail Modal */}
      {selectedTicket && (
        <div onClick={() => setSelectedTicket(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 20 }}>
          <div onClick={e => e.stopPropagation()} style={{ background: "var(--color-background-primary)", border: "0.5px solid var(--color-border-secondary)", borderRadius: 14, padding: "1.25rem", maxWidth: 520, width: "100%", maxHeight: "80vh", overflow: "auto" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
              <div>
                <span style={{ fontSize: 11, fontFamily: "'IBM Plex Mono', monospace", color: "var(--color-text-secondary)" }}>{selectedTicket.id}</span>
                <h3 style={{ margin: "4px 0 0", fontSize: 15, fontWeight: 600, lineHeight: 1.35 }}>{selectedTicket.title}</h3>
              </div>
              <button onClick={() => setSelectedTicket(null)} style={{ background: "transparent", border: "none", cursor: "pointer", fontSize: 16, color: "var(--color-text-secondary)", padding: 4 }}>
                <i className="ti ti-x" />
              </button>
            </div>

            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 14 }}>
              {[
                { label: selectedTicket.type, bg: "#E6F1FB", color: "#185FA5" },
                { label: selectedTicket.priority, bg: priorityStyle[selectedTicket.priority]?.bg, color: priorityStyle[selectedTicket.priority]?.color },
                { label: selectedTicket.column, bg: columnStyle[selectedTicket.column]?.bg, color: columnStyle[selectedTicket.column]?.color },
                { label: `${selectedTicket.story_points} pts`, bg: "#EAF3DE", color: "#3B6D11" },
              ].map(b => (
                <span key={b.label} style={{ fontSize: 11, fontWeight: 500, padding: "2px 8px", borderRadius: 20, background: b.bg, color: b.color }}>{b.label}</span>
              ))}
            </div>

            <div style={{ fontSize: 12.5, color: "var(--color-text-secondary)", marginBottom: 6 }}>
              <i className="ti ti-user" style={{ marginRight: 4 }} />{selectedTicket.assignee}
            </div>

            <div style={{ marginBottom: 12 }}>
              <p style={{ margin: "0 0 5px", fontSize: 12, fontWeight: 500, color: "var(--color-text-secondary)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Deskripsi</p>
              <p style={{ margin: 0, fontSize: 13, lineHeight: 1.6, whiteSpace: "pre-wrap" }}>{selectedTicket.description}</p>
            </div>

            {selectedTicket.acceptance_criteria?.length > 0 && (
              <div style={{ marginBottom: 12 }}>
                <p style={{ margin: "0 0 5px", fontSize: 12, fontWeight: 500, color: "var(--color-text-secondary)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Acceptance Criteria</p>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  {selectedTicket.acceptance_criteria.map((ac, i) => (
                    <div key={i} style={{ display: "flex", gap: 8, fontSize: 12.5, lineHeight: 1.5 }}>
                      <i className="ti ti-circle-check" style={{ color: "#3B6D11", fontSize: 13, marginTop: 2, flexShrink: 0 }} />
                      {ac}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* ─── TARUH DI SINI: DECISION AUDIT TRAIL ─── */}
            {selectedTicket.ai_reasoning && (
              <div style={{
                marginBottom: 12, // Sesuaikan jarak bawah dengan elemen lain
                padding: "10px 14px",
                backgroundColor: "#f0fdf4",
                borderLeft: "4px solid #22c55e",
                borderRadius: "6px",
                fontSize: "12.5px", // Samakan ukuran font dengan bagian lain
              }}>
                <p style={{ 
                  margin: "0 0 4px 0", 
                  fontWeight: 600, 
                  color: "#166534", 
                  display: "flex", 
                  alignItems: "center", 
                  gap: "6px",
                  fontSize: 12,
                  textTransform: "uppercase",
                  letterSpacing: "0.05em"
                }}>
                  <span>🤖</span> AI Reasoning (Audit)
                </p>
                <p style={{ margin: 0, color: "#1e4620", lineHeight: "1.5" }}>
                  {selectedTicket.ai_reasoning}
                </p>
              </div>
            )}
            {/* ─────────────────────────────────────────── */}

            

            {/* TAMBAHKAN: Dropdown untuk mengubah status kolom tiket */}
            <div style={{ marginBottom: 15 }}>
              <p style={{ margin: "0 0 5px", fontSize: 12, fontWeight: 500, color: "var(--color-text-secondary)", textTransform: "uppercase", letterSpacing: "0.05em" }}>
                Ubah Status / Kolom
              </p>
              <select
                value={selectedTicket.column}
                onChange={(e) => handleUpdateTicketStatus(selectedTicket.id, e.target.value)}
                style={{
                  width: "100%",
                  padding: "8px 12px",
                  borderRadius: 8,
                  border: "0.5px solid var(--color-border-secondary)",
                  background: "var(--color-background-secondary)",
                  color: "var(--color-text-primary)",
                  fontSize: 13,
                  fontFamily: "inherit",
                  cursor: "pointer",
                  outline: "none"
                }}
              >
                <option value="Backlog">Backlog</option>
                <option value="To Do">To Do</option>
                <option value="In Progress">In Progress</option>
                <option value="Review">Review</option>
                <option value="Done">Done</option>
              </select>
            </div>

            {/* Kode bawaan Anda sebelumnya (Labels) */}
            {selectedTicket.labels?.length > 0 && (
              <div style={{ marginBottom: 15 }}>
                <p style={{ margin: "0 0 5px", fontSize: 12, fontWeight: 500, color: "var(--color-text-secondary)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Label</p>
                <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                  {selectedTicket.labels.map(l => (
                    <span key={l} style={{ fontSize: 11, padding: "2px 8px", borderRadius: 20, background: "var(--color-background-secondary)", color: "var(--color-text-secondary)", border: "0.5px solid var(--color-border-tertiary)" }}>{l}</span>
                  ))}
                </div>
              </div>
            )}

            {selectedTicket.dependencies?.length > 0 && (
              <div>
                <p style={{ margin: "0 0 5px", fontSize: 12, fontWeight: 500, color: "var(--color-text-secondary)", textTransform: "uppercase", letterSpacing: "0.05em" }}>Depends On</p>
                <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                  {selectedTicket.dependencies.map(d => (
                    <span key={d} style={{ fontSize: 11, fontFamily: "'IBM Plex Mono', monospace", padding: "2px 8px", borderRadius: 5, background: "#FAEEDA", color: "#854F0B" }}>{d}</span>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:0.3; } }
        @keyframes loading { 0% { transform: translateX(-100%); } 100% { transform: translateX(250%); } }
      `}</style>
    </div>
  );
}