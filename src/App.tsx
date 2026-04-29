import React, { useState, useRef } from 'react';
import * as pdfjsLib from 'pdfjs-dist';
import * as XLSX from 'xlsx';
import { GoogleGenAI, Type } from "@google/genai";
import { 
  Upload, 
  FileText, 
  Search, 
  Mail, 
  CheckCircle2, 
  AlertCircle, 
  Loader2, 
  ExternalLink,
  ChevronRight,
  ArrowRight
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// Configure PDF.js worker
// We use a CDN to ensure the worker is available without complex build steps in this environment
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

interface AuditResult {
  firstName: string;
  lastName: string;
  candidateId: string;
  email: string;
  subject: string;
  body: string;
  trainingNo?: string;
  trainingName?: string;
  status: 'matched' | 'similar' | 'notfound' | 'manual' | 'duplicate';
  confidence?: number;
  matchOptions?: { name: string; email: string; score: number }[];
}

interface AuditSession {
  id: string;
  name: string;
  userName: string;
  userEmail: string;
  timestamp: Date;
  status: 'draft' | 'completed';
  resultCount: number;
  results: AuditResult[];
  instruction: string;
  ccList?: string;
}

const DELOITTE_GREEN = "#86BC25";

type AppView = 'login' | 'dashboard' | 'audit-step-1' | 'audit-step-2' | 'audit-step-3';

export default function App() {
  const [view, setView] = useState<AppView>('login');
  const [currentUser, setCurrentUser] = useState<{name: string, email: string} | null>(null);
  const [sessions, setSessions] = useState<AuditSession[]>([]);
  const [activeSessionId, setActiveSessionId] = useState<string | null>(null);

  // Automation State
  const [authStatus, setAuthStatus] = useState({ microsoft: false, google: false });
  const [automationSettings, setAutomationSettings] = useState({
    subjectTrigger: localStorage.getItem('audit_subject_trigger') || "Training Report",
    sourceId: localStorage.getItem('audit_sheet_id') || "1a29I0sU51Awvnl2P2Pr8RXuQuftTW2WSKyXnK4yeQb0"
  });
  const [isSyncing, setIsSyncing] = useState(false);

  const [uploadedFiles, setUploadedFiles] = useState<File[]>([]);
  const [instruction, setInstruction] = useState("The source file (File A) is a definitive list of people who have NOT completed training. Extract their First Name, Last Name, Training Name, and Training No. Then, find their email addresses in the Directory (File B) and create nudge emails for them.");
  const [results, setResults] = useState<AuditResult[]>([]);
  const [filter, setFilter] = useState<'all' | 'with-email' | 'no-email'>('all');
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [ccMemory, setCcMemory] = useState("");
  const [isDraftOpen, setIsDraftOpen] = useState(false);
  const [currentDraft, setCurrentDraft] = useState({ subject: "", body: "", cc: "" });

  const baseInputRef = useRef<HTMLInputElement>(null);

  const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || '' });

  // Persistence logic
  React.useEffect(() => {
    const saved = localStorage.getItem('audit_agent_sessions');
    const savedCC = localStorage.getItem('audit_agent_cc_memory');
    if (saved) {
      try {
        setSessions(JSON.parse(saved));
      } catch (e) {
        console.error("Failed to load sessions", e);
      }
    }
    if (savedCC) setCcMemory(savedCC);
    
    checkAuthStatus();
  }, []);

  const checkAuthStatus = async () => {
    try {
      const res = await fetch('/api/status');
      const data = await res.json();
      setAuthStatus(data);
    } catch (e) {
      console.error("Failed to check auth status", e);
    }
  };

  React.useEffect(() => {
    const handleMessage = (event: MessageEvent) => {
      if (event.data?.type === 'MS_AUTH_SUCCESS' || event.data?.type === 'GOOGLE_AUTH_SUCCESS') {
        checkAuthStatus();
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const handleConnect = async (type: 'microsoft' | 'google') => {
    try {
      const res = await fetch(`/api/auth/${type}/url`);
      const { url } = await res.json();
      window.open(url, `${type}_oauth`, 'width=600,height=700');
    } catch (e) {
      setError(`Failed to initiate ${type} connection`);
    }
  };

  React.useEffect(() => {
    localStorage.setItem('audit_subject_trigger', automationSettings.subjectTrigger);
    localStorage.setItem('audit_sheet_id', automationSettings.sourceId);
  }, [automationSettings]);

  React.useEffect(() => {
    if (sessions.length > 0) {
      localStorage.setItem('audit_agent_sessions', JSON.stringify(sessions));
    }
  }, [sessions]);

  React.useEffect(() => {
    localStorage.setItem('audit_agent_cc_memory', ccMemory);
  }, [ccMemory]);

  const handleLogin = (name: string, email: string) => {
    const emailLower = email.toLowerCase().trim();
    setCurrentUser({ name, email: emailLower });
    setView('dashboard');
  };

  const startNewSession = () => {
    const id = Math.random().toString(36).substring(7);
    setActiveSessionId(id);
    setResults([]);
    setUploadedFiles([]);
    setView('audit-step-1');
  };

  const loadSession = (id: string) => {
    const session = sessions.find(s => s.id === id);
    if (session) {
      setActiveSessionId(id);
      setResults(session.results);
      setInstruction(session.instruction);
      setView('audit-step-3');
    }
  };

  const saveCurrentSession = (updatedResults: AuditResult[]) => {
    if (!currentUser || !activeSessionId) return;

    setSessions(prev => {
      const existingIdx = prev.findIndex(s => s.id === activeSessionId);
      const newSession: AuditSession = {
        id: activeSessionId,
        userName: currentUser.name,
        userEmail: currentUser.email,
        name: `Audit Session - ${new Date().toLocaleDateString()}`,
        timestamp: new Date(),
        status: updatedResults.length > 0 ? 'completed' : 'draft',
        resultCount: updatedResults.length,
        results: updatedResults,
        instruction
      };

      if (existingIdx >= 0) {
        const next = [...prev];
        next[existingIdx] = newSession;
        return next;
      }
      return [newSession, ...prev];
    });
  };
  const extractTextFromFile = async (file: File): Promise<string> => {
    const extension = file.name.split('.').pop()?.toLowerCase();

    if (extension === 'pdf') {
      return await extractTextFromPDF(file);
    } else if (extension === 'csv') {
      return await file.text();
    } else if (extension === 'xlsx' || extension === 'xls') {
      return await extractTextFromExcel(file);
    } else {
      throw new Error(`Unsupported file type: ${extension}`);
    }
  };

  const extractTextFromPDF = async (file: File): Promise<string> => {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
      const pdf = await loadingTask.promise;
      let fullText = "";
      
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const pageText = textContent.items
          .map((item: any) => item.str)
          .join(" ");
        fullText += pageText + "\n";
      }
      return fullText;
    } catch (err) {
      console.error("PDF Extraction Error:", err);
      throw new Error(`Failed to extract text from ${file.name}`);
    }
  };

  const extractTextFromExcel = async (file: File): Promise<string> => {
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      let fullText = "";
      
      workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const csv = XLSX.utils.sheet_to_csv(worksheet);
        fullText += `Sheet: ${sheetName}\n${csv}\n\n`;
      });
      
      return fullText;
    } catch (err) {
      console.error("Excel Extraction Error:", err);
      throw new Error(`Failed to extract text from Excel file: ${file.name}`);
    }
  };

  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        const base64String = reader.result?.toString().split(',')[1];
        if (base64String) resolve(base64String);
        else reject(new Error("Failed to convert file to base64"));
      };
      reader.onerror = (error) => reject(error);
    });
  };

  const updateResultEmail = (idx: number, email: string) => {
    setResults(prev => {
      const next = [...prev];
      next[idx] = { 
        ...next[idx], 
        email, 
        status: email ? 'manual' : 'notfound' 
      };
      saveCurrentSession(next);
      return next;
    });
  };

  const handleMailto = (res: AuditResult) => {
    const subject = encodeURIComponent(res.subject);
    const body = encodeURIComponent(res.body);
    window.location.href = `mailto:${res.email}?subject=${subject}&body=${body}`;
  };

  const handleMasterDraft = () => {
    if (results.length === 0) return;
    const targets = results.filter(r => r.email);
    
    // Create a plain text table for the email body
    const tableHeader = "NAME              | EMAIL                         | TRAINING ID | TRAINING NAME       | STATUS\n";
    const tableSeparator = "------------------|-------------------------------|-------------|---------------------|---------------\n";
    let tableRows = "";
    
    targets.forEach(r => {
      const fullName = `${r.firstName} ${r.lastName}`;
      const name = fullName.padEnd(17).slice(0, 17);
      const email = r.email.padEnd(29).slice(0, 29);
      const tNo = (r.trainingNo || 'N/A').padEnd(11).slice(0, 11);
      const tName = (r.trainingName || 'N/A').padEnd(19).slice(0, 19);
      tableRows += `${name} | ${email} | ${tNo} | ${tName} | Not Completed\n`;
    });

    const body = `Dear All,

Please find the reconciliation for pending training completion items below:

${tableHeader}${tableSeparator}${tableRows}

Please prioritize these training items as they are now overdue.

Regards,
Audit Team`;

    setCurrentDraft({ 
      subject: "ACTION REQUIRED: Training Completion Reconciliation Nudge", 
      body, 
      cc: ccMemory 
    });
    setIsDraftOpen(true);
  };

  const handleSyncAndTrigger = async () => {
    if (!authStatus.google) {
      setError("Please connect Google Sheets first. This is required for the Directory repository.");
      handleConnect('google');
      return;
    }
    if (!automationSettings.sourceId) {
      setError("Please provide the Source of Truth Sheet ID.");
      return;
    }
    if (uploadedFiles.length === 0) {
      setError("Please upload at least one training file (PDF/Excel) to begin.");
      return;
    }

    setIsSyncing(true);
    setIsProcessing(true);
    setError(null);
    setResults([]);
    setView('audit-step-2');

    try {
      // 1. Fetch Directory from Google Sheets
      const sheetRes = await fetch(`/api/sheets/fetch?sheetId=${automationSettings.sourceId}`);
      if (!sheetRes.ok) throw new Error("Google Sheets sync failed: " + (await sheetRes.json()).error);
      const sheetData = await sheetRes.json();
      const directoryText = `[DIRECTORY REPOSITORY FROM GOOGLE SHEETS]:\n${JSON.stringify(sheetData.values)}`;

      // 2. Prepare Uploaded Files
      const fileParts = await Promise.all(uploadedFiles.map(async (file) => {
        const extension = file.name.split('.').pop()?.toLowerCase();
        if (extension === 'pdf') {
          const b64 = await fileToBase64(file);
          return { inlineData: { data: b64, mimeType: "application/pdf" } };
        } else {
          const text = await extractTextFromFile(file);
          return { text: `[DATA FROM UPLOADED ${file.name}]:\n${text.slice(0, 30000)}` };
        }
      }));

      // 3. Run Audit with Gemini
      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [
          {
            role: "user",
            parts: [
              { text: "You are a Deloitte Audit Assistant. I am providing you with one or more Source files (uploaded by the user) and a Directory Repository from Google Sheets." },
              ...fileParts as any,
              { text: directoryText },
              {
                text: `Your task is as follows:
                1. EXHAUSTIVE EXTRACTION: Go through ALL provided Source Files. You MUST extract "First Name", "Last Name", "Candidate ID" (or Employee ID), "Training No", and "Training Name" for EVERY SINGLE PERSON mentioned as not having completed training.
                2. MATCHING: Cross-reference every person against the Directory Repository to find their email address. Match using First Name, Last Name, and Candidate ID.
                3. FUZZY MATCH: If names are slightly different but Candidate ID matches, consider it a match.
                4. STATUS: 
                   - 'matched' if you found an email.
                   - 'notfound' if no email is found in the directory.
                5. REQUIRED OUTPUT (JSON Array of objects):
                   - firstName: string
                   - lastName: string
                   - candidateId: string
                   - email: string (empty string if not found)
                   - trainingNo: string
                   - trainingName: string
                   - subject: string (Reminder subject)
                   - body: string (Nudge message)
                   - status: "matched" | "notfound"
                
                Additional Context: "${instruction}".`
              }
            ]
          }
        ],
        config: {
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                firstName: { type: Type.STRING },
                lastName: { type: Type.STRING },
                candidateId: { type: Type.STRING },
                email: { type: Type.STRING },
                trainingNo: { type: Type.STRING },
                trainingName: { type: Type.STRING },
                subject: { type: Type.STRING },
                body: { type: Type.STRING },
                status: { type: Type.STRING }
              },
              required: ["firstName", "lastName", "candidateId", "email", "status"]
            }
          }
        }
      });

      const parsedResults: AuditResult[] = JSON.parse(response.text || "[]");
      setResults(parsedResults);
      saveCurrentSession(parsedResults);
      setView('audit-step-3');
    } catch (err: any) {
      console.error("Automation Error:", err);
      setError(err.message);
      setView('audit-step-1');
    } finally {
      setIsSyncing(false);
      setIsProcessing(false);
    }
  };

  const finalizeDraft = async () => {
    const to = results.filter(r => r.email).map(r => r.email).join('; ');
    if (!to) return;

    try {
      setIsProcessing(true);
      const res = await fetch('/api/outlook/send', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          to,
          cc: currentDraft.cc,
          subject: currentDraft.subject,
          body: currentDraft.body
        })
      });

      if (!res.ok) throw new Error("Failed to send via Outlook API");
      
      setIsDraftOpen(false);
      alert("Success! Audit reconciliation dispatched via Outlook.");
    } catch (err: any) {
      setError(err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="h-screen bg-slate-50 flex flex-col font-sans text-slate-900 overflow-hidden">
      {/* Dynamic Header */}
      <header className="h-16 bg-black text-white flex items-center justify-between px-8 border-b-4 border-deloitte shrink-0">
        <div className="flex items-center gap-4">
          <div className="text-2xl font-black tracking-tighter text-white">
            Deloitte<span className="text-deloitte">.</span>
          </div>
          <div className="h-6 w-px bg-slate-700"></div>
          <div className="text-sm font-bold tracking-widest uppercase text-slate-400">AuditAgent AI</div>
        </div>
        
        {currentUser && (
          <div className="flex items-center gap-6">
            <div className="hidden md:flex gap-4 items-center">
              <button 
                onClick={() => setView('dashboard')}
                className={`text-[10px] font-black uppercase tracking-widest px-2 ${view === 'dashboard' ? 'text-deloitte' : 'text-slate-400 hover:text-white'}`}
              >
                Projects
              </button>
              {activeSessionId && (
                <>
                  <ChevronRight className="w-3 h-3 text-slate-700" />
                  <button 
                    onClick={() => setView('audit-step-3')}
                    className={`text-[10px] font-black uppercase tracking-widest px-2 ${view.startsWith('audit-step') ? 'text-deloitte' : 'text-slate-400 hover:text-white'}`}
                  >
                    Active Audit
                  </button>
                </>
              )}
            </div>
            <div className="flex items-center gap-3">
              <div className="w-8 h-8 rounded-full bg-deloitte flex items-center justify-center text-black font-black uppercase text-[10px]">
                {currentUser.name.split(' ').map(n => n[0]).join('')}
              </div>
              <div className="flex flex-col items-start leading-none">
                <span className="text-[10px] font-black uppercase tracking-widest text-white">{currentUser.name}</span>
                <span className="text-[9px] font-bold text-slate-500 uppercase tracking-tighter">{currentUser.email}</span>
              </div>
            </div>
          </div>
        )}
      </header>

      <AnimatePresence mode="wait">
        {view === 'login' && (
          <motion.div 
            key="login"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="flex-1 flex"
          >
            {/* Left side: Product Intro */}
            <div className="w-1/2 bg-black p-20 flex flex-col justify-center text-white border-r border-slate-800">
               <motion.div
                initial={{ x: -20, opacity: 0 }}
                animate={{ x: 0, opacity: 1 }}
                transition={{ delay: 0.2 }}
              >
                <h1 className="text-8xl font-black tracking-tighter leading-[0.85] mb-8 uppercase">
                  Reconcile<br />Everything<span className="text-deloitte">.</span>
                </h1>
                <p className="text-slate-400 text-xl max-w-md font-medium leading-relaxed italic border-l-2 border-deloitte pl-6">
                  "Precision is not an option; it's our standard." 
                  <span className="block mt-4 text-xs font-black uppercase tracking-widest not-italic">Deloitte Audit Services</span>
                </p>

                <div className="mt-16 grid grid-cols-2 gap-8 text-[10px] font-black uppercase tracking-widest">
                  <div className="p-6 border border-slate-800 rounded-xl bg-slate-900/50">
                    <div className="text-deloitte mb-4"><Search className="w-6 h-6" /></div>
                    Exhaustive<br />Extraction
                  </div>
                  <div className="p-6 border border-slate-800 rounded-xl bg-slate-900/50">
                    <div className="text-deloitte mb-4"><Mail className="w-6 h-6" /></div>
                    Automated<br />Nudges
                  </div>
                </div>
              </motion.div>
            </div>

            {/* Right side: User ID */}
            <div className="w-1/2 bg-slate-100 flex flex-col items-center justify-center p-20">
              <div className="w-full max-w-md">
                <h2 className="text-4xl font-black tracking-tighter uppercase mb-2">Identify Yourself<span className="text-deloitte">.</span></h2>
                <p className="text-slate-500 text-sm font-bold uppercase tracking-widest mb-10">Access your historical audit sessions.</p>
                
                <form 
                  onSubmit={(e) => {
                    e.preventDefault();
                    const formData = new FormData(e.currentTarget);
                    handleLogin(formData.get('name') as string, formData.get('email') as string);
                  }}
                  className="space-y-6"
                >
                  <div className="space-y-1">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">Full Name</label>
                    <input 
                      name="name"
                      required
                      placeholder="e.g. John Doe"
                      className="w-full h-16 bg-white border border-slate-200 rounded-lg px-6 font-black uppercase tracking-widest text-sm focus:outline-none focus:border-deloitte transition-all"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">Deloitte ID (Email)</label>
                    <input 
                      name="email"
                      type="email"
                      required
                      placeholder="jdoe@deloitte.com"
                      className="w-full h-16 bg-white border border-slate-200 rounded-lg px-6 font-black uppercase tracking-widest text-sm focus:outline-none focus:border-deloitte transition-all"
                    />
                  </div>
                  <button 
                    type="submit"
                    className="w-full h-16 bg-black text-white rounded-lg font-black uppercase tracking-widest hover:bg-deloitte hover:text-black transition-all group flex items-center justify-center gap-3 active:scale-95"
                  >
                    Initialize Environment
                    <ArrowRight className="w-5 h-5 group-hover:translate-x-2 transition-transform" />
                  </button>
                </form>
              </div>
            </div>
          </motion.div>
        )}

        {view === 'dashboard' && currentUser && (
          <motion.div 
            key="dashboard"
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            className="flex-1 p-20 flex flex-col gap-12 overflow-y-auto"
          >
            <div className="flex justify-between items-end">
              <div>
                <h1 className="text-7xl font-black tracking-tighter leading-none uppercase">Project<br />Directory<span className="text-deloitte">.</span></h1>
                <p className="text-slate-400 font-black uppercase text-sm tracking-[0.3em] mt-4">Welcome back, Agent {currentUser.name.split(' ')[0]}</p>
              </div>
              <div className="flex gap-4">
                <div className="flex flex-col gap-1">
                  <label className="text-[9px] font-black uppercase tracking-widest text-slate-400">CC Persistence Memory</label>
                  <input 
                    value={ccMemory}
                    onChange={(e) => setCcMemory(e.target.value)}
                    placeholder="manager@deloitte.com; hr@deloitte.com"
                    className="w-64 h-12 bg-white border border-slate-200 rounded-xl px-4 text-[10px] font-bold uppercase tracking-widest focus:border-deloitte focus:outline-none shadow-sm"
                  />
                </div>
                <button 
                  onClick={startNewSession}
                  className="px-10 h-12 mt-auto bg-black text-white rounded-xl text-xs font-black uppercase tracking-[0.2em] hover:bg-deloitte hover:text-black transition-all shadow-2xl active:scale-95 flex items-center gap-3"
                >
                  <ArrowRight className="w-4 h-4" />
                  Initialize New Audit
                </button>
              </div>
            </div>

            <div className="bg-white border border-slate-200 rounded-3xl overflow-hidden shadow-xl">
              <div className="p-8 border-b border-slate-100 flex justify-between items-center">
                <h3 className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-400">Historical Session History</h3>
                <span className="px-3 py-1 bg-slate-100 rounded text-[9px] font-black uppercase tracking-widest">
                  {sessions.filter(s => s.userEmail === currentUser.email).length} Total Sessions
                </span>
              </div>
              
              <div className="divide-y divide-slate-100">
                {sessions.filter(s => s.userEmail === currentUser.email).length === 0 ? (
                  <div className="p-20 text-center text-slate-300">
                    <FileText className="w-16 h-16 mx-auto mb-4 opacity-10" />
                    <p className="font-black text-xl tracking-widest uppercase">No Prior Sessions Detected</p>
                    <p className="text-[10px] font-black uppercase tracking-widest mt-2">Start a new session to begin auditing.</p>
                  </div>
                ) : (
                  sessions.filter(s => s.userEmail === currentUser.email).map(s => (
                    <div 
                      key={s.id}
                      className="group p-8 flex items-center hover:bg-slate-50 transition-all cursor-pointer"
                      onClick={() => loadSession(s.id)}
                    >
                      <div className="w-12 h-12 rounded-xl bg-slate-100 flex items-center justify-center text-slate-400 group-hover:bg-deloitte group-hover:text-black transition-colors shrink-0">
                        <FileText className="w-6 h-6" />
                      </div>
                      <div className="ml-6 flex-1">
                        <h4 className="text-xl font-black tracking-tighter uppercase">{s.name}</h4>
                        <div className="flex gap-4 mt-1">
                          <span className="text-[9px] font-black uppercase tracking-widest text-slate-400">Created: {new Date(s.timestamp).toLocaleDateString()}</span>
                          <span className="text-[9px] font-black uppercase tracking-widest text-slate-400">{s.resultCount} Targets Audited</span>
                        </div>
                      </div>
                      <div className="flex items-center gap-4">
                        <span className={`px-2 py-1 rounded text-[8px] font-black uppercase tracking-widest ${s.status === 'completed' ? 'bg-[#E6F4D7] text-black' : 'bg-slate-200 text-slate-600'}`}>
                          {s.status}
                        </span>
                        <ChevronRight className="w-5 h-5 text-slate-300 group-hover:text-black group-hover:translate-x-1 transition-all" />
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          </motion.div>
        )}

        {view === 'audit-step-1' && (
          <motion.div 
            key="audit-step-1"
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: -20 }}
            className="flex-1 flex flex-col items-center justify-center p-12 overflow-y-auto"
          >
            <div className="w-full max-w-4xl grid grid-cols-1 md:grid-cols-2 gap-12">
              <div className="space-y-8">
                <div>
                  <h2 className="text-6xl font-black tracking-tighter uppercase leading-none mb-4">Step 1<br />Ingestion<span className="text-deloitte">.</span></h2>
                  <p className="text-slate-500 font-bold uppercase tracking-widest text-xs leading-relaxed">
                    Upload the training report (PDF or Excel). This will be cross-referenced against our global Directory Repository.
                  </p>
                </div>

                <div className="p-8 bg-black rounded-3xl border border-slate-800 space-y-6">
                  <h3 className="text-xs font-black uppercase tracking-widest text-deloitte flex items-center gap-2">
                    <div className="w-2 h-2 bg-deloitte rounded-full"></div>
                    Directory Repository
                  </h3>
                  <div>
                    <label className="text-[8px] font-black uppercase tracking-widest text-slate-500 block mb-2">Connected Source of Truth (Google Sheets)</label>
                    <div className="flex gap-2">
                      <input 
                        value={automationSettings.sourceId}
                        onChange={(e) => setAutomationSettings(s => ({ ...s, sourceId: e.target.value }))}
                        className="flex-1 h-12 bg-slate-900 border border-slate-800 rounded-xl px-4 text-[11px] font-mono text-deloitte focus:border-deloitte focus:outline-none"
                        placeholder="Google Sheet ID"
                      />
                      {!authStatus.google && (
                        <button 
                          onClick={() => handleConnect('google')}
                          className="h-12 px-4 bg-slate-800 text-white rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-white hover:text-black transition-all"
                        >
                          Link
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              <div className="space-y-6">
                <div 
                  onClick={() => baseInputRef.current?.click()}
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={(e) => {
                    e.preventDefault();
                    const files = Array.from(e.dataTransfer.files);
                    setUploadedFiles(prev => [...prev, ...files]);
                  }}
                  className="border-4 border-dashed border-slate-200 rounded-3xl p-16 flex flex-col items-center justify-center gap-6 hover:border-deloitte hover:bg-slate-50 transition-all cursor-pointer group relative bg-white shadow-sm"
                >
                  <div className="w-20 h-20 bg-slate-50 rounded-full flex items-center justify-center group-hover:scale-110 transition-transform">
                    <Upload className="w-10 h-10 text-deloitte" />
                  </div>
                  <div className="text-center">
                    <p className="text-lg font-black uppercase tracking-widest">Drop Training File</p>
                    <p className="text-[10px] font-bold text-slate-400 mt-2 uppercase tracking-widest">Supports PDF, XLSX, CSV</p>
                  </div>
                  <input type="file" multiple ref={baseInputRef} className="hidden" onChange={(e) => {
                    const files = Array.from(e.target.files || []) as File[];
                    setUploadedFiles(prev => [...prev, ...files]);
                  }} />
                </div>

                {uploadedFiles.length > 0 && (
                  <div className="space-y-2">
                    {uploadedFiles.map((f, i) => (
                      <motion.div 
                        key={i} 
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        className="p-4 bg-white border border-slate-200 rounded-2xl flex items-center justify-between"
                      >
                        <div className="flex items-center gap-3 overflow-hidden">
                          <FileText className="w-5 h-5 text-deloitte shrink-0" />
                          <span className="text-xs font-bold truncate">{f.name}</span>
                        </div>
                        <button 
                          onClick={() => setUploadedFiles(prev => prev.filter((_, idx) => idx !== i))}
                          className="w-8 h-8 flex items-center justify-center text-slate-400 hover:text-red-500 transition-colors"
                        >
                          ×
                        </button>
                      </motion.div>
                    ))}
                    
                    <button 
                      onClick={handleSyncAndTrigger}
                      disabled={isSyncing || uploadedFiles.length === 0}
                      className="w-full h-20 bg-deloitte text-black rounded-3xl font-black uppercase tracking-[0.4em] hover:bg-black hover:text-white transition-all shadow-2xl mt-4 flex items-center justify-center gap-4 group active:scale-95"
                    >
                      {isSyncing ? <Loader2 className="w-6 h-6 animate-spin" /> : "Run Automated Audit"}
                      {!isSyncing && <ArrowRight className="w-6 h-6 group-hover:translate-x-2 transition-transform" />}
                    </button>
                  </div>
                )}
              </div>
            </div>
            {error && (
              <motion.div 
                initial={{ opacity: 0, y: 10 }}
                animate={{ opacity: 1, y: 0 }}
                className="mt-8 p-4 bg-red-50 border border-red-100 rounded-2xl flex gap-3 text-red-600 max-w-4xl w-full"
              >
                <AlertCircle className="w-4 h-4 shrink-0" />
                <span className="text-[10px] font-black uppercase tracking-widest">{error}</span>
              </motion.div>
            )}
          </motion.div>
        )}

        {view === 'audit-step-2' && (
          <motion.div 
            key="audit-step-2"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="flex-1 flex flex-col items-center justify-center bg-black text-white"
          >
            <div className="relative w-64 h-64 mb-12">
              <motion.div 
                animate={{ rotate: 360 }}
                transition={{ duration: 4, repeat: Infinity, ease: "linear" }}
                className="absolute inset-0 border-[16px] border-slate-900 border-t-deloitte rounded-full shadow-[0_0_50px_rgba(134,188,37,0.3)]"
              />
              <div className="absolute inset-0 flex items-center justify-center">
                <Search className="w-20 h-20 text-deloitte animate-pulse" />
              </div>
            </div>
            
            <div className="text-center space-y-4">
              <h2 className="text-5xl font-black tracking-tighter uppercase leading-none">Step 2<br />Reconciliation<span className="text-deloitte">.</span></h2>
              <div className="flex flex-col items-center gap-2">
                <p className="text-slate-400 font-black uppercase text-[10px] tracking-[0.4em] animate-pulse">Cross-Referencing Repository</p>
                <div className="flex gap-1">
                  {[0, 1, 2].map(i => (
                    <motion.div 
                      key={i}
                      animate={{ opacity: [0.2, 1, 0.2] }}
                      transition={{ duration: 1, repeat: Infinity, delay: i * 0.2 }}
                      className="w-1.5 h-1.5 bg-deloitte rounded-full"
                    />
                  ))}
                </div>
              </div>
            </div>
            
            <div className="mt-20 max-w-lg w-full px-8">
              <div className="h-1 bg-slate-900 rounded-full overflow-hidden">
                <motion.div 
                  initial={{ width: "0%" }}
                  animate={{ width: "100%" }}
                  transition={{ duration: 8, ease: "easeInOut" }}
                  className="h-full bg-deloitte"
                />
              </div>
              <p className="text-[8px] font-black uppercase text-slate-600 mt-4 tracking-widest flex justify-between">
                <span>Ingesting Source Data</span>
                <span>Gemini 3 Flash Environment</span>
              </p>
            </div>
          </motion.div>
        )}

        {view === 'audit-step-3' && (
          <motion.div 
            key="audit-step-3"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="flex-1 flex flex-col overflow-hidden"
          >
            <div className="bg-white border-b border-slate-200 p-8 flex justify-between items-end">
              <div>
                <h2 className="text-6xl font-black tracking-tighter uppercase leading-none mb-2">Step 3<br />Results<span className="text-deloitte">.</span></h2>
                <p className="text-slate-400 font-black uppercase text-[10px] tracking-[0.3em]">
                  {results.length} INDIVIDUALS IDENTIFIED · {results.filter(r => r.status === 'matched').length} MATCHED
                </p>
              </div>

              <div className="flex items-center gap-4">
                <div className="flex gap-2 p-1 bg-slate-100 rounded-xl">
                  {[
                    { id: 'all', label: 'All' },
                    { id: 'with-email', label: 'Matched' },
                    { id: 'no-email', label: 'Missing' }
                  ].map((f) => (
                    <button
                      key={f.id}
                      onClick={() => setFilter(f.id as any)}
                      className={`px-4 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all ${
                        filter === f.id 
                        ? 'bg-black text-white shadow-lg' 
                        : 'text-slate-500 hover:bg-white'
                      }`}
                    >
                      {f.label}
                    </button>
                  ))}
                </div>
                <button 
                  onClick={handleMasterDraft}
                  className="h-14 px-8 bg-deloitte text-black rounded-xl text-xs font-black uppercase tracking-[0.2em] hover:bg-black hover:text-white transition-all shadow-xl flex items-center gap-3 active:scale-95"
                >
                  <Mail className="w-5 h-5" />
                  Bulk Dispatch (Outlook)
                </button>
                <button 
                  onClick={() => setView('audit-step-1')}
                  className="h-14 px-6 border border-slate-200 rounded-xl text-xs font-black uppercase tracking-widest hover:bg-slate-50 transition-all text-slate-400"
                >
                  New Audit
                </button>
              </div>
            </div>

            <div className="flex-1 bg-slate-50 overflow-y-auto p-8 pt-0">
              <div className="max-w-6xl mx-auto py-8">
                <div className="grid grid-cols-12 px-8 py-4 text-[9px] font-black uppercase tracking-widest text-slate-400 shrink-0 mb-4 bg-transparent">
                  <div className="col-span-1">ID</div>
                  <div className="col-span-4">Audit Target</div>
                  <div className="col-span-4">Repository Match</div>
                  <div className="col-span-3 text-right">Dispatch Status</div>
                </div>

                <div className="space-y-4">
                  {results
                    .filter(res => {
                      if (filter === 'with-email') return res.status === 'matched' || res.status === 'manual' || !!res.email;
                      if (filter === 'no-email') return res.status === 'notfound' && !res.email;
                      return true;
                    })
                    .map((res, idx) => (
                    <motion.div 
                      key={idx}
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      className="bg-white border border-slate-200 rounded-3xl p-8 grid grid-cols-12 items-center hover:shadow-xl transition-all group"
                    >
                      <div className="col-span-1 text-slate-300 font-black font-mono text-xs">{res.candidateId || 'N/A'}</div>
                      <div className="col-span-4">
                        <p className="font-black text-2xl tracking-tighter leading-none mb-1 uppercase">{res.firstName} {res.lastName}</p>
                        <div className="flex gap-2">
                          <span className="px-2 py-1 bg-slate-100 text-slate-500 rounded text-[7px] font-black uppercase tracking-widest leading-none">
                            {res.trainingName || 'COURSE PENDING'}
                          </span>
                        </div>
                      </div>
                      <div className="col-span-4 pr-12">
                        {res.email ? (
                          <div className="flex flex-col">
                            <p className="font-mono text-sm text-slate-900 font-bold tracking-tight lowercase">{res.email}</p>
                            <span className="text-[8px] font-black uppercase text-deloitte mt-1">Verified Match</span>
                          </div>
                        ) : (
                          <div className="flex flex-col gap-2">
                            <div className="flex items-center gap-2 text-red-500">
                              <AlertCircle className="w-3 h-3" />
                              <span className="text-[8px] font-black uppercase tracking-widest">Not Found in Repository</span>
                            </div>
                            <input 
                              type="email" 
                              placeholder="MANUAL ENTRY..."
                              className="w-full h-10 bg-slate-50 border border-slate-200 rounded-lg px-4 text-[10px] font-black uppercase tracking-wider focus:bg-white focus:border-deloitte focus:outline-none transition-all placeholder:text-slate-400"
                              onBlur={(e) => e.target.value && updateResultEmail(idx, e.target.value)}
                            />
                          </div>
                        )}
                      </div>
                      <div className="col-span-3 flex justify-end gap-3">
                         <button 
                            onClick={() => handleMailto(res)}
                            disabled={!res.email}
                            className="h-14 px-8 bg-black text-white text-[10px] font-black uppercase tracking-widest rounded-2xl shadow-lg hover:bg-deloitte hover:text-black transition-all flex items-center gap-2 disabled:opacity-5 group/btn"
                          >
                            <Mail className="w-4 h-4" />
                            Dispatch Nudge
                          </button>
                      </div>
                    </motion.div>
                  ))}

                  {results.length === 0 && (
                     <div className="h-64 flex flex-col items-center justify-center text-slate-300 border-2 border-dashed border-slate-200 rounded-3xl">
                        <Search className="w-12 h-12 mb-4 opacity-10" />
                        <p className="font-black text-lg tracking-[0.2em] uppercase">No Match Identified</p>
                     </div>
                  )}
                </div>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Overlays / Modals */}
      <AnimatePresence>
        {isDraftOpen && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[100] bg-black/80 backdrop-blur-md flex items-center justify-center p-8"
          >
              <motion.div 
                initial={{ scale: 0.9, y: 20 }}
                animate={{ scale: 1, y: 0 }}
                className="bg-white rounded-3xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col shadow-2xl"
              >
                <div className="p-8 border-b border-slate-100 bg-slate-50 flex justify-between items-center">
                  <div>
                    <h3 className="text-3xl font-black uppercase tracking-tighter">Master Draft Preview</h3>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mt-1">Review and finalize the reconciliation nudge.</p>
                  </div>
                  <button onClick={() => setIsDraftOpen(false)} className="w-10 h-10 rounded-full hover:bg-slate-200 flex items-center justify-center transition-colors">
                    <AlertCircle className="w-6 h-6 rotate-45" />
                  </button>
                </div>

                <div className="p-8 flex-1 overflow-y-auto space-y-6">
                  <div className="space-y-4">
                    <div className="grid grid-cols-6 gap-4 items-center">
                      <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">Recipient (To)</label>
                      <div className="col-span-5 p-4 bg-slate-100 rounded-xl text-[10px] font-mono text-slate-600 break-all">
                        {results.filter(r => r.email).map(r => r.email).join('; ')}
                      </div>
                    </div>
                    <div className="grid grid-cols-6 gap-4 items-center">
                      <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">CC List</label>
                      <input 
                        value={currentDraft.cc}
                        onChange={(e) => setCurrentDraft(prev => ({ ...prev, cc: e.target.value }))}
                        className="col-span-5 h-12 bg-white border border-slate-200 rounded-xl px-4 text-[11px] font-bold uppercase tracking-widest focus:border-deloitte focus:outline-none"
                        placeholder="Add CC recipients..."
                      />
                    </div>
                    <div className="grid grid-cols-6 gap-4 items-center">
                      <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">Subject</label>
                      <input 
                        value={currentDraft.subject}
                        onChange={(e) => setCurrentDraft(prev => ({ ...prev, subject: e.target.value }))}
                        className="col-span-5 h-12 bg-white border border-slate-200 rounded-xl px-4 text-[11px] font-black uppercase tracking-widest focus:border-deloitte focus:outline-none"
                      />
                    </div>
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">Message Body</label>
                    <textarea 
                      value={currentDraft.body}
                      onChange={(e) => setCurrentDraft(prev => ({ ...prev, body: e.target.value }))}
                      className="w-full h-80 bg-slate-50 border border-slate-200 rounded-2xl p-8 text-sm font-medium leading-relaxed focus:outline-none focus:border-deloitte transition-all resize-none"
                    />
                  </div>
                </div>

                <div className="p-8 border-t border-slate-100 bg-slate-50 flex gap-4">
                  <button 
                    onClick={() => setIsDraftOpen(false)}
                    className="flex-1 h-16 border border-slate-200 rounded-xl text-xs font-black uppercase tracking-widest hover:bg-white transition-all"
                  >
                    Cancel Draft
                  </button>
                  <button 
                    onClick={finalizeDraft}
                    className="flex-[2] h-16 bg-deloitte text-black rounded-xl text-xs font-black uppercase tracking-widest hover:bg-black hover:text-white transition-all shadow-xl flex items-center justify-center gap-3"
                  >
                    <ExternalLink className="w-5 h-5" />
                    Dispatch to Outlook
                  </button>
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

      <footer className="h-12 bg-black flex items-center px-8 border-t border-[#1E1E1E] shrink-0">
        <div className="flex gap-6 text-[9px] font-black text-slate-500 uppercase tracking-[0.3em] w-full items-center">
          <span className="flex items-center gap-1.5">
            <div className={`w-1.5 h-1.5 rounded-full ${view === 'login' ? 'bg-slate-700' : 'bg-deloitte animate-pulse'}`}></div>
            Node: PPR-882-AUDIT
          </span>
          <span className="opacity-20">|</span>
          <span className="text-slate-400 group flex items-center gap-2">
            Status: {view === 'login' ? 'Identification Required' : 'System Operational'}
          </span>
          <div className="ml-auto text-deloitte font-black tracking-[0.5em]">INTERNAL USE ONLY</div>
        </div>
      </footer>
    </div>
  );
}
