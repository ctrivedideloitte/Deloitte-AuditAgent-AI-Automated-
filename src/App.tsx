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
  name: string;
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

type AppView = 'login' | 'dashboard' | 'audit';

export default function App() {
  const [view, setView] = useState<AppView>('login');
  const [currentUser, setCurrentUser] = useState<{name: string, email: string} | null>(null);
  const [sessions, setSessions] = useState<AuditSession[]>([]);
  const [activeSessionId, setActiveSessionId] = useState<string | null>(null);

  // Automation State
  const [authStatus, setAuthStatus] = useState({ microsoft: false, google: false });
  const [automationSettings, setAutomationSettings] = useState({
    subjectTrigger: localStorage.getItem('audit_subject_trigger') || "Training Report",
    sourceId: localStorage.getItem('audit_sheet_id') || "" // Migrate existing key
  });
  const [isSyncing, setIsSyncing] = useState(false);

  const [baseFiles, setBaseFiles] = useState<File[]>([]);
  const [targetFiles, setTargetFiles] = useState<File[]>([]);
  const [instruction, setInstruction] = useState("The source file (File A) is a definitive list of people who have NOT completed training. Extract their First Name, Last Name, Training Name, and Training No. Then, find their email addresses in the Directory (File B) and create nudge emails for them.");
  const [results, setResults] = useState<AuditResult[]>([]);
  const [filter, setFilter] = useState<'all' | 'with-email' | 'no-email'>('all');
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [ccMemory, setCcMemory] = useState("");
  const [isDraftOpen, setIsDraftOpen] = useState(false);
  const [currentDraft, setCurrentDraft] = useState({ subject: "", body: "", cc: "" });
  const [pdfToConvert, setPdfToConvert] = useState<File | null>(null);
  const [isConverting, setIsConverting] = useState(false);

  const baseInputRef = useRef<HTMLInputElement>(null);
  const addBaseInputRef = useRef<HTMLInputElement>(null);
  const targetInputRef = useRef<HTMLInputElement>(null);
  const addTargetInputRef = useRef<HTMLInputElement>(null);

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
    setCurrentUser({ name, email });
    setView('dashboard');
  };

  const startNewSession = () => {
    const id = Math.random().toString(36).substring(7);
    setActiveSessionId(id);
    setResults([]);
    setBaseFiles([]);
    setTargetFiles([]);
    setView('audit');
  };

  const loadSession = (id: string) => {
    const session = sessions.find(s => s.id === id);
    if (session) {
      setActiveSessionId(id);
      setResults(session.results);
      setInstruction(session.instruction);
      setView('audit');
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

  const convertPDFToExcel = async (file: File) => {
    setIsConverting(true);
    try {
      const b64 = await fileToBase64(file);
      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [{
          role: "user",
          parts: [
            { inlineData: { data: b64, mimeType: "application/pdf" } },
            { text: "Extract all individuals from this training report. For each person, find their Name, Training ID (if any), and the Person ID/Employee ID. Return as a JSON array of objects with keys: name, trainingId, personId." }
          ]
        }],
        config: {
          responseMimeType: "application/json",
          responseSchema: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                name: { type: Type.STRING },
                trainingId: { type: Type.STRING },
                personId: { type: Type.STRING },
              }
            }
          }
        }
      });

      const data = JSON.parse(response.text || "[]");
      const worksheet = XLSX.utils.json_to_sheet(data);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "ExtractedData");
      
      const fileName = file.name.replace('.pdf', '_converted.xlsx');
      XLSX.writeFile(workbook, fileName);
      
      // Also add it to our base files
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const excelBlob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const convertedFile = new File([excelBlob], fileName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      setBaseFiles(prev => [...prev, convertedFile]);
      setPdfToConvert(null);
    } catch (err: any) {
      console.error("Conversion Error:", err);
      setError("Failed to convert PDF. Please try again.");
    } finally {
      setIsConverting(false);
    }
  };

  const handleProcess = async () => {
    if (baseFiles.length === 0 || targetFiles.length === 0) {
      setError("Please upload at least one file for both Source and Target.");
      return;
    }

    setIsProcessing(true);
    setError(null);
    setResults([]);

    try {
      const prepareFileParts = async (files: File[]) => {
        return Promise.all(files.map(async (file) => {
          const extension = file.name.split('.').pop()?.toLowerCase();
          if (extension === 'pdf') {
            const b64 = await fileToBase64(file);
            return { inlineData: { data: b64, mimeType: "application/pdf" } };
          } else {
            const text = await extractTextFromFile(file);
            return { text: `[DATA FROM ${file.name}]:\n${text.slice(0, 20000)}` };
          }
        }));
      };

      const [baseParts, targetParts] = await Promise.all([
        prepareFileParts(baseFiles),
        prepareFileParts(targetFiles)
      ]);

      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [
          {
            role: "user",
            parts: [
              { text: "You are a Deloitte Audit Assistant. I am providing you with multiple Source (File A) and Target (File B) documents. Some are PDFs, others are CSV/Excel." },
              ...baseParts,
              ...targetParts,
              {
                text: `Your task is as follows:
                1. EXHAUSTIVE EXTRACTION: Go through ALL provided Source Data (File A). You MUST extract "First Name", "Last Name", "Training No", and "Training Name" for EVERY SINGLE PERSON mentioned or listed in those documents. Do not miss anyone, even if the list is long.
                2. NO ONE LEFT BEHIND: Even if you are 100% sure you cannot find a match in File B, you MUST still return that person in the output JSON array.
                3. MATCHING: Cross-reference every name against the Directory Data (File B) to find their email address.
                4. SUGGESTIONS (CRITICAL): If you cannot find an EXACT match but find people with SIMILAR names (e.g. "John Doe" vs "Jon Doe"), include those in the 'matchOptions' array with their name, email, and a reason/score.
                5. STATUS & EMPTY EMAILS:
                   - If a match is found: set status to 'matched' and provide the 'email'.
                   - If NO match is found: set status to 'notfound', return the 'email' as an empty string (""), and populate 'matchOptions' with 1-3 best guesses from File B.
                6. Use additional context: "${instruction}".
                
                CRITICAL: The output must be an EXHAUSTIVE list of everyone found in File A. If there are 50 names in File A, there must be 50 objects in your JSON response. DO NOT SUMMARIZE OR TRUNCATE THE LIST.
                
                REQUIRED OUTPUT (JSON Array):
                - name: Full name (First + Last)
                - email: Email address (EMPTY STRING if not found)
                - subject: Professional reminder subject line
                - body: Specific nudge body mentioning their training status
                - trainingNo: Extracted Training No (or "N/A" if missing)
                - trainingName: Name of the training session (e.g. "Ethics 2024")
                - status: "matched" | "notfound"
                - matchOptions: Array of { name: string, email: string, score: number } for potential matches if primary match is weak/missing.
                
                Strict JSON output format required.`
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
                name: { type: Type.STRING },
                email: { type: Type.STRING },
                subject: { type: Type.STRING },
                body: { type: Type.STRING },
                trainingNo: { type: Type.STRING },
                trainingName: { type: Type.STRING },
                status: { type: Type.STRING },
                matchOptions: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      name: { type: Type.STRING },
                      email: { type: Type.STRING },
                      score: { type: Type.NUMBER }
                    }
                  }
                }
              },
              required: ["name", "email", "subject", "body", "status"]
            }
          }
        }
      });

      const parsedResults: AuditResult[] = JSON.parse(response.text || "[]");
      setResults(parsedResults);
      saveCurrentSession(parsedResults);
    } catch (err: any) {
      console.error("Processing Error:", err);
      setError(err.message || "An unexpected error occurred.");
    } finally {
      setIsProcessing(false);
    }
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
      const name = r.name.padEnd(17).slice(0, 17);
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
    if (!authStatus.microsoft) {
      setError("Please connect Outlook first.");
      return;
    }
    if (!automationSettings.sourceId) {
      setError("Please provide a Google Sheet ID or SharePoint URL.");
      return;
    }

    setIsSyncing(true);
    setIsProcessing(true);
    setError(null);
    setResults([]);

    try {
      // 1. Fetch Attachment from Outlook
      const outlookRes = await fetch(`/api/outlook/fetch-attachment?subject=${encodeURIComponent(automationSettings.subjectTrigger)}`);
      if (!outlookRes.ok) throw new Error("Outlook sync failed: " + (await outlookRes.json()).error);
      const outlookData = await outlookRes.json();
      
      // 2. Fetch Directory (Google Sheets or SharePoint)
      let directoryText = "";
      if (automationSettings.sourceId.includes('sharepoint.com')) {
        const spRes = await fetch(`/api/microsoft/fetch-sharepoint?url=${encodeURIComponent(automationSettings.sourceId)}`);
        if (!spRes.ok) throw new Error("SharePoint sync failed. Ensure the file is shared and Outlook is connected.");
        const blob = await spRes.blob();
        const arrayBuffer = await blob.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer);
        workbook.SheetNames.forEach(name => {
          directoryText += `Sheet: ${name}\n${XLSX.utils.sheet_to_csv(workbook.Sheets[name])}\n`;
        });
      } else {
        if (!authStatus.google) throw new Error("Please connect Google Sheets to use a Sheet ID.");
        const sheetRes = await fetch(`/api/sheets/fetch?sheetId=${automationSettings.sourceId}`);
        if (!sheetRes.ok) throw new Error("Google Sheets sync failed: " + (await sheetRes.json()).error);
        const sheetData = await sheetRes.json();
        directoryText = `[DIRECTORY FROM GOOGLE SHEETS]:\n${JSON.stringify(sheetData.values)}`;
      }

      // 3. Prepare for Gemini
      const sourcePart = { 
        inlineData: { 
          data: outlookData.contentBytes, 
          mimeType: outlookData.contentType === 'application/pdf' ? 'application/pdf' : 'text/csv' 
        } 
      };
      const directoryPart = { text: directoryText };

      // 4. Run Audit
      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [
          {
            role: "user",
            parts: [
              { text: "You are a Deloitte Audit Assistant. I am providing you with a Source file from Outlook and a Directory from either Google Sheets or SharePoint." },
              sourcePart as any,
              directoryPart,
              {
                text: `Your task is as follows:
                1. EXHAUSTIVE EXTRACTION: Go through ALL provided Source Data. You MUST extract "First Name", "Last Name", "Training No", and "Training Name" for EVERY SINGLE PERSON mentioned.
                2. MATCHING: Cross-reference every name against the Directory to find their email address.
                3. Use context: "${instruction}".
                4. REQUIRED OUTPUT (JSON Array): name, email, subject, body, trainingNo, trainingName, status, matchOptions.`
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
                name: { type: Type.STRING },
                email: { type: Type.STRING },
                subject: { type: Type.STRING },
                body: { type: Type.STRING },
                trainingNo: { type: Type.STRING },
                trainingName: { type: Type.STRING },
                status: { type: Type.STRING },
                matchOptions: {
                  type: Type.ARRAY,
                  items: {
                    type: Type.OBJECT,
                    properties: {
                      name: { type: Type.STRING },
                      email: { type: Type.STRING },
                      score: { type: Type.NUMBER }
                    }
                  }
                }
              },
              required: ["name", "email", "subject", "body", "status"]
            }
          }
        }
      });

      const parsedResults: AuditResult[] = JSON.parse(response.text || "[]");
      setResults(parsedResults);
      saveCurrentSession(parsedResults);
    } catch (err: any) {
      console.error("Automation Error:", err);
      setError(err.message);
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
                    onClick={() => setView('audit')}
                    className={`text-[10px] font-black uppercase tracking-widest px-2 ${view === 'audit' ? 'text-deloitte' : 'text-slate-400 hover:text-white'}`}
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

        {view === 'audit' && (
          <motion.div 
            key="audit"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="flex-1 flex overflow-hidden relative"
          >
            {/* Sidebar: Configuration */}
            <motion.section 
              animate={{ width: isSidebarCollapsed ? 0 : '33.3333%', opacity: isSidebarCollapsed ? 0 : 1 }}
              className="bg-white border-r border-slate-200 flex flex-col overflow-y-auto shrink-0 transition-all duration-300"
              style={{ padding: isSidebarCollapsed ? 0 : '2rem' }}
            >
              <div className="flex justify-between items-start">
                {!isSidebarCollapsed && (
                  <div>
                    <h1 className="text-4xl font-black leading-none mb-2 tracking-tighter uppercase">Audit<br />Agent<span className="text-deloitte">.</span></h1>
                    <p className="text-slate-500 text-sm font-medium italic">Active Reconstruction Phase</p>
                  </div>
                )}
              </div>

              {!isSidebarCollapsed && (
                <div className="space-y-6 flex-1 flex flex-col">
                  {/* Automation Panel */}
                  <div className="p-4 bg-slate-900 rounded-2xl border border-slate-800 space-y-4">
                    <div className="flex justify-between items-center">
                      <h3 className="text-[10px] font-black uppercase tracking-widest text-deloitte flex items-center gap-2">
                        <div className="w-2 h-2 bg-deloitte rounded-full animate-pulse"></div>
                        Auto-Process Suite
                      </h3>
                    </div>

                    <div className="grid grid-cols-2 gap-2">
                      <button 
                        onClick={() => handleConnect('microsoft')}
                        className={`h-10 rounded-xl text-[8px] font-black uppercase tracking-widest transition-all flex items-center justify-center gap-2 ${authStatus.microsoft ? 'bg-deloitte text-black' : 'bg-slate-800 text-slate-400 hover:text-white'}`}
                      >
                        {authStatus.microsoft ? <CheckCircle2 className="w-3 h-3" /> : <Mail className="w-3 h-3" />}
                        Outlook {authStatus.microsoft ? 'Linked' : 'Connect'}
                      </button>
                      <button 
                        onClick={() => handleConnect('google')}
                        className={`h-10 rounded-xl text-[8px] font-black uppercase tracking-widest transition-all flex items-center justify-center gap-2 ${authStatus.google ? 'bg-deloitte text-black' : 'bg-slate-800 text-slate-400 hover:text-white'}`}
                      >
                        {authStatus.google ? <CheckCircle2 className="w-3 h-3" /> : <FileText className="w-3 h-3" />}
                        Sheets {authStatus.google ? 'Linked' : 'Connect'}
                      </button>
                    </div>

                    <div className="space-y-3">
                      <div>
                        <label className="text-[8px] font-black uppercase tracking-widest text-slate-500 block mb-1">Email Subject Trigger</label>
                        <input 
                          value={automationSettings.subjectTrigger}
                          onChange={(e) => setAutomationSettings(s => ({ ...s, subjectTrigger: e.target.value }))}
                          className="w-full h-10 bg-black border border-slate-800 rounded-lg px-3 text-[10px] font-mono text-deloitte focus:border-deloitte focus:outline-none"
                          placeholder="e.g. Training Report"
                        />
                      </div>
                      <div>
                        <label className="text-[8px] font-black uppercase tracking-widest text-slate-500 block mb-1">Source of Truth (Sheet ID or SharePoint URL)</label>
                        <input 
                          value={automationSettings.sourceId}
                          onChange={(e) => setAutomationSettings(s => ({ ...s, sourceId: e.target.value }))}
                          className="w-full h-10 bg-black border border-slate-800 rounded-lg px-3 text-[10px] font-mono text-deloitte focus:border-deloitte focus:outline-none"
                          placeholder="Spreadsheet ID or Deloitte SharePoint .xlsm URL"
                        />
                      </div>
                    </div>

                    <button 
                      onClick={handleSyncAndTrigger}
                      disabled={isSyncing || !authStatus.microsoft || !authStatus.google}
                      className="w-full h-12 bg-white text-black rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-deloitte transition-all disabled:opacity-20 flex items-center justify-center gap-2 shadow-xl"
                    >
                      {isSyncing ? <Loader2 className="w-4 h-4 animate-spin" /> : <Search className="w-4 h-4" />}
                      Sync & Automated Audit
                    </button>
                  </div>

                  <div className="h-px bg-slate-100"></div>

                  {/* Manual Section Label */}
                  <div className="text-[10px] font-black uppercase tracking-widest text-slate-400">Manual Fallback Ingestion</div>
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 block flex justify-between">
                      <span>Source Data (File A)</span>
                    </label>
                    
                    {baseFiles.length > 0 ? (
                      <div className="space-y-2">
                         {baseFiles.map((f, i) => (
                          <div key={i} className="p-4 bg-slate-900 rounded-xl text-white flex items-center justify-between">
                            <div className="flex items-center gap-3">
                              <FileText className="w-4 h-4 text-deloitte" />
                              <span className="text-[10px] font-black uppercase tracking-widest truncate max-w-[150px]">{f.name}</span>
                            </div>
                            <button 
                              onClick={(e) => { e.stopPropagation(); setBaseFiles(baseFiles.filter((_, idx) => idx !== i)); }}
                              className="text-slate-500 hover:text-white font-bold px-1"
                            >
                              ×
                            </button>
                          </div>
                        ))}
                        <button 
                          onClick={() => addBaseInputRef.current?.click()}
                          className="w-full p-4 border border-dashed border-slate-300 rounded-xl text-[9px] font-black uppercase tracking-widest hover:border-deloitte transition-colors flex items-center justify-center gap-2"
                        >
                          <Upload className="w-3 h-3" /> Add More Files
                        </button>
                        <input type="file" multiple ref={addBaseInputRef} className="hidden" onChange={(e) => {
                          const newFiles = Array.from(e.target.files || []) as File[];
                          const pdfs = newFiles.filter(f => f.name.toLowerCase().endsWith('.pdf'));
                          if (pdfs.length > 0) setPdfToConvert(pdfs[0]);
                          setBaseFiles(prev => [...prev, ...newFiles]);
                        }} />
                      </div>
                    ) : (
                      <div 
                        onClick={() => baseInputRef.current?.click()}
                        className="border-2 border-dashed border-slate-200 rounded-2xl p-10 flex flex-col items-center justify-center gap-4 hover:border-deloitte transition-all cursor-pointer group bg-slate-50/50 active:scale-95"
                      >
                        <div className="p-4 bg-white rounded-2xl shadow-sm group-hover:shadow-xl group-hover:-translate-y-1 transition-all">
                          <Upload className="w-8 h-8 text-deloitte" />
                        </div>
                        <div className="text-center">
                          <p className="text-[10px] font-black uppercase tracking-[0.2em]">Drop Source File</p>
                          <p className="text-[9px] font-medium text-slate-400 mt-1 uppercase tracking-widest">XLSX, CSV, PDF (Limit 10MB)</p>
                        </div>
                        <input type="file" multiple ref={baseInputRef} className="hidden" onChange={(e) => {
                          const newFiles = Array.from(e.target.files || []) as File[];
                          const pdfs = newFiles.filter(f => f.name.toLowerCase().endsWith('.pdf'));
                          if (pdfs.length > 0) setPdfToConvert(pdfs[0]);
                          setBaseFiles(prev => [...prev, ...newFiles]);
                        }} />
                      </div>
                    )}
                  </div>

                  {/* Target Files (File B) */}
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 block flex justify-between">
                      <span>Target Directory (File B)</span>
                    </label>

                    {targetFiles.length > 0 ? (
                      <div className="space-y-2">
                         {targetFiles.map((f, i) => (
                          <div key={i} className="p-4 border border-slate-200 rounded-xl text-slate-900 flex items-center justify-between">
                            <div className="flex items-center gap-3">
                              <Search className="w-4 h-4 text-deloitte" />
                              <span className="text-[10px] font-black uppercase tracking-widest truncate max-w-[150px]">{f.name}</span>
                            </div>
                            <button 
                              onClick={(e) => { e.stopPropagation(); setTargetFiles(targetFiles.filter((_, idx) => idx !== i)); }}
                              className="text-slate-400 hover:text-black font-bold px-1"
                            >
                              ×
                            </button>
                          </div>
                        ))}
                        <button 
                          onClick={() => addTargetInputRef.current?.click()}
                          className="w-full p-4 border border-dashed border-slate-300 rounded-xl text-[9px] font-black uppercase tracking-widest hover:border-deloitte transition-colors flex items-center justify-center gap-2"
                        >
                          <Upload className="w-3 h-3" /> Add Directory Files
                        </button>
                        <input type="file" multiple ref={addTargetInputRef} className="hidden" onChange={(e) => setTargetFiles(prev => [...prev, ...Array.from(e.target.files || [])])} />
                      </div>
                    ) : (
                      <div 
                        onClick={() => targetInputRef.current?.click()}
                        className="border-2 border-dashed border-slate-200 rounded-2xl p-10 flex flex-col items-center justify-center gap-4 hover:border-deloitte transition-all cursor-pointer group bg-slate-50/50 active:scale-95"
                      >
                        <div className="p-4 bg-white rounded-2xl shadow-sm group-hover:shadow-xl group-hover:-translate-y-1 transition-all">
                          <FileText className="w-8 h-8 text-slate-400 group-hover:text-deloitte" />
                        </div>
                        <div className="text-center">
                          <p className="text-[10px] font-black uppercase tracking-[0.2em]">Upload Directory</p>
                          <p className="text-[9px] font-medium text-slate-400 mt-1 uppercase tracking-widest">Internal DB, Lists, CSV</p>
                        </div>
                        <input type="file" multiple ref={targetInputRef} className="hidden" onChange={(e) => setTargetFiles(Array.from(e.target.files || []))} />
                      </div>
                    )}
                  </div>

                  {/* Scope Instruction */}
                  <div>
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 block">Processing Context</label>
                    <textarea 
                      value={instruction}
                      onChange={(e) => setInstruction(e.target.value)}
                      className="w-full h-32 bg-slate-50 border border-slate-200 rounded-2xl p-6 text-[11px] font-bold tracking-tight focus:outline-none focus:border-deloitte transition-all resize-none leading-relaxed"
                      placeholder="Provide specific extraction or matching rules..."
                    />
                  </div>

                  <div className="mt-auto pt-4 border-t border-slate-100">
                    <button 
                      onClick={handleProcess}
                      disabled={isProcessing || isConverting || baseFiles.length === 0 || targetFiles.length === 0}
                      className="w-full h-20 bg-deloitte text-black rounded-2xl font-black uppercase tracking-[0.4em] hover:bg-black hover:text-white transition-all shadow-2xl disabled:opacity-30 disabled:grayscale disabled:scale-100 flex flex-col items-center justify-center relative overflow-hidden group active:scale-95"
                    >
                      <span className="relative z-10 text-xs">{isProcessing ? 'Analyzing Patterns...' : 'Run Audit Process'}</span>
                      <span className="text-[8px] opacity-40 font-black relative z-10 tracking-widest mt-1 uppercase">
                        {isProcessing ? 'Extracting Data Points' : 'Execute Reconstruction'}
                      </span>
                      
                      {isProcessing && (
                        <motion.div 
                          className="absolute inset-0 bg-white/30 skew-x-12"
                          animate={{ x: ['-100%', '200%'] }}
                          transition={{ duration: 1.5, repeat: Infinity, ease: "easeInOut" }}
                        />
                      )}
                    </button>
                    
                    {error && (
                      <motion.div 
                        initial={{ opacity: 0, y: 10 }}
                        animate={{ opacity: 1, y: 0 }}
                        className="mt-4 p-4 bg-red-50 border border-red-100 rounded-xl flex gap-3 text-red-600"
                      >
                        <AlertCircle className="w-4 h-4 shrink-0" />
                        <span className="text-[9px] font-black uppercase tracking-widest leading-none">{error}</span>
                      </motion.div>
                    )}
                  </div>
                  <p className="text-[9px] font-bold text-slate-400 text-center uppercase tracking-widest leading-relaxed mt-4">
                    By executing, you confirm data adherence<br />to regional privacy protocols.
                  </p>
                </div>
              )}
            </motion.section>

          {/* Sidebar Collapse Toggle */}
          <button 
            onClick={() => setIsSidebarCollapsed(!isSidebarCollapsed)}
            className="absolute top-1/2 -translate-y-1/2 z-50 bg-black text-white p-2 rounded-r-xl shadow-2xl hover:bg-deloitte hover:text-black transition-all flex items-center justify-center active:scale-75"
            style={{ left: isSidebarCollapsed ? 0 : '33.3333%' }}
          >
            <ChevronRight className={`w-4 h-4 transition-transform duration-500 ${isSidebarCollapsed ? '' : 'rotate-180'}`} />
          </button>

          {/* Main Content: Results View */}
            <section className="flex-1 bg-slate-50 p-8 flex flex-col gap-6 overflow-hidden">
               <div className="flex justify-between items-end">
                <div className="flex flex-col gap-4">
                  <div>
                    <h2 className="text-6xl font-black tracking-tighter uppercase leading-[0.8] mb-2">Results<span className="text-deloitte">.</span></h2>
                    <p className="text-slate-400 font-black uppercase text-[10px] tracking-[0.3em]">
                      {results.length > 0 ? `${results.length} INDIVIDUALS IDENTIFIED IN SOURCE DATA` : 'AWAITING DATA INGESTION'}
                    </p>
                  </div>
                  
                  {results.length > 0 && (
                    <div className="flex gap-2 p-1 bg-slate-200 rounded-lg w-fit">
                      {[
                        { id: 'all', label: 'All Users' },
                        { id: 'with-email', label: 'Matched' },
                        { id: 'no-email', label: 'Missing Email' }
                      ].map((f) => (
                        <button
                          key={f.id}
                          onClick={() => setFilter(f.id as any)}
                          className={`px-4 py-2 rounded-md text-[10px] font-black uppercase tracking-widest transition-all ${
                            filter === f.id 
                            ? 'bg-black text-white shadow-lg' 
                            : 'text-slate-500 hover:bg-slate-300'
                          }`}
                        >
                          {f.label}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
                {results.length > 0 && (
                  <button 
                    onClick={handleMasterDraft}
                    className="px-10 py-5 bg-deloitte text-black rounded-lg text-xs font-black uppercase tracking-[0.2em] hover:bg-black hover:text-white transition-all shadow-2xl active:scale-95 flex items-center gap-3"
                  >
                    <Mail className="w-4 h-4" />
                    Master Draft (BCC)
                  </button>
                )}
              </div>

              <div className="flex-1 bg-white border border-slate-200 rounded-2xl overflow-hidden flex flex-col">
                <div className="grid grid-cols-12 px-8 py-4 bg-slate-100 border-b border-slate-200 text-[9px] font-black uppercase tracking-widest text-slate-500 shrink-0">
                  <div className="col-span-4">Audit Target</div>
                  <div className="col-span-5">Matched Email / Suggestions</div>
                  <div className="col-span-3 text-right">Actions</div>
                </div>

                <div className="flex-1 overflow-y-auto divide-y divide-slate-100">
                  {isProcessing ? (
                    <div className="h-full flex flex-col items-center justify-center p-20 text-center">
                      <div className="w-12 h-12 border-4 border-slate-100 border-t-deloitte rounded-full animate-spin mb-6"></div>
                      <p className="font-black text-2xl tracking-tighter uppercase">CROSS-REFERENCING DOCUMENTS</p>
                      <p className="text-slate-400 text-xs font-bold uppercase tracking-widest mt-2 animate-pulse font-mono">Multimodal Extraction in Progress...</p>
                    </div>
                  ) : results.length > 0 ? (
                    results
                      .filter(res => {
                        if (filter === 'with-email') return res.status === 'matched' || res.status === 'manual' || res.status === 'similar' || !!res.email;
                        if (filter === 'no-email') return res.status === 'notfound' && !res.email;
                        return true;
                      })
                      .map((res, idx) => (
                      <motion.div 
                        key={idx}
                        initial={{ opacity: 0, x: -10 }}
                        animate={{ opacity: 1, x: 0 }}
                        className="grid grid-cols-12 px-8 py-8 items-center hover:bg-slate-50 transition-colors group"
                      >
                        <div className="col-span-4">
                          <p className="font-black text-2xl tracking-tighter leading-none mb-1 uppercase">{res.name}</p>
                          <div className="flex gap-2 items-center">
                            <span className="px-2 py-1 bg-slate-100 text-slate-500 rounded text-[8px] font-black uppercase tracking-widest leading-none">
                              {res.trainingName || 'COURSE PENDING'}
                            </span>
                            <span className="px-2 py-1 bg-slate-900 text-white rounded text-[8px] font-black uppercase tracking-widest leading-none">
                              TRAINING ID: {res.trainingNo && res.trainingNo !== 'N/A' ? res.trainingNo : 'PENDING'}
                            </span>
                            {!res.email && (
                              <span className="px-2 py-1 bg-red-100 text-red-600 rounded text-[8px] font-black uppercase tracking-widest leading-none">
                                DISCREPANCY DETECTED
                              </span>
                            )}
                          </div>
                        </div>
                        <div className="col-span-5 pr-8">
                          {res.email ? (
                            <div className="flex flex-col">
                              <p className="font-mono text-xs text-slate-600 font-bold tracking-tight lowercase">{res.email}</p>
                              {res.status === 'manual' && (
                                <span className="text-[7px] font-black uppercase text-deloitte mt-1">Manual Override</span>
                              )}
                            </div>
                          ) : (
                            <div className="flex flex-col gap-3">
                              {res.matchOptions && res.matchOptions.length > 0 && (
                                <div className="space-y-1.5">
                                  <p className="text-[8px] font-black uppercase text-amber-600 tracking-widest flex items-center gap-1">
                                    <AlertCircle className="w-2 h-2" />
                                    Potential Matches Found:
                                  </p>
                                  <div className="flex flex-wrap gap-2">
                                    {res.matchOptions.slice(0, 3).map((opt, i) => (
                                      <button 
                                        key={i}
                                        onClick={() => updateResultEmail(idx, opt.email)}
                                        className="h-8 px-3 border border-amber-200 bg-amber-50 rounded text-[9px] font-bold text-amber-800 hover:bg-amber-500 hover:text-white hover:border-amber-500 transition-all flex items-center gap-2"
                                      >
                                        <Search className="w-3 h-3 opacity-50" />
                                        {opt.name} · {opt.email}
                                      </button>
                                    ))}
                                  </div>
                                </div>
                              )}
                              <div className="relative group/input">
                                <input 
                                  type="email" 
                                  placeholder="IDENTIFY MANUALLY..."
                                  className="w-full h-10 bg-slate-100 border border-slate-200 rounded px-4 text-[10px] font-black uppercase tracking-wider focus:bg-white focus:border-deloitte focus:outline-none transition-all placeholder:text-slate-400"
                                  onBlur={(e) => e.target.value && updateResultEmail(idx, e.target.value)}
                                  onKeyDown={(e) => e.key === 'Enter' && updateResultEmail(idx, (e.target as HTMLInputElement).value)}
                                />
                                <ExternalLink className="absolute right-3 top-1/2 -translate-y-1/2 w-3 h-3 text-slate-300 pointer-events-none" />
                              </div>
                            </div>
                          )}
                        </div>
                        <div className="col-span-3 flex justify-end gap-2">
                          <button 
                            onClick={() => handleMailto(res)}
                            disabled={!res.email}
                            className="h-12 px-8 bg-slate-900 text-white text-[10px] font-black uppercase tracking-widest rounded-lg shadow-lg hover:bg-deloitte hover:text-black transition-all flex items-center gap-2 disabled:opacity-10 group/btn"
                          >
                            <Mail className="w-4 h-4 group-hover/btn:scale-110 transition-transform" />
                            Initialize Nudge
                          </button>
                        </div>
                      </motion.div>
                    ))
                  ) : (
                    <div className="h-full flex flex-col items-center justify-center p-20 text-slate-300 text-center">
                      <Search className="w-16 h-16 mb-4 opacity-10" />
                      <p className="font-black text-2xl tracking-widest uppercase">No Active Audit Session</p>
                      <p className="text-[10px] font-black uppercase tracking-[0.3em] mt-3">Upload source files to begin the reconciliation process.</p>
                    </div>
                  )}
                </div>
              </div>
            </section>
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

        {/* PDF Conversion Prompt Modal */}
        <AnimatePresence>
          {pdfToConvert && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 z-[110] bg-black/60 backdrop-blur-sm flex items-center justify-center p-8"
            >
              <motion.div 
                initial={{ scale: 0.9 }}
                animate={{ scale: 1 }}
                className="bg-white rounded-3xl p-10 max-w-md w-full shadow-2xl text-center"
              >
                <div className="w-20 h-20 bg-deloitte/10 rounded-full flex items-center justify-center mx-auto mb-6">
                  <FileText className="w-10 h-10 text-deloitte" />
                </div>
                <h3 className="text-2xl font-black uppercase tracking-tighter mb-2">Optimize Data Source?</h3>
                <p className="text-slate-500 text-xs font-bold uppercase tracking-widest leading-relaxed mb-8">
                  You uploaded {pdfToConvert.name}. Would you like to convert this PDF into a structured Excel document for better audit accuracy?
                </p>
                <div className="flex flex-col gap-3">
                  <button 
                    onClick={() => convertPDFToExcel(pdfToConvert)}
                    className="h-14 bg-black text-white rounded-xl font-black uppercase tracking-widest hover:bg-deloitte hover:text-black transition-all flex items-center justify-center gap-3"
                  >
                    {isConverting ? <Loader2 className="w-4 h-4 animate-spin" /> : <Upload className="w-4 h-4" />}
                    Convert to Excel
                  </button>
                  <button 
                    onClick={() => setPdfToConvert(null)}
                    className="h-14 border border-slate-200 rounded-xl font-black uppercase tracking-widest hover:bg-slate-50 transition-all text-slate-400"
                  >
                    Use Original PDF
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
