import React, { useState, useRef, useEffect } from 'react';
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
  ArrowRight,
  Lock,
  Unlock,
  Settings
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
  const [directoryFile, setDirectoryFile] = useState<File | null>(null);
  
  // Use user-specific keys for persistence
  const getStorageKey = (suffix: string) => `deloitte_directory_${currentUser?.email || 'guest'}_${suffix}`;

  const [directoryMetadata, setDirectoryMetadata] = useState<{ name: string, updatedAt: string, size: string } | null>(() => {
    const email = currentUser?.email;
    if (!email) return null;
    const saved = localStorage.getItem(`deloitte_directory_${email}_meta`);
    if (saved) return JSON.parse(saved);
    
    // Default initialization if no directory exists
    const defaultMeta = {
      name: 'Master_Directory_V1.csv',
      updatedAt: new Date().toLocaleString(),
      size: '4.2 KB'
    };
    return defaultMeta;
  });

  const [isDirectoryLocked, setIsDirectoryLocked] = useState(true);

  // Initialize default content if missing
  useEffect(() => {
    if (!currentUser) return;

    // Sync metadata if null (happens after login)
    if (!directoryMetadata) {
      const saved = localStorage.getItem(`deloitte_directory_${currentUser.email}_meta`);
      if (saved) {
        setDirectoryMetadata(JSON.parse(saved));
      } else {
        setDirectoryMetadata({
          name: 'Master_Directory_V1.csv',
          updatedAt: new Date().toLocaleString(),
          size: '4.2 KB'
        });
      }
    }

    const contentKey = getStorageKey('content');
    const metaKey = getStorageKey('meta');
    if (!localStorage.getItem(contentKey)) {
      const defaultCSV = `name,email
"Devine, Sean",sedevine@deloitte.com
"Saha , Barun",barsaha@deloitte.com
"Thube, Gauri",gthube@deloitte.com
"Mathew, Ansa",ansmathew@deloitte.com
"Vamsi, Ramabathina",ravamsi@deloitte.com
"Goenka, Mayank",magoenka@deloitte.com
"S, Ramkumar",ramlnu@deloitte.com
"M, Sruthi",sruthm@deloitte.com
"Arora, Titiksha",tiarora@deloitte.com
"BHASKAR, BATTALA",batbhaskar@deloitte.com
"Viswa Teja, Adabala",aviswateja@deloitte.com
"Smith, Ash",aybala@deloitte.com
"Bandekar, Abhishek",abbandekar@deloitte.com
"Bhargavi, Vayaluru",vaybhargavi@deloitte.com
"Jha, Suman",sumajha@deloitte.com
"Rai, Ekta",ekrai@deloitte.com
"G, Sneha",sneg@deloitte.com
"Soni, Shweta",shwsoni@deloitte.ie
"Kumar, Varun",varunkumar6@deloitte.com
"Ezzat, Ola",oezzat@deloitte.com
"Halim, Nada",nhalim@deloitte.com
"Venkateswarlu , M",vmuthyala.ext@deloitte.com
"Aggarwal, Priyanshi",priyansh@deloitte.com
Kumar Raghavendra Venkata Raghav Aravapalli,vrkaravapalli@deloitte.com
"Y, Pradhiksha",pradhlnu@deloitte.com
"Mac Coitir, Diarmuid",dmaccoitir@deloitte.ie
"Maan, Harsh",hmaan@deloitte.com
"Deshmukh, Shivathmika",shivatdeshmukh@deloitte.com
"Guntreddy, Renuka",gurenuka@deloitte.com
"Karmakar, Shruti",shrkarmakar@deloitte.com
"Sachin Dahake, Vaishnav",vsachindahake@deloitte.com
"Arora, Nikhil",nikhilarora@deloitte.com
"Griffin, Niall",niagriffin@deloitte.ie
"Jain, Dev",devjain6@deloitte.com
"Kassiotis, Jamie",jkassiotis@deloitte.com
"Desai, Arti",ardesai@deloitte.com
"Popli, Labhesh",lpopli@deloitte.com
"S, Karthiga",karthigas@deloitte.com
"Jaiswal, Monika",mojaiswal@deloitte.com
"Sajja, Tanuja",tsajja@deloitte.com
"Verma , Shubham",shubhaverma@deloitte.com
"Jawa, Bhumika",bhjawa@deloitte.com
"Anand Kashikar, Vidya",vanandkashikar@deloitte.com
"Malviya, Aadesh",aamalviya@deloitte.com
"Hernandez, Alex",AlexanHernandez@deloitte.com
"Poivey-Olson, Harrison",hpoivey-olson@deloitte.com
"C M, Mahadevan",mahcm@deloitte.com
"Mansfield, Chris",chmansfield@deloitte.ie
"Rathore, Rajni",rajnirathore@deloitte.com
"Magdy, Mina",mimagdy@deloitte.com
"Haddad, Jack",jahaddad@deloitte.com
"Amir, Elana",eamir@deloitte.com
"McCarthy, Donnchadh",domccarthy@deloitte.ie
"Mohapatra, Soumyajit",soumyajmohapatra@deloitte.com
"Behera, Tanmay",tanbehera@deloitte.com
"Babu, Shubhasmita",shubabu@deloitte.com
"Plamada, Stefan-Sorin",splamada@deloitte.com
"M, Russel",rmondayapurath@deloitte.ie
"Kokate, Onkar",osubhashkokate@deloitte.com
"Roy, Suraj",surroy@deloitte.com
"Divakar T D, Suprith",sdivakartd@deloitte.com
"Timbadia, Dhvani",dtimbadia@deloitte.com
"Kodali, Anuhya",kanuhya@deloitte.com
"Rehman U, Afsal",afsau@deloitte.com
"Kumar, Tadipathri",tadipkumar@deloitte.com
"Aman, Alwani,",aalwani@deloitte.com
"Behera, Rudrakshya",rudbehera@deloitte.com
"Daya, Dhiliban,",dadp@deloitte.com
"Chen, Melissa",melischen@deloitte.com
"O'Rourke, Cian",cianorourke@deloitte.ie
"Prakasam, Medha",meprakasam@deloitte.ie
"babu, Vijay, Vijay",vvijaybabu@deloitte.com
"Tiwari, Shivangi",shivapathak@deloitte.com
"Prabhakar, Tanvi",taprabhakar@deloitte.com
"Banerjee, Abishek",abanerjee5@deloitte.com
"Ray Doocey, John",jdoocey@deloitte.ie
"Panchal, Shyam",shypanchal@deloitte.ie
"Singh, Rashmi",rassingh@deloitte.ie
"Gupta, Renu",rengupta@deloitte.com
"Pasam, Mounika",mpasam@deloitte.com
"Venkatesh, Chittimelli",chvenkatesh@deloitte.com
"Srivastava, Rishabh",rishsrivastava@deloitte.com
"Sudhakar Tawar, Sushil",sutawar@deloitte.com
"reddy bobbili, Aditi",areddybobbili@deloitte.com
"Maskar, Abishek",amaskar@deloitte.com
"Shaikh, Anhar",anshaikh@deloitte.com
"A, Prince",prina@deloitte.com
"Saklani, Upasana",usaklani@deloitte.com
"Raju, Kalyan",kalraju@deloitte.com
"Burkholder, Nate",nburkholder@deloitte.com
"Lingenfelter, Isaac",ilingenfelter@deloitte.com
"Stoica, Ionut-Tudorel",iostoic@deloitte.com
"Beniwal, Sahil",sbeniwal@deloitte.com
"Gaurav, Limbore,",glimbore@deloitte.com
"Suthar, Dheeraj",dhsuthar@deloitte.com
"Bandu Godham, Omkar",ogodham@deloitte.com
"LNU, Ravi",rlnu18@deloitte.com
"Pal, Niharika",nihapal@deloitte.com
"Hirphode, Niranjan",nhirphode@deloitte.com
"Vijaya Lakshmi, Gandham, Sowmya",gsowmyavijayalaks@deloitte.com
"Sayesh Reddy, Pallapolu",psayeshreddy@deloitte.com
"Kumar, Sumit",sumitkumar67@deloitte.com
"Barik, Satyajit",satbarik@deloitte.com
"Kumar Biswal, Santosh",santobiswal@deloitte.com
"Salvarajan, Praveen",prselvarajan@deloitte.com
"Devi Natarajan, Nithya",nithnatarajan@deloitte.com
"Sarkar, Sauvik",sauvsarkar@deloitte.com
"K, Akshatha",akshathk@deloitte.com
"Ranjan, Lalan",lranjan@deloitte.com
"Dey, Prithviraj",pritdey@deloitte.com
"Sabyasachi, Dey,",sabdey@deloitte.com
"Sahebrao Chormale, Suraj",ssahebraochormale@deloitte.com
"Acharyya, Sagnik",saacharyya@deloitte.com
"Kishor Mali, Samadhan",skishormali@deloitte.com
"Rishu, Kumar",kurishu@deloitte.com
"Kedarisetti, Vishnu",vkedarisetti@deloitte.com
"Pachauri, Ankit",ankpachauri@deloitte.com
"Santhoshkumar, S",ssanthoshkumar2@deloitte.com
"Bhapkar, Shivraj",sbhapkar@deloitte.com
"Kumar G, Dinesh",dkumarg@deloitte.com
"Bhargava, Kartik",kartbhargava@deloitte.com
"Iqbal Shaikh, Anjum",anjshaikh@deloitte.com
"Pullareddygari, Pravallika",ppullareddygari@deloitte.com
"Jain, Madhur",madhurjain@deloitte.com
"Kumar Gupta, Ankit",ankumargupta@deloitte.com
"Panjwani, Akber",akpanjwani@deloitte.com
"Yogeshvaran, S,",yogeshvs@deloitte.com
"Anusha, T,",anushat@deloitte.com
"Trivedi, Awadhesh",awtrivedi@deloitte.ie
"Ravindra Reddy, Sirasani",sravindrareddy@deloitte.com
"Srikanth, Nettem",nettsrikanth@deloitte.com
"Kumar Annupoojari Anigalale, Akshay",aannupoojarianigalal@deloitte.com
"Matt, Pan,",matpan@deloitte.com
"Srivastava, Ayushman",ayushmsrivastava@deloitte.com
"Kulkarni, Prathamesh",prathameskulkarni@deloitte.com
"Bagwan, Mafaruk",mbagwan@deloitte.com
"Mukherjee, Subarna",subamukherjee@deloitte.ie
"Dutta, Aniket",anikdutta@deloitte.com
"Acharya, Pratyush",pratyacharya@deloitte.com
"Vidyananda Sagar, Boorle",bvidyanandasagar@deloitte.com
"Yadav, Ranjeet",ranjeyadav@deloitte.com
"P, Sargar",psagar9@deloitte.com
"Trivedi, Chitransh",ctrivedi@deloitte.com
"Davis, Kai",kdavis2@deloitte.com
"Ramachandran, Sreekanth",sreramachandran@deloitte.com
"Patel, Sunil",sunilpatel@deloitte.com
"Sonawane, Saili",ssonawane@deloitte.com
"Markandeya, Akshata",anerkar@deloitte.com
"Shah, Rishil",rishilshah@deloitte.com
"Naveena Spoorthi, Tankasala",ntankasala@deloitte.com
"Travers, Patrick",ptravers@deloitte.ie
"T, Sivakumar",sivt@deloitte.com
"Tandon, Ojasvini",otandon@deloitte.ie
"Madhukar, Vaibhav",vmadhukar@deloitte.com
"Raj Ayyappa, Nagender",nayyappa@deloitte.com
"Law, Dana",danalaw@deloitte.ie
"Reddy Jakkam, Maheswar",mjakkam@deloitte.com
"Gupta, Neha",negupta.ext@deloitte.com
"Palanisamy, Thilakavathy",thpalanisamy@deloitte.com
"Kumar, Shloka",shlkumar@deloitte.com
"N Kumar, Varsha",varsnkumar@deloitte.com
"Chang, Alice",alchang@deloitte.com
"Lyons, Eimear",eilyons@deloitte.ie
"Taher, Mohamad",mohataher@deloitte.com
"Markle, Rich",rmarkle@deloitte.com
"Abraham, Sonu",ssonuabraham@deloitte.com
"O'Callaghan Smith, Orla",oocallaghansmith@deloitte.ie
"Montoya, Daniel",damontoya@deloitte.com
"Lynch, Deiric",deilynch@deloitte.ie
"Hession, Etain",ehession@deloitte.ie
"Barrie, Owen",obarrie@deloitte.ie
"Vanarajan, Mullai",mvanarajan@deloitte.com
"Rasamsetty, Anusha,",rasanusha@deloitte.com
"Henry, Randy",rahenry@deloitte.com
"Kumar , Dhiraj",dhirkumar@deloitte.com
"Mendoza, Adrian",admendoza@deloitte.com
"Akshaya, M,",akshayam@deloitte.com
"Yadav, Pooja",pyadav29@deloitte.com
"Hema, P,",hemap@deloitte.com
"Kherha, Jasveer",jkherha@deloitte.com
"Paramjyothi, Aketi",aparamjyothi@deloitte.com
"Pawar, Akanksha",aamarjeetpawar@deloitte.com
"Dand, Jinesha",jdand@deloitte.com
"Kumar, Dhiraj",dhirkumar@deloitte.com
"Sahu, Alok",alosahu@deloitte.com
"B Bhandari, Lavanya",labhandari@deloitte.com
"V G, Ramya",ramvg@deloitte.com
"Fox, Craig",crfox@deloitte.com
"Siddiq, Kamran",ksiddiq@deloitte.com
"Merriman, Luke",lumerriman@deloitte.ie
"Asfour, Ramez",raasfour@deloitte.com
"Moroney, Brian",bmoroney@deloitte.ie
"Kumar, Ashish",askumar10@deloitte.com
"Marnane, Aidan",amarnane@deloitte.ie
"Shree Venugopalan, Vidya",vidvenugopalan@deloitte.com
"Kumar, Yogesh",ykumar5@deloitte.com
"Eisenhut, Tanner",teisenhut@deloitte.com
"Fahad, Mohammad",mfahad@deloitte.com
"El-Hanafy, Nadin",nelhanafy@deloitte.com
"Zoheb Jahagirdar, Mohammed",mjahagirdar@deloitte.com
"Thomas, Shijo",shithomas@deloitte.com
"Saini, Jatin",jatisaini@deloitte.com
"Abdul Rab Ansari, Fareen",fabdulrabansari@deloitte.com
"Vajula, Rajesh",rvajula@deloitte.com
"Carroll, Riona",caitcarroll@deloitte.ie
"Mourya, Nirbhay",nmourya@deloitte.com
"Kumar, Sachin",skumar48@deloitte.ie
"Surabhi, Sweta",ssurabhi@deloitte.com
"Gopal, Priyanka",priygopal@deloitte.com
"Ganguly, Varsha",vaganguly@deloitte.com
"Bhaskaruni, Maheswari",mbhaskaruni@deloitte.com
"Tomar, Sonali",sotomar@deloitte.com
"Sharma, Manisha",manishasharma4@deloitte.com
"Dash, Sushant",susdash@deloitte.com
"Bhat, Aravinda",arbhat@deloitte.com
"Gorai, Sandeep",sgorai@deloitte.com
"Singh, Govind",gosingh@deloitte.com
"Ghanukota, Sahaja",gsahaja@deloitte.com
"O'Malley, Allan",alomalley@deloitte.ie
"Hazra , Sourav",souhazra@deloitte.com
"Suneelkumar, Parvatala",ksuneelkumar@deloitte.com
"Gupta, Naman",ngupta65@deloitte.com
"Wintz, Finn",fiwintz@deloitte.com
"Pattanayak, Amrit",ampattanayak@deloitte.com
"Shravya, Shravya",glshravya@deloitte.com
"Das Mahapatra, Debanjan",ddasmahapatra@deloitte.com
"Vasavi, Vasavi",kavasavi@deloitte.com
"Lnu, Karthikeyan",karthikeyan73@deloitte.com
"Butt, Nomaan",nombutt@deloitte.com
"Cabrera, Stephanie",stcabrera@deloitte.com
"Ferry, Zachary",zferry@deloitte.com
"Shaikh Salaam, Jamesha",jshaikhsalaam@deloitte.com
"G Nomula, Makarand",mgnomula@deloitte.com
"Pani, Puratan",pupani@deloitte.com
"Kota, Susmitha",susmkota@deloitte.com
"Chateker, Sheetal",schateker@deloitte.com
"Mahajan, Gaurav",gaurmahajan@deloitte.com
"Bhatt, Pavitra",pabhatt.ext@deloitte.com
"Ali Bohra, Hussain",hubohra@deloitte.com
"Varige, Vidyasri,",vvidyasri@deloitte.com
"Rohani, Ashkan",arohani@deloitte.com
"Arroyo, Federico",farroyo@deloitte.com
"Avani, Khoti,",avkhoti@deloitte.com
"Prasannakumari, Matangi,",pramatangi@deloitte.com
"Basak, Deeptesh",deebasak@deloitte.com
"TS, Balaji",bats@deloitte.com
"Rath, Debashish",derath@deloitte.com
"Embiricos, Saya",sayuno@deloitte.com`;
      localStorage.setItem(contentKey, defaultCSV);
      localStorage.setItem(metaKey, JSON.stringify({
        name: 'Master_Directory_V1.csv',
        updatedAt: new Date().toLocaleString(),
        size: '4.2 KB'
      }));
    }
  }, [currentUser?.email]);

  const saveDirectoryMetadata = async (file: File) => {
    if (!currentUser) return;
    const meta = {
      name: file.name,
      updatedAt: new Date().toLocaleString(),
      size: (file.size / 1024).toFixed(1) + ' KB'
    };
    setDirectoryMetadata(meta);
    localStorage.setItem(`deloitte_directory_${currentUser?.email}_meta`, JSON.stringify(meta));
    
    // Persist content user-specifically
    const text = await extractTextFromFile(file);
    localStorage.setItem(`deloitte_directory_${currentUser?.email}_content`, text);
  };
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

  // Lazy initialization for Gemini AI
  const getAI = () => {
    const key = process.env.GEMINI_API_KEY;
    if (!key || key === 'MY_GEMINI_API_KEY' || key === '') {
      throw new Error("Gemini API key is missing. Please configure it in the 'Secrets' panel in AI Studio.");
    }
    return new GoogleGenAI({ apiKey: key });
  };

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
      const data = await res.json();
      
      if (!res.ok) {
        throw new Error(data.error || `Server returned ${res.status}`);
      }

      if (!data.url) {
        throw new Error("No authorization URL received from server");
      }

      window.open(data.url, `${type}_oauth`, 'width=600,height=700');
    } catch (e: any) {
      console.error(`Oauth initialization error (${type}):`, e);
      setError(`Failed to initiate ${type} connection: ${e.message}`);
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
      if (!currentUser) return prev;
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
      // 1. Fetch Directory Persistent Storage OR Session Override
      let directoryText = "";
      const contentKey = `deloitte_directory_${currentUser?.email}_content`;
      
      if (isDirectoryLocked && directoryMetadata) {
        const savedContent = localStorage.getItem(contentKey);
        if (savedContent) {
          directoryText = `[DIRECTORY REPOSITORY FROM PERSISTENT CACHE (${directoryMetadata.name})]:\n${savedContent.slice(0, 50000)}`;
        } else {
          throw new Error("Base directory content missing. Please re-upload on the dashboard.");
        }
      } else if (directoryFile) {
        const text = await extractTextFromFile(directoryFile);
        directoryText = `[DIRECTORY REPOSITORY FROM UPLOADED SESSION OVERRIDE ${directoryFile.name}]:\n${text.slice(0, 50000)}`;
      } else {
        throw new Error("No Directory available. Unlock to upload a session override or use the dashboard to set a base Excel.");
      }

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
      const ai = getAI();
      const response = await ai.models.generateContent({
        model: "gemini-2.0-flash",
        contents: [
          {
            role: "user",
            parts: [
              { text: "You are a Deloitte Audit Assistant. I am providing you with one or more Source files (uploaded by the user) and a Directory Repository for reconciliation." },
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

      const resultText = (response as any).text || (response as any).response?.text?.() || "";
      const parsedResults: AuditResult[] = JSON.parse(resultText || "[]");
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
              <div className="space-y-4">
                <h1 className="text-7xl font-black tracking-tighter leading-none uppercase">Project<br />Directory<span className="text-deloitte">.</span></h1>
                <div className="flex items-center gap-6">
                  <p className="text-slate-400 font-black uppercase text-sm tracking-[0.3em]">Welcome, Agent {currentUser.name}</p>
                  <div className="h-4 w-px bg-slate-300"></div>
                  <div className="flex items-center gap-2">
                    <div className="w-2 h-2 bg-deloitte rounded-full"></div>
                    <span className="text-[10px] font-black uppercase tracking-widest text-slate-500">System Ready</span>
                  </div>
                </div>
              </div>
              
              <div className="flex gap-4">
                 <button 
                  onClick={() => setIsProcessing(!isProcessing)}
                  className="px-6 h-16 border border-slate-200 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-slate-50 transition-all flex items-center gap-2"
                >
                  <Settings className="w-4 h-4" />
                  Quick Tools
                </button>
              </div>
            </div>

            <div className="grid grid-cols-12 gap-8">
              {/* Directory Management Card */}
              <div className="col-span-12 lg:col-span-5 space-y-6">
                <div className="bg-black text-white p-10 rounded-3xl space-y-8 shadow-2xl relative overflow-hidden group">
                  <div className="absolute top-0 right-0 w-64 h-64 bg-deloitte/10 blur-[100px] -translate-y-1/2 translate-x-1/2"></div>
                  
                  <div className="relative z-10">
                    <h3 className="text-xs font-black uppercase tracking-[0.3em] text-deloitte mb-6 flex items-center gap-3">
                      <FileText className="w-4 h-4" />
                      Directory Repository
                    </h3>
                    
                    {directoryMetadata ? (
                      <div className="space-y-6">
                        <div>
                          <p className="text-4xl font-black tracking-tighter uppercase leading-tight mb-2 truncate group-hover:text-deloitte transition-colors">{directoryMetadata.name}</p>
                          <div className="flex items-center gap-3">
                            <span className="px-2 py-1 bg-white/10 text-[9px] font-black uppercase tracking-widest rounded">
                              {directoryMetadata.size}
                            </span>
                            <span className="text-slate-400 text-[10px] font-bold uppercase tracking-widest">
                              Updated {directoryMetadata.updatedAt}
                            </span>
                          </div>
                        </div>
                        
                        <div className="p-4 bg-white/5 border border-white/10 rounded-2xl">
                          <p className="text-[10px] font-bold text-slate-300 uppercase leading-relaxed tracking-wider">
                            <span className="text-deloitte">IMPORTANT:</span> This is your persistent base Excel used for all audit reconciliations. If you have a more updated master list, replace it here.
                          </p>
                        </div>
                      </div>
                    ) : (
                      <div className="space-y-4">
                        <p className="text-slate-400 text-xs font-bold uppercase tracking-widest leading-relaxed">
                          No base directory found. Upload the master Employee List to enable reconciliation.
                        </p>
                      </div>
                    )}
                  </div>

                  <div className="space-y-4 relative z-10">
                    <button 
                      onClick={() => {
                        const input = document.createElement('input');
                        input.type = 'file';
                        input.accept = '.xlsx, .xls, .csv';
                        input.onchange = (e) => {
                          const file = (e.target as HTMLInputElement).files?.[0];
                          if (file) {
                            saveDirectoryMetadata(file);
                          }
                        };
                        input.click();
                      }}
                      className="w-full h-14 bg-white text-black rounded-xl font-black uppercase tracking-widest hover:bg-deloitte transition-all flex items-center justify-center gap-3 shadow-xl"
                    >
                      <Upload className="w-4 h-4" />
                      Update Base Repository
                    </button>

                    <button 
                      onClick={startNewSession}
                      className="w-full h-16 bg-deloitte text-black rounded-xl text-xs font-black uppercase tracking-[0.2em] hover:bg-black hover:text-white transition-all shadow-2xl active:scale-95 flex items-center justify-center gap-3"
                    >
                      <Search className="w-5 h-5" />
                      Initialize New Audit
                    </button>
                  </div>
                </div>

                <div className="p-8 border border-slate-200 rounded-3xl bg-white shadow-sm">
                  <div className="flex justify-between items-center mb-4">
                    <label className="text-[10px] font-black uppercase tracking-widest text-slate-400">CC Persistence Memory</label>
                    <Mail className="w-4 h-4 text-slate-300" />
                  </div>
                  <input 
                    value={ccMemory}
                    onChange={(e) => setCcMemory(e.target.value)}
                    placeholder="manager@deloitte.com; hr@deloitte.com"
                    className="w-full h-14 bg-slate-50 border border-slate-100 rounded-xl px-6 text-xs font-bold uppercase tracking-widest focus:border-deloitte focus:outline-none transition-all"
                  />
                  <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mt-4 leading-relaxed">
                    Set a persistent list of BCC recipients that will be automatically added to all outgoing nudges.
                  </p>
                </div>
              </div>

              {/* Session History List */}
              <div className="col-span-7 bg-white border border-slate-200 rounded-3xl overflow-hidden shadow-sm flex flex-col">
                <div className="p-8 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                  <h3 className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Historical Audit Records</h3>
                  <div className="flex gap-2">
                    <span className="px-3 py-1 bg-white border border-slate-200 rounded-full text-[9px] font-black uppercase tracking-widest text-slate-500 shadow-sm">
                      {sessions.filter(s => s.userEmail === currentUser?.email).length} Records
                    </span>
                  </div>
                </div>
                
                <div className="flex-1 overflow-y-auto divide-y divide-slate-100 min-h-[400px]">
                  {sessions.filter(s => s.userEmail === currentUser?.email).length === 0 ? (
                    <div className="h-full flex flex-col items-center justify-center p-20 text-center text-slate-300">
                      <FileText className="w-12 h-12 mx-auto mb-4 opacity-10" />
                      <p className="font-black text-lg tracking-widest uppercase">Vault Empty</p>
                      <p className="text-[10px] font-black uppercase tracking-widest mt-2">Initialize an audit to create a record.</p>
                    </div>
                  ) : (
                    sessions
                      .filter(s => s.userEmail === currentUser?.email)
                      .sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime())
                      .map(s => (
                      <div 
                        key={s.id}
                        onClick={() => {
                          setResults(s.metadata.results || []);
                          setView('audit-step-3');
                        }}
                        className="p-8 hover:bg-slate-50 transition-all cursor-pointer group flex justify-between items-center"
                      >
                        <div className="space-y-1">
                          <div className="flex items-center gap-3">
                            <span className="text-xl font-black uppercase tracking-tighter group-hover:text-deloitte transition-colors">Audit #{s.id.slice(-4)}</span>
                            <span className="px-2 py-0.5 bg-black text-white text-[8px] font-black uppercase tracking-widest rounded-full">
                              Completed
                            </span>
                          </div>
                          <div className="flex gap-3 text-[10px] items-center">
                            <span className="text-slate-400 font-bold uppercase tracking-widest">{new Date(s.timestamp).toLocaleDateString()}</span>
                            <span className="w-1 h-1 bg-slate-200 rounded-full"></span>
                            <span className="text-slate-500 font-black uppercase tracking-widest">
                              {s.resultCount || 0} Targets Audited
                            </span>
                          </div>
                        </div>
                        <div className="w-12 h-12 rounded-xl bg-slate-50 border border-slate-100 flex items-center justify-center group-hover:bg-black group-hover:text-white transition-all">
                          <ChevronRight className="w-5 h-5" />
                        </div>
                      </div>
                    ))
                  )}
                </div>
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

                <div className="p-8 bg-black rounded-3xl border border-slate-800 space-y-6 relative overflow-hidden">
                  <div className="flex justify-between items-center">
                    <h3 className="text-xs font-black uppercase tracking-widest text-deloitte flex items-center gap-2">
                      <div className="w-2 h-2 bg-deloitte rounded-full"></div>
                      Directory Repository
                    </h3>
                    {directoryMetadata && (
                      <button 
                        onClick={() => setIsDirectoryLocked(!isDirectoryLocked)}
                        className={`px-3 py-1 rounded-full text-[8px] font-black uppercase tracking-widest transition-all flex items-center gap-2 ${isDirectoryLocked ? 'bg-slate-800 text-slate-400' : 'bg-deloitte text-black'}`}
                      >
                        {isDirectoryLocked ? <Lock className="w-2 h-2" /> : <Unlock className="w-2 h-2" />}
                        {isDirectoryLocked ? 'Locked' : 'Unlocked'}
                      </button>
                    )}
                  </div>

                  <AnimatePresence mode="wait">
                    {isDirectoryLocked && directoryMetadata ? (
                      <motion.div 
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        exit={{ opacity: 0 }}
                        key="locked"
                        className="space-y-4"
                      >
                        <div>
                          <p className="text-white text-2xl font-black uppercase tracking-tighter truncate">{directoryMetadata.name}</p>
                          <p className="text-slate-500 text-[8px] font-bold uppercase tracking-[0.2em] mt-1">LATEST UPDATED SOURCE · {directoryMetadata.updatedAt}</p>
                        </div>
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest leading-relaxed">
                          This base directory is currently active for this session. Unlock to replace with a temporary session repository.
                        </p>
                      </motion.div>
                    ) : (
                      <motion.div 
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        exit={{ opacity: 0 }}
                        key="unlocked"
                        className="space-y-4"
                      >
                         {directoryFile ? (
                            <div className="p-4 bg-slate-900 border border-deloitte/30 rounded-xl flex items-center justify-between">
                              <div className="flex items-center gap-3 overflow-hidden">
                                <FileText className="w-5 h-5 text-deloitte shrink-0" />
                                <div>
                                  <p className="text-[10px] font-black uppercase text-white truncate">{directoryFile.name}</p>
                                  <p className="text-[8px] font-bold text-deloitte uppercase tracking-widest">Session Override Ready</p>
                                </div>
                              </div>
                              <button 
                                onClick={() => setDirectoryFile(null)}
                                className="text-slate-500 hover:text-white"
                              >
                                ×
                              </button>
                            </div>
                          ) : (
                            <div className="space-y-4">
                              <div>
                                <label className="text-[8px] font-black uppercase tracking-widest text-slate-500 block mb-2">Upload Session Override (Excel/CSV)</label>
                                <button 
                                  onClick={() => {
                                    const input = document.createElement('input');
                                    input.type = 'file';
                                    input.accept = '.xlsx, .xls, .csv';
                                    input.onchange = (e) => {
                                      const file = (e.target as HTMLInputElement).files?.[0];
                                      if (file) setDirectoryFile(file);
                                    };
                                    input.click();
                                  }}
                                  className="w-full h-12 border border-dashed border-slate-700 rounded-xl text-[10px] font-black uppercase tracking-widest text-slate-400 hover:border-deloitte hover:text-white transition-all flex items-center justify-center gap-2"
                                >
                                  <Upload className="w-4 h-4" />
                                  Choose Session File
                                </button>
                                <p className="text-[8px] font-bold text-slate-500 uppercase tracking-widest mt-2">Note: This will only be used for the current audit session.</p>
                              </div>
                            </div>
                          )}
                      </motion.div>
                    )}
                  </AnimatePresence>
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
                      disabled={isSyncing || uploadedFiles.length === 0 || (!directoryFile && !automationSettings.sourceId)}
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
