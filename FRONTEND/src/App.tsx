import React, { useState, useEffect, Component } from "react";
import { 
  LayoutDashboard, 
  BookOpen, 
  Users, 
  DoorOpen, 
  Calendar, 
  Grid, 
  FileText, 
  Plus, 
  Upload, 
  Download, 
  Trash2,
  Search,
  AlertCircle,
  LogIn,
  LogOut,
  Zap,
  Printer,
  Edit2,
  Menu,
  X,
  ChevronLeft,
  ChevronRight
} from "lucide-react";
import { 
  collection, 
  addDoc, 
  deleteDoc, 
  doc, 
  onSnapshot, 
  setDoc,
  getDocFromServer,
  writeBatch
} from "firebase/firestore";
import { signInWithPopup, GoogleAuthProvider, onAuthStateChanged, signOut, User } from "firebase/auth";
import { db, auth } from "./firebase";
import { 
  Programme, 
  Subject, 
  Student, 
  Room, 
  TimetableEntry, 
  SeatingPlan
} from "./types";
import { generateSeatingArrangement } from "./services/seatingService";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { motion, AnimatePresence } from "motion/react";
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer
} from "recharts";

// --- Types for Error Handling ---
enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | null | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | null | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}

function handleFirestoreError(error: any, operationType: OperationType, path: string | null): string {
  const code = error?.code;
  let friendlyMessage = "";
  
  switch (code) {
    case 'permission-denied':
      friendlyMessage = "Access Denied: You don't have permission to perform this action. Ensure you are logged in as an authorized admin.";
      break;
    case 'unauthenticated':
      friendlyMessage = "Authentication Required: Please sign in to continue.";
      break;
    case 'unavailable':
      friendlyMessage = "Service Unavailable: The database is currently unreachable. Please check your internet connection.";
      break;
    case 'not-found':
      friendlyMessage = "Not Found: The requested record could not be located.";
      break;
    case 'already-exists':
      friendlyMessage = "Conflict: This record already exists in the system.";
      break;
    case 'resource-exhausted':
      friendlyMessage = "Quota Exceeded: The system has reached its usage limit for today. Please try again tomorrow.";
      break;
    case 'deadline-exceeded':
      friendlyMessage = "Timeout: The operation took too long to complete. Please try again.";
      break;
    case 'cancelled':
      friendlyMessage = "Operation Cancelled: The request was aborted.";
      break;
    default:
      friendlyMessage = error?.message || "An unexpected database error occurred. Please try again.";
  }

  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData.map(provider => ({
        providerId: provider.providerId,
        displayName: provider.displayName,
        email: provider.email,
        photoUrl: provider.photoURL
      })) || []
    },
    operationType,
    path
  }
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  return friendlyMessage;
}

interface ErrorBoundaryProps {
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: any;
}

class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  state: ErrorBoundaryState;
  props: ErrorBoundaryProps;
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: any) {
    return { hasError: true, error };
  }

  componentDidCatch(error: any, errorInfo: any) {
    console.error("ErrorBoundary caught an error", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      let errorMessage = "Something went wrong.";
      try {
        const parsed = JSON.parse(this.state.error.message);
        if (parsed.error) errorMessage = parsed.error;
      } catch (e) {
        errorMessage = this.state.error.message || errorMessage;
      }

      return (
        <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
          <div className="bg-white p-8 rounded-2xl shadow-xl border border-slate-200 max-w-md w-full text-center">
            <div className="w-16 h-16 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center mx-auto mb-6">
              <AlertCircle size={32} />
            </div>
            <h2 className="text-2xl font-bold text-slate-800 mb-2">Application Error</h2>
            <p className="text-slate-600 mb-6">{errorMessage}</p>
            <button 
              onClick={() => window.location.reload()}
              className="w-full bg-indigo-600 text-white py-3 rounded-xl font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-200"
            >
              Reload Application
            </button>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

// --- Components ---

const SidebarItem = ({ icon: Icon, label, active, onClick, collapsed }: { icon: any, label: string, active: boolean, onClick: () => void, collapsed?: boolean }) => (
  <button
    onClick={onClick}
    className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all relative group ${
      active 
        ? "bg-indigo-600 text-white shadow-lg shadow-indigo-100" 
        : "text-slate-600 hover:bg-indigo-50 hover:text-indigo-600"
    } ${collapsed ? "justify-center px-2" : ""}`}
    title={collapsed ? label : ""}
  >
    <Icon size={20} className="shrink-0" />
    {!collapsed && <span className="font-medium truncate">{label}</span>}
    {collapsed && !active && (
      <div className="absolute left-full ml-2 px-2 py-1 bg-slate-800 text-white text-xs rounded opacity-0 group-hover:opacity-100 pointer-events-none transition-opacity whitespace-nowrap z-50">
        {label}
      </div>
    )}
  </button>
);

const Card = ({ children, title, action, noPadding }: { children: React.ReactNode, title?: string, action?: React.ReactNode, noPadding?: boolean }) => (
  <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
    {(title || action) && (
      <div className="px-4 md:px-6 py-4 border-b border-slate-100 flex flex-col xl:flex-row xl:items-center justify-between gap-4">
        {title && <h3 className="font-semibold text-slate-800 shrink-0">{title}</h3>}
        {action && <div className="flex shrink-0 max-w-full overflow-x-auto no-scrollbar">{action}</div>}
      </div>
    )}
    <div className={noPadding ? "" : "p-4 md:p-6"}>{children}</div>
  </div>
);

const Modal = ({ isOpen, onClose, title, children }: { isOpen: boolean, onClose: () => void, title: string, children: React.ReactNode }) => (
  <AnimatePresence>
    {isOpen && (
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          onClick={onClose}
          className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm"
        />
        <motion.div
          initial={{ scale: 0.95, opacity: 0 }}
          animate={{ scale: 1, opacity: 1 }}
          exit={{ scale: 0.95, opacity: 0 }}
          className="relative bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden flex flex-col max-h-[90vh] md:max-h-[80vh]"
        >
          <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between shrink-0">
            <h3 className="font-bold text-lg text-slate-800">{title}</h3>
            <button onClick={onClose} className="text-slate-400 hover:text-slate-600 p-2">
              <Plus className="rotate-45" size={24} />
            </button>
          </div>
          <div className="p-6 overflow-y-auto">{children}</div>
        </motion.div>
      </div>
    )}
  </AnimatePresence>
);

// --- Main App ---

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [isAuthReady, setIsAuthReady] = useState(false);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [programmes, setProgrammes] = useState<Programme[]>([]);
  const [subjects, setSubjects] = useState<Subject[]>([]);
  const [students, setStudents] = useState<Student[]>([]);
  const [rooms, setRooms] = useState<Room[]>([]);
  const [timetable, setTimetable] = useState<TimetableEntry[]>([]);
  const [seatingPlans, setSeatingPlans] = useState<SeatingPlan[]>([]);

  // Modal States (Moved to top level to avoid hook violations)
  const [isStudentUploadModalOpen, setIsStudentUploadModalOpen] = useState(false);
  const [isStudentClearModalOpen, setIsStudentClearModalOpen] = useState(false);
  const [isSubjectUploadModalOpen, setIsSubjectUploadModalOpen] = useState(false);
  const [isRoomAddModalOpen, setIsRoomAddModalOpen] = useState(false);
  const [isTimetableAddModalOpen, setIsTimetableAddModalOpen] = useState(false);
  const [isTimetableAutoModalOpen, setIsTimetableAutoModalOpen] = useState(false);
  const [isTimetableClearModalOpen, setIsTimetableClearModalOpen] = useState(false);
  const [isTimetableEditModalOpen, setIsTimetableEditModalOpen] = useState(false);
  const [isSubjectEditModalOpen, setIsSubjectEditModalOpen] = useState(false);
  const [isStudentEditModalOpen, setIsStudentEditModalOpen] = useState(false);
  const [isDeleteSemesterModalOpen, setIsDeleteSemesterModalOpen] = useState(false);
  const [isDeleteConfirmModalOpen, setIsDeleteConfirmModalOpen] = useState(false);
  const [isSeatingClearModalOpen, setIsSeatingClearModalOpen] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState<{ collection: string, id: string, label: string } | null>(null);
  const [deleteSemesterTarget, setDeleteSemesterTarget] = useState<{ progId: string, semester: number, progName: string } | null>(null);
  const [editingTimetableEntry, setEditingTimetableEntry] = useState<TimetableEntry | null>(null);
  const [editingStudent, setEditingStudent] = useState<Student | null>(null);
  const [editingSubject, setEditingSubject] = useState<Subject | null>(null);
  const [editSubjectData, setEditSubjectData] = useState<Omit<Subject, "id">>({
    programmeId: "",
    semester: 1,
    code: "",
    name: "",
    type: "Theory",
    specialization: "General"
  });
  const [editStudentData, setEditStudentData] = useState<Omit<Student, "id">>({
    studentId: "",
    name: "",
    programmeId: "",
    specialization: "",
    semester: 1
  });
  const [statusMessage, setStatusMessage] = useState<{ type: 'success' | 'error' | 'info', text: string } | null>(null);

  // New states for manual inputs
  const [newProgName, setNewProgName] = useState("");
  const [newSubjProg, setNewSubjProg] = useState("");
  const [newSubjSem, setNewSubjSem] = useState("");
  const [newSubjCode, setNewSubjCode] = useState("");
  const [newSubjName, setNewSubjName] = useState("");
  const [newSubjType, setNewSubjType] = useState<"Theory" | "Practical">("Theory");
  const [newSubjSpec, setNewSubjSpec] = useState("General");
  const [subjectSearchTerm, setSubjectSearchTerm] = useState("");
  const [subjectSpecFilter, setSubjectSpecFilter] = useState<string | "All">("All");
  const [subjectProgFilter, setSubjectProgFilter] = useState<string | "All">("All");
  const [subjectSemFilter, setSubjectSemFilter] = useState<number | "All">("All");
  const [selectedSemesterFilter, setSelectedSemesterFilter] = useState<number | "All">("All");

  // Form States
  const [studentSearchTerm, setStudentSearchTerm] = useState("");
  const [selectedStudentProg, setSelectedStudentProg] = useState<string | "All">("All");
  const [selectedStudentSpec, setSelectedStudentSpec] = useState<string | "All">("All");
  const [newRoom, setNewRoom] = useState({ roomNumber: "", capacity: 0, columns: 0, seatsPerColumn: [] as number[] });
  const [seatsPerColumnInput, setSeatsPerColumnInput] = useState("");
  const [autoGenConfig, setAutoGenConfig] = useState({
    startDate: "",
    preferredEndDate: "",
    shiftType: "mixed" as "single" | "double" | "mixed",
    mixedPattern: "2,1,2,2,1",
    slots: ["09:00 AM - 12:00 PM", "02:00 PM - 05:00 PM"],
    excludeSundays: true,
    publicHolidays: "", // Comma separated dates YYYY-MM-DD
    semester: 0,
    totalExamsToSchedule: 0
  });
  const [calculatedMinEndDate, setCalculatedMinEndDate] = useState("");

  useEffect(() => {
    if (!autoGenConfig.startDate) {
      setCalculatedMinEndDate("");
      return;
    }

    const totalCapacity = (rooms || []).reduce((acc, r) => acc + (r.capacity || 0), 0);
    if (totalCapacity === 0) {
      setCalculatedMinEndDate(autoGenConfig.startDate);
      return;
    }

    // Get subjects to schedule
    const scheduledSubjectIds = new Set(timetable.map(t => t.subjectId));
    let subjectsToSchedule = subjects.filter(s => !scheduledSubjectIds.has(s.id));

    if (autoGenConfig.semester !== 0) {
      subjectsToSchedule = subjectsToSchedule.filter(s => s.semester === autoGenConfig.semester);
    }

    let count = autoGenConfig.totalExamsToSchedule > 0 
      ? Math.min(autoGenConfig.totalExamsToSchedule, subjectsToSchedule.length)
      : subjectsToSchedule.length;

    if (count === 0) {
      setCalculatedMinEndDate(autoGenConfig.startDate);
      return;
    }

    // Pre-calculate student counts
    const subjectsWithCounts = subjectsToSchedule.slice(0, count === 0 ? undefined : count).map(s => {
      const studentCount = students.filter(st => 
        st.programmeId === s.programmeId && 
        st.semester === s.semester &&
        (s.specialization === "General" || !s.specialization || st.specialization === s.specialization)
      ).length;
      return { ...s, studentCount };
    });

    // Sort descending for better packing (First Fit Decreasing)
    const sortedSubjects = [...subjectsWithCounts].sort((a, b) => b.studentCount - a.studentCount);

    // Check if any single subject exceeds capacity
    if (sortedSubjects.some(s => s.studentCount > totalCapacity)) {
      setCalculatedMinEndDate("Capacity Error");
      return;
    }

    const getShiftsForDay = (dayIndex: number) => {
      let count = 0;
      if (autoGenConfig.shiftType === "single") count = 1;
      else if (autoGenConfig.shiftType === "double") count = 2;
      else {
        const pattern = autoGenConfig.mixedPattern.split(",").map(n => parseInt(n.trim()) || 0);
        count = pattern[dayIndex % pattern.length] || 0;
      }
      return Math.min(count, autoGenConfig.slots.length);
    };

    // Simulate packing to find min shifts needed
    // We need to track: which subjects are in which shift, and which programmes are on which day
    let shiftsNeeded = 0;
    let subjectsRemaining = [...sortedSubjects];
    let dayIndex = 0;
    let dayOffset = 0;
    let minEndDate = new Date(autoGenConfig.startDate);

    while (subjectsRemaining.length > 0) {
      const current = new Date(autoGenConfig.startDate);
      current.setDate(current.getDate() + dayOffset);
      const dayOfWeek = current.getDay();
      const dateStr = current.toISOString().split('T')[0];
      const holidays = autoGenConfig.publicHolidays.split(",").map(h => h.trim());
      
      const isHoliday = holidays.includes(dateStr);
      const isSunday = dayOfWeek === 0;
      
      if (!isSunday && !isHoliday) {
        const shiftsInDay = getShiftsForDay(dayIndex);
        const dailyProgrammes = new Set<string>();

        for (let s = 0; s < shiftsInDay; s++) {
          if (subjectsRemaining.length === 0) break;
          
          let shiftCapacity = totalCapacity;
          const shiftSubjectsIndices: number[] = [];

          for (let i = 0; i < subjectsRemaining.length; i++) {
            const sub = subjectsRemaining[i];
            const progKey = `${sub.programmeId}-${sub.semester}`;
            
            if (sub.studentCount <= shiftCapacity && !dailyProgrammes.has(progKey)) {
              shiftCapacity -= sub.studentCount;
              dailyProgrammes.add(progKey);
              shiftSubjectsIndices.push(i);
            }
          }

          // Remove scheduled subjects
          subjectsRemaining = subjectsRemaining.filter((_, idx) => !shiftSubjectsIndices.includes(idx));
          if (shiftSubjectsIndices.length > 0) {
            minEndDate = new Date(current);
          }
        }
        dayIndex++;
      }
      dayOffset++;
      if (dayOffset > 2000) break; // Safety
    }

    const minDateStr = minEndDate.toISOString().split('T')[0];
    setCalculatedMinEndDate(minDateStr);
    
    if (!autoGenConfig.preferredEndDate || autoGenConfig.preferredEndDate < minDateStr) {
      setAutoGenConfig(prev => ({ ...prev, preferredEndDate: minDateStr }));
    }
  }, [autoGenConfig.startDate, autoGenConfig.shiftType, autoGenConfig.mixedPattern, autoGenConfig.excludeSundays, autoGenConfig.publicHolidays, autoGenConfig.slots, autoGenConfig.semester, autoGenConfig.totalExamsToSchedule, subjects, timetable, students, rooms]);

  const studentStats: {
    overallTotal: number;
    byProgram: Record<string, {
      total: number;
      bySemester: Record<number, {
        total: number;
        bySpecialization: Record<string, number>;
      }>;
    }>;
  } = React.useMemo(() => {
    const stats = {
      overallTotal: students.length,
      byProgram: {} as Record<string, {
        total: number,
        bySemester: Record<number, {
          total: number,
          bySpecialization: Record<string, number>
        }>
      }>
    };

    students.forEach(s => {
      const progId = s.programmeId;
      const sem = s.semester;
      const spec = s.specialization || "General";

      if (!stats.byProgram[progId]) {
        stats.byProgram[progId] = { total: 0, bySemester: {} };
      }
      stats.byProgram[progId].total++;

      if (!stats.byProgram[progId].bySemester[sem]) {
        stats.byProgram[progId].bySemester[sem] = { total: 0, bySpecialization: {} };
      }
      stats.byProgram[progId].bySemester[sem].total++;

      if (!stats.byProgram[progId].bySemester[sem].bySpecialization[spec]) {
        stats.byProgram[progId].bySemester[sem].bySpecialization[spec] = 0;
      }
      stats.byProgram[progId].bySemester[sem].bySpecialization[spec]++;
    });

    return stats;
  }, [students, programmes]);

  const [newTimetableEntry, setNewTimetableEntry] = useState<Omit<TimetableEntry, "id">>({
    date: "",
    timeSlot: "",
    programmeId: "",
    semester: 1,
    subjectId: ""
  });
  const [selectedSeatingDate, setSelectedSeatingDate] = useState(() => localStorage.getItem("selectedSeatingDate") || "");
  const [selectedSeatingSlot, setSelectedSeatingSlot] = useState(() => localStorage.getItem("selectedSeatingSlot") || "");

  useEffect(() => {
    localStorage.setItem("selectedSeatingDate", selectedSeatingDate);
  }, [selectedSeatingDate]);

  useEffect(() => {
    localStorage.setItem("selectedSeatingSlot", selectedSeatingSlot);
  }, [selectedSeatingSlot]);

  const [loading, setLoading] = useState(true);
  const [connectionError, setConnectionError] = useState<string | null>(null);

  // --- Connection Test ---
  useEffect(() => {
    async function testConnection() {
      try {
        // Test connection to Firestore
        await getDocFromServer(doc(db, 'test', 'connection'));
        console.log("Firestore connection verified.");
        setConnectionError(null);
      } catch (error) {
        if (error instanceof Error && (error.message.includes('the client is offline') || error.message.includes('unavailable'))) {
          console.error("Please check your Firebase configuration. Firestore is unreachable.");
          setConnectionError("Could not reach Cloud Firestore backend. Please check your Firebase configuration or internet connection.");
        }
        // Skip logging for other errors (like permission denied on the test doc), 
        // as this is simply a connection test.
      }
    }
    testConnection();
  }, []);

  // --- Auth Handling ---
  useEffect(() => {
    // Verify API connectivity
    fetch("/api/health")
      .then(res => res.json())
      .then(data => console.log("API connectivity verified:", data))
      .catch(err => console.error("API connectivity error:", err));

    const unsub = onAuthStateChanged(auth, (u) => {
      setUser(u);
      setIsAuthReady(true);
      setLoading(false);
    });
    return () => unsub();
  }, []);

  const [isGenerating, setIsGenerating] = useState(false);
  const [isDataLoading, setIsDataLoading] = useState(true);

  // --- Data Fetching ---
  useEffect(() => {
    if (!isAuthReady || !user) return;
    
    setIsDataLoading(true);
    let loadedCount = 0;
    const totalCollections = 6;

    const checkLoaded = () => {
      loadedCount++;
      if (loadedCount >= totalCollections) {
        setIsDataLoading(false);
      }
    };

    const unsubProgrammes = onSnapshot(collection(db, "programmes"), 
      (s) => {
        setProgrammes(s.docs.map(d => ({ id: d.id, ...d.data() } as Programme)));
        checkLoaded();
      },
      (err) => setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.LIST, "programmes") })
    );
    const unsubSubjects = onSnapshot(collection(db, "subjects"), 
      (s) => {
        setSubjects(s.docs.map(d => ({ id: d.id, ...d.data() } as Subject)));
        checkLoaded();
      },
      (err) => setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.LIST, "subjects") })
    );
    const unsubStudents = onSnapshot(collection(db, "students"), 
      (s) => {
        setStudents(s.docs.map(d => ({ id: d.id, ...d.data() } as Student)));
        checkLoaded();
      },
      (err) => setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.LIST, "students") })
    );
    const unsubRooms = onSnapshot(collection(db, "rooms"), 
      (s) => {
        setRooms(s.docs.map(d => ({ id: d.id, ...d.data() } as Room)));
        checkLoaded();
      },
      (err) => setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.LIST, "rooms") })
    );
    const unsubTimetable = onSnapshot(collection(db, "timetable"), 
      (s) => {
        setTimetable(s.docs.map(d => ({ id: d.id, ...d.data() } as TimetableEntry)));
        checkLoaded();
      },
      (err) => setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.LIST, "timetable") })
    );
    const unsubSeating = onSnapshot(collection(db, "seating_plans"), 
      (s) => {
        setSeatingPlans(s.docs.map(d => ({ id: d.id, ...d.data() } as SeatingPlan)));
        checkLoaded();
      },
      (err) => setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.LIST, "seating_plans") })
    );

    return () => {
      unsubProgrammes(); unsubSubjects(); unsubStudents(); unsubRooms(); unsubTimetable(); unsubSeating();
    };
  }, [isAuthReady, user]);

  const [isInitializingRooms, setIsInitializingRooms] = useState(false);

  const PREDEFINED_ROOMS = [
    { roomNumber: "PG-FC1", capacity: 54, columns: 6, seatsPerColumn: [9, 9, 9, 9, 9, 9] },
    { roomNumber: "PG-FC2", capacity: 46, columns: 5, seatsPerColumn: [10, 10, 10, 8, 8] },
    { roomNumber: "PG-SC1", capacity: 28, columns: 4, seatsPerColumn: [7, 7, 7, 7] },
    { roomNumber: "PG-SC2", capacity: 54, columns: 6, seatsPerColumn: [9, 9, 9, 9, 9, 9] },
    { roomNumber: "PG-SC3", capacity: 46, columns: 5, seatsPerColumn: [10, 10, 10, 8, 8] },
    { roomNumber: "CS-SC1", capacity: 59, columns: 6, seatsPerColumn: [10, 10, 10, 10, 10, 9] },
    { roomNumber: "CS-SC2", capacity: 49, columns: 5, seatsPerColumn: [10, 10, 10, 10, 9] },
  ];

  const initializePredefinedRooms = async () => {
    if (isInitializingRooms) return;
    setIsInitializingRooms(true);
    try {
      let addedCount = 0;
      // Use a local set to track rooms being added in this session
      // to avoid race conditions before state updates
      const currentRoomNumbers = new Set(rooms.map(r => r.roomNumber.toLowerCase().trim()));
      
      for (const room of PREDEFINED_ROOMS) {
        const roomKey = room.roomNumber.toLowerCase().trim();
        if (!currentRoomNumbers.has(roomKey)) {
          await handleAddRoom(room);
          currentRoomNumbers.add(roomKey);
          addedCount++;
        }
      }
      if (addedCount > 0) {
        setStatusMessage({ type: 'success', text: `${addedCount} predefined rooms initialized successfully!` });
      } else {
        setStatusMessage({ type: 'info', text: "All predefined rooms already exist." });
      }
    } catch (err) {
      console.error("Error initializing rooms:", err);
      setStatusMessage({ type: 'error', text: "Failed to initialize predefined rooms." });
    } finally {
      setIsInitializingRooms(false);
    }
  };

  // Remove automatic room initialization as per user request
  /*
  useEffect(() => {
    if (isAuthReady && user && rooms.length === 0 && !isInitializingRooms) {
      const timer = setTimeout(() => {
        if (rooms.length === 0) {
          initializePredefinedRooms();
        }
      }, 1500);
      return () => clearTimeout(timer);
    }
  }, [isAuthReady, user, rooms.length]);
  */

  // Clear status message after 5 seconds
  useEffect(() => {
    if (statusMessage) {
      const timer = setTimeout(() => {
        setStatusMessage(null);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [statusMessage]);

  // Clear status message when switching tabs to avoid persistent warnings on every page
  useEffect(() => {
    setStatusMessage(null);
  }, [activeTab]);

  const handleLogin = async () => {
    const provider = new GoogleAuthProvider();
    try {
      const result = await signInWithPopup(auth, provider);
      const user = result.user;
      // Create user document if it doesn't exist
      await setDoc(doc(db, "users", user.uid), {
        name: user.displayName,
        email: user.email,
        role: user.email === "p96540114@gmail.com" ? "admin" : "user",
        lastLogin: new Date().toISOString()
      }, { merge: true });
    } catch (error) {
      console.error("Login error:", error);
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
    } catch (error) {
      console.error("Logout error:", error);
    }
  };

  const handleAddProgramme = async (name: string) => {
    if (!name.trim()) {
      setStatusMessage({ type: 'error', text: "Programme name cannot be empty." });
      return;
    }
    try {
      await addDoc(collection(db, "programmes"), { name: name.trim() });
      setStatusMessage({ type: 'success', text: `Programme "${name}" added successfully.` });
      setNewProgName("");
    } catch (err) {
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.CREATE, "programmes") });
    }
  };

  const handleAddSubject = async (data: Omit<Subject, "id">) => {
    try {
      await addDoc(collection(db, "subjects"), data);
      setStatusMessage({ type: 'success', text: `Subject "${data.name}" added successfully.` });
      setNewSubjSem("");
      setNewSubjCode("");
      setNewSubjName("");
      setNewSubjSpec("General");
    } catch (err) {
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.CREATE, "subjects") });
    }
  };

  const handleAddRoom = async (data: Omit<Room, "id">) => {
    try {
      // Use roomNumber as the document ID to prevent duplicates at the database level
      const roomId = data.roomNumber.trim().toUpperCase();
      if (!roomId) throw new Error("Room number is required");
      
      await setDoc(doc(db, "rooms", roomId), {
        ...data,
        roomNumber: data.roomNumber.trim() // Keep original casing for display but trimmed
      });
    } catch (err) {
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.CREATE, "rooms") });
    }
  };

  const handleAutoGenerateTimetable = async () => {
    const today = new Date().toISOString().split('T')[0];
    
    if (!autoGenConfig.startDate) {
      setStatusMessage({ type: 'error', text: "Please provide a start date." });
      return;
    }

    if (autoGenConfig.startDate < today) {
      setStatusMessage({ type: 'error', text: "Start date cannot be in the past." });
      return;
    }

    if (calculatedMinEndDate === "Capacity Error") {
      setStatusMessage({ type: 'error', text: "One or more exams exceed the total seating capacity. Please add more rooms or check student counts." });
      return;
    }

    if (autoGenConfig.preferredEndDate && autoGenConfig.preferredEndDate < calculatedMinEndDate) {
      setStatusMessage({ type: 'error', text: `End date cannot be earlier than the minimum required date (${calculatedMinEndDate}).` });
      return;
    }

    const totalCapacity = (rooms || []).reduce((acc, r) => acc + (r.capacity || 0), 0);
    if (totalCapacity === 0) {
      setStatusMessage({ type: 'error', text: "No rooms available. Please add rooms with capacity first." });
      return;
    }

    // Get subjects to schedule
    const scheduledSubjectIds = new Set(timetable.map(t => t.subjectId));
    let subjectsToSchedule = subjects.filter(s => !scheduledSubjectIds.has(s.id));

    if (autoGenConfig.semester !== 0) {
      subjectsToSchedule = subjectsToSchedule.filter(s => s.semester === autoGenConfig.semester);
    }

    if (autoGenConfig.totalExamsToSchedule > 0) {
      subjectsToSchedule = subjectsToSchedule.slice(0, autoGenConfig.totalExamsToSchedule);
    }

    if (subjectsToSchedule.length === 0) {
      setStatusMessage({ type: 'info', text: "No subjects to schedule." });
      return;
    }

    // Pre-calculate student counts
    const subjectsWithCounts = subjectsToSchedule.map(s => {
      const count = students.filter(st => 
        st.programmeId === s.programmeId && 
        st.semester === s.semester &&
        (s.specialization === "General" || !s.specialization || st.specialization === s.specialization)
      ).length;
      return { ...s, studentCount: count };
    });

    // Sort by Programme and Semester for program-wise organization
    const sortedSubjects = [...subjectsWithCounts].sort((a, b) => {
      const progA = programmes.find(p => p.id === a.programmeId)?.name || "";
      const progB = programmes.find(p => p.id === b.programmeId)?.name || "";
      if (progA !== progB) return progA.localeCompare(progB);
      return a.semester - b.semester;
    });

    // Helper to get shifts for a day index
    const getShiftsForDay = (dayIndex: number) => {
      let count = 0;
      if (autoGenConfig.shiftType === "single") count = 1;
      else if (autoGenConfig.shiftType === "double") count = 2;
      else {
        const pattern = autoGenConfig.mixedPattern.split(",").map(n => parseInt(n.trim()) || 0);
        count = pattern[dayIndex % pattern.length] || 0;
      }
      return Math.min(count, autoGenConfig.slots.length);
    };

    // 1. Identify all available shifts until finalEndDate
    const finalEndDate = new Date(autoGenConfig.preferredEndDate || calculatedMinEndDate);
    const availableShifts: { date: string, slot: string, dayIndex: number }[] = [];
    let d = new Date(autoGenConfig.startDate);
    let dayIndex = 0;
    const holidays = autoGenConfig.publicHolidays.split(",").map(h => h.trim());

    while (d <= finalEndDate) {
      const dayOfWeek = d.getDay();
      const dateStr = d.toISOString().split('T')[0];
      const isHoliday = holidays.includes(dateStr);
      const isSunday = dayOfWeek === 0;

      if (!isSunday && !isHoliday) {
        const shiftsInDay = getShiftsForDay(dayIndex);
        for (let i = 0; i < shiftsInDay; i++) {
          if (i < autoGenConfig.slots.length) {
            availableShifts.push({ 
              date: dateStr, 
              slot: autoGenConfig.slots[i],
              dayIndex
            });
          }
        }
        dayIndex++;
      }
      d.setDate(d.getDate() + 1);
    }

    // 2. Pack subjects into virtual shifts
    // This is the "Minimum Shifts Needed" calculation
    const packedShifts: { subjects: typeof sortedSubjects }[] = [];
    let subjectsRemaining = [...sortedSubjects];
    let virtualDayIndex = 0;
    let virtualDayOffset = 0;

    while (subjectsRemaining.length > 0) {
      const shiftsInDay = getShiftsForDay(virtualDayIndex);
      const dailyProgrammes = new Set<string>();

      for (let s = 0; s < shiftsInDay; s++) {
        if (subjectsRemaining.length === 0) break;
        
        let shiftCapacity = totalCapacity;
        const shiftSubjects: typeof sortedSubjects = [];
        const shiftIndices: number[] = [];

        for (let i = 0; i < subjectsRemaining.length; i++) {
          const sub = subjectsRemaining[i];
          const progKey = `${sub.programmeId}-${sub.semester}`;
          
          if (sub.studentCount <= shiftCapacity && !dailyProgrammes.has(progKey)) {
            shiftCapacity -= sub.studentCount;
            dailyProgrammes.add(progKey);
            shiftSubjects.push(sub);
            shiftIndices.push(i);
          }
        }

        if (shiftSubjects.length > 0) {
          packedShifts.push({ subjects: shiftSubjects });
          subjectsRemaining = subjectsRemaining.filter((_, idx) => !shiftIndices.includes(idx));
        }
      }
      virtualDayIndex++;
      virtualDayOffset++;
      if (virtualDayOffset > 2000) break;
    }

    const numPackedShifts = packedShifts.length;
    const numAvailableShifts = availableShifts.length;

    if (numAvailableShifts < numPackedShifts) {
      setStatusMessage({ 
        type: 'error', 
        text: `Not enough capacity to schedule all exams. Needed: ${numPackedShifts} shifts, Available: ${numAvailableShifts}.` 
      });
      return;
    }

    // 3. Distribute packed shifts across available shifts
    const newEntries: Omit<TimetableEntry, "id">[] = [];
    // Use sequential distribution to respect the shifting rule (Odd/Even shifts)
    // Spreading with step > 1 often skips the second shift entirely in double-shift setups
    const step = 1; 

    for (let i = 0; i < numPackedShifts; i++) {
      const targetShiftIdx = i;
      const targetShift = availableShifts[targetShiftIdx];
      const packed = packedShifts[i];

      packed.subjects.forEach(subject => {
        newEntries.push({
          date: targetShift.date,
          timeSlot: targetShift.slot,
          programmeId: subject.programmeId,
          semester: subject.semester,
          subjectId: subject.id
        });
      });
    }

    try {
      for (const entry of newEntries) {
        await addDoc(collection(db, "timetable"), entry);
      }
      setStatusMessage({ 
        type: 'success', 
        text: `Successfully generated ${newEntries.length} exams across ${numPackedShifts} shifts.` 
      });
      setIsTimetableAutoModalOpen(false);
    } catch (err) {
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.CREATE, "timetable") });
    }
  };

  // API Connectivity Check
  useEffect(() => {
    const checkApi = async () => {
      try {
        console.log("[App] Checking API connectivity at /api/ping...");
        const res = await fetch("/api/ping");
        const contentType = res.headers.get("content-type");
        console.log(`[App] API Response status: ${res.status}, content-type: ${contentType}`);
        
        if (contentType && contentType.includes("application/json")) {
          const data = await res.json();
          console.log("[App] API Connectivity Check Success:", data);
        } else {
          const text = await res.text();
          console.error("[App] API Connectivity Check returned non-JSON:", text.substring(0, 200));
          console.warn("[App] This usually means the request hit the SPA fallback. Check server routing.");
        }
      } catch (err) {
        console.error("[App] API Connectivity Check Failed:", err);
      }
    };
    checkApi();
  }, []);

  const handleExcelUpload = async (file: File, type: "students" | "timetable" | "subjects") => {
    console.log(`Starting ${type} upload for file:`, file.name);
    const formData = new FormData();
    formData.append("file", file);
    let endpoint = "";
    if (type === "students") endpoint = "/api/upload-students";
    else if (type === "timetable") endpoint = "/api/upload-timetable";
    else endpoint = "/api/upload-subjects";
    
    try {
      console.log(`[App] Uploading ${type} to ${endpoint}`);
      const res = await fetch(endpoint, { method: "POST", body: formData });
      
      const contentType = res.headers.get("content-type");
      console.log(`[App] Response status: ${res.status}, content-type: ${contentType}`);
      if (!res.ok) {
        let errorMessage = `Upload failed with status ${res.status}`;
        try {
          if (contentType && contentType.includes("application/json")) {
            const errorData = await res.json();
            errorMessage = errorData.error || errorMessage;
          } else {
            const text = await res.text();
            console.error("Non-JSON error response:", text);
          }
        } catch (e) {
          console.error("Error parsing error response:", e);
        }
        throw new Error(errorMessage);
      }

      if (!contentType || !contentType.includes("application/json")) {
        const text = await res.text();
        console.error("Expected JSON but received:", text.substring(0, 100) + "...");
        throw new Error(`Server returned an invalid response format (expected JSON, got ${contentType || 'unknown'}). This usually means the API route was not found or the server encountered an error.`);
      }

      const result = await res.json();
      let successCount = 0;
      let skipCount = 0;
      
      if (type === "students") {
        const localProgrammes = [...programmes];
        for (const s of result.students) {
          try {
            let prog = localProgrammes.find(p => p.name.toLowerCase() === s.programme.toLowerCase());
            
            // If programme doesn't exist, create it
            if (!prog) {
              const newProgRef = await addDoc(collection(db, "programmes"), {
                name: s.programme,
                description: `Automatically created during student upload`
              });
              prog = { id: newProgRef.id, name: s.programme, description: "" };
              localProgrammes.push(prog);
            }

            const semester = isNaN(Number(s.semester)) ? 0 : Number(s.semester);
            
            // Use studentId as the document ID and setDoc with merge: true
            // This prevents "Document already exists" errors and allows updating existing students
            await setDoc(doc(db, "students", s.studentId), {
              studentId: s.studentId,
              name: s.name,
              programmeId: prog.id,
              specialization: s.specialization || "General",
              semester: semester
            }, { merge: true });
            successCount++;
          } catch (err) {
            console.error("Error adding student:", err);
            console.error("Student data:", s);
            skipCount++;
          }
        }
        setStatusMessage({ type: 'success', text: `Successfully uploaded ${successCount} students. ${skipCount > 0 ? `Skipped ${skipCount} due to errors.` : ""}` });
      } else if (type === "timetable") {
        const today = new Date().toISOString().split('T')[0];
        for (const t of result.timetable) {
          if (t.date < today) {
            skipCount++;
            continue;
          }
          const prog = programmes.find(p => p.name.toLowerCase() === t.programme.toLowerCase());
          const subj = subjects.find(s => s.code.toLowerCase() === t.subjectCode.toLowerCase());
          if (prog && subj) {
            await addDoc(collection(db, "timetable"), {
              date: t.date,
              timeSlot: t.timeSlot,
              programmeId: prog.id,
              semester: t.semester,
              subjectId: subj.id
            });
            successCount++;
          } else {
            skipCount++;
          }
        }
        setStatusMessage({ type: 'success', text: `Successfully uploaded ${successCount} timetable entries. ${skipCount > 0 ? `Skipped ${skipCount} due to missing programmes/subjects.` : ""}` });
      } else if (type === "subjects") {
        const localProgrammes = [...programmes];
        for (const s of result.subjects) {
          let prog = localProgrammes.find(p => p.name.toLowerCase() === s.programme.toLowerCase());
          
          if (!prog) {
            const newProgRef = await addDoc(collection(db, "programmes"), {
              name: s.programme,
              description: `Automatically created during subject upload`
            });
            prog = { id: newProgRef.id, name: s.programme, description: "" };
            localProgrammes.push(prog);
          }

          await addDoc(collection(db, "subjects"), {
            programmeId: prog.id,
            semester: s.semester,
            code: s.subjectCode,
            name: s.subjectName,
            type: s.subjectType || "Theory"
          });
          successCount++;
        }
        setStatusMessage({ type: 'success', text: `Successfully uploaded ${successCount} subjects.` });
      }
      
      setIsStudentUploadModalOpen(false);
      setIsSubjectUploadModalOpen(false);
    } catch (err) {
      console.error(err);
      setStatusMessage({ 
        type: 'error', 
        text: err instanceof Error ? err.message : "An unexpected error occurred during upload." 
      });
    }
  };

  const handleGenerateSeating = async (date: string, timeSlot: string) => {
    if (!date || !timeSlot) return;
    
    setIsGenerating(true);
    setStatusMessage({ type: 'info', text: "Generating seating arrangement..." });
    
    try {
      const plans = generateSeatingArrangement(students, rooms, timetable, subjects, date, timeSlot);
      
      if (plans.length === 0) {
        setStatusMessage({ type: 'error', text: "No eligible students found for this shift." });
        setIsGenerating(false);
        return;
      }

      // Use batches for better performance (max 500 operations per batch)
      const batchSize = 500;
      for (let i = 0; i < plans.length; i += batchSize) {
        const batch = writeBatch(db);
        const chunk = plans.slice(i, i + batchSize);
        
        chunk.forEach(plan => {
          batch.set(doc(db, "seating_plans", plan.id), plan);
        });
        
        await batch.commit();
      }
      
      setStatusMessage({ type: 'success', text: `Successfully generated seating for ${plans.length} students!` });
    } catch (err) {
      console.error("Generation error:", err);
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.WRITE, "seating_plans") });
    } finally {
      setIsGenerating(false);
    }
  };

  const exportToExcel = (data: any[], fileName: string) => {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `${fileName}.xlsx`);
  };

  const exportSeatingToExcel = (plans: SeatingPlan[], fileName: string) => {
    const wb = XLSX.utils.book_new();
    
    // Group by room
    const roomsMap: Record<string, SeatingPlan[]> = {};
    plans.forEach(p => {
      const room = rooms.find(r => r.id === p.roomId)?.roomNumber || "Unknown";
      if (!roomsMap[room]) roomsMap[room] = [];
      roomsMap[room].push(p);
    });

    // Create a sheet for each room
    Object.keys(roomsMap).sort().forEach(roomName => {
      const roomPlans = roomsMap[roomName].sort((a, b) => {
        if (a.seat !== b.seat) return a.seat - b.seat;
        return a.row - b.row;
      });

      const data = roomPlans.map(p => {
        const student = students.find(s => s.id === p.studentId);
        const prog = programmes.find(pr => pr.id === student?.programmeId);
        return {
          "Room Name": roomName,
          "Row Number": p.row,
          "Seat Number": p.seat,
          "Student Name": student?.name || "N/A",
          "Roll Number": student?.studentId || "N/A",
          "Program": prog?.name || "N/A",
          "Semester": student?.semester || "N/A"
        };
      });

      const ws = XLSX.utils.json_to_sheet(data);
      
      // Set column widths for better readability
      const wscols = [
        {wch: 15}, // Room Name
        {wch: 12}, // Row Number
        {wch: 12}, // Seat Number
        {wch: 25}, // Student Name
        {wch: 15}, // Roll Number
        {wch: 25}, // Program
        {wch: 10}, // Semester
      ];
      ws['!cols'] = wscols;

      XLSX.utils.book_append_sheet(wb, ws, `Room ${roomName}`);
    });

    XLSX.writeFile(wb, `${fileName}.xlsx`);
  };

  const exportSeatingToPDF = (plans: SeatingPlan[], fileName: string, title: string) => {
    const doc = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
    });
    
    // Group by room
    const roomsWithPlans = rooms.filter(r => plans.some(p => p.roomId === r.id));

    if (roomsWithPlans.length === 0) {
      doc.setFontSize(18);
      doc.text(title, 14, 22);
      doc.setFontSize(11);
      doc.setTextColor(100);
      doc.text("No seating data available", 14, 40);
      doc.save(`${fileName}.pdf`);
      return;
    }

    roomsWithPlans.forEach((room, roomIdx) => {
      if (roomIdx > 0) doc.addPage();

      // Header
      doc.setFontSize(18);
      doc.setTextColor(30, 41, 59); // Slate-800
      doc.text(title, 14, 15);
      
      doc.setFontSize(14);
      doc.setTextColor(79, 70, 229); // Indigo-600
      doc.text(`Room: ${room.roomNumber}`, 14, 25);

      doc.setFontSize(9);
      doc.setTextColor(100);
      const dateStr = new Date().toLocaleString();
      doc.text(`Generated on: ${dateStr} • Total Capacity: ${room.capacity}`, 14, 32);

      // Prepare Grid Data
      const maxRows = Math.max(...room.seatsPerColumn);
      const tableHead = Array.from({ length: room.columns }, (_, i) => `Column ${i + 1}`);
      const tableBody: string[][] = [];

      for (let r = 0; r < maxRows; r++) {
        const rowData: string[] = [];
        for (let c = 0; c < room.columns; c++) {
          const plan = plans.find(p => p.roomId === room.id && p.row === r + 1 && p.seat === c + 1);
          if (plan && r < room.seatsPerColumn[c]) {
            const student = students.find(s => s.id === plan.studentId);
            const subject = subjects.find(s => s.id === plan.subjectId);
            rowData.push(`Col ${plan.seat} • Seat ${plan.row}\n${student?.name || 'N/A'}\n${student?.studentId || ''}\nSem: ${student?.semester || ''}\n${subject?.code || ''}`);
          } else if (r < room.seatsPerColumn[c]) {
            rowData.push("Empty");
          } else {
            rowData.push(""); // Not a seat
          }
        }
        tableBody.push(rowData);
      }

      autoTable(doc, {
        startY: 38,
        head: [tableHead],
        body: tableBody,
        theme: 'grid',
        headStyles: { 
          fillColor: [79, 70, 229],
          textColor: [255, 255, 255],
          fontSize: 10,
          halign: 'center'
        },
        styles: { 
          fontSize: 8,
          cellPadding: 3,
          halign: 'center',
          valign: 'middle',
          overflow: 'linebreak',
          cellWidth: 'auto'
        },
        columnStyles: {
          // Distribute width evenly
        },
        didDrawPage: (data) => {
          const str = "Page " + doc.getNumberOfPages();
          doc.setFontSize(10);
          const pageSize = doc.internal.pageSize;
          const pageHeight = pageSize.height ? pageSize.height : pageSize.getHeight();
          doc.text(str, data.settings.margin.left, pageHeight - 10);
        }
      });
    });

    doc.save(`${fileName}.pdf`);
  };

  const exportTimetableToPDF = (data: TimetableEntry[], fileName: string, title: string) => {
    const doc = new jsPDF();
    
    doc.setFontSize(18);
    doc.setTextColor(30, 41, 59);
    doc.text(title, 14, 20);
    
    doc.setFontSize(9);
    doc.setTextColor(100);
    const dateStr = new Date().toLocaleString();
    doc.text(`Generated on: ${dateStr}`, 14, 28);

    // Group by program and semester
    const grouped = data.reduce((acc, t) => {
      const key = `${t.programmeId}-${t.semester}`;
      if (!acc[key]) acc[key] = [];
      acc[key].push(t);
      return acc;
    }, {} as Record<string, TimetableEntry[]>);

    let currentY = 35;

    const sortedGroups = Object.entries(grouped).sort(([keyA], [keyB]) => {
      const progA = programmes.find(p => p.id === keyA.split('-')[0])?.name || "";
      const progB = programmes.find(p => p.id === keyB.split('-')[0])?.name || "";
      if (progA !== progB) return progA.localeCompare(progB);
      return parseInt(keyA.split('-')[1]) - parseInt(keyB.split('-')[1]);
    });

    sortedGroups.forEach(([key, entries], idx) => {
      const [progId, sem] = key.split('-');
      const prog = programmes.find(p => p.id === progId);
      
      if (idx > 0) {
        if (currentY > 240) {
          doc.addPage();
          currentY = 20;
        } else {
          currentY += 10;
        }
      }

      doc.setFontSize(12);
      doc.setTextColor(51, 65, 85);
      doc.setFont("helvetica", "bold");
      doc.text(`${prog?.name} - Semester ${sem}`, 14, currentY);
      currentY += 5;

      const tableHead = ["Sr. No.", "Subject", "Code", "Date", "Time"];
      const tableBody = entries
        .sort((a, b) => a.date.localeCompare(b.date))
        .map((t, i) => {
          const subj = subjects.find(s => s.id === t.subjectId);
          const dateObj = new Date(t.date);
          const day = dateObj.toLocaleDateString('en-GB', { weekday: 'short' });
          return [
            i + 1,
            subj?.name || 'N/A',
            subj?.code || 'N/A',
            `${t.date} (${day})`,
            t.timeSlot
          ];
        });

      autoTable(doc, {
        startY: currentY,
        head: [tableHead],
        body: tableBody,
        theme: 'striped',
        headStyles: { 
          fillColor: [79, 70, 229],
          textColor: [255, 255, 255],
          fontSize: 9
        },
        styles: { 
          fontSize: 8,
          cellPadding: 2
        },
        margin: { left: 14, right: 14 }
      });
      
      currentY = (doc as any).lastAutoTable.finalY + 10;
    });

    const pageCount = doc.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
      doc.setPage(i);
      doc.setFontSize(10);
      doc.setTextColor(150);
      const str = "Page " + i + " of " + pageCount;
      const pageSize = doc.internal.pageSize;
      const pageHeight = pageSize.height ? pageSize.height : pageSize.getHeight();
      doc.text(str, 14, pageHeight - 10);
    }

    doc.save(`${fileName}.pdf`);
  };

  const downloadStudentTemplate = () => {
    const templateData = [
      {
        "S.No": 1,
        "AdmissionNo": "2024001",
        "StudentName": "John Doe",
        "Program": "B.Tech (CSE)",
        "Semester": 1
      },
      {
        "S.No": 2,
        "AdmissionNo": "2024002",
        "StudentName": "Jane Smith",
        "Program": "MBA (Finance)",
        "Semester": 3
      }
    ];
    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, "Student_Upload_Template.xlsx");
  };

  const handleDeleteDoc = async (collectionName: string, id: string) => {
    const label = deleteTarget?.label || 'Item';
    try {
      await deleteDoc(doc(db, collectionName, id));
      setIsDeleteConfirmModalOpen(false);
      setDeleteTarget(null);
      setStatusMessage({ type: 'success', text: `${label} deleted successfully.` });
    } catch (err) {
      setIsDeleteConfirmModalOpen(false);
      setDeleteTarget(null);
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.DELETE, `${collectionName}/${id}`) });
    }
  };

  const confirmDelete = (collection: string, id: string, label: string) => {
    setDeleteTarget({ collection, id, label });
    setIsDeleteConfirmModalOpen(true);
  };

  const handleEditStudent = (student: Student) => {
    setEditingStudent(student);
    setEditStudentData({
      studentId: student.studentId,
      name: student.name,
      programmeId: student.programmeId,
      specialization: student.specialization || "",
      semester: student.semester
    });
    setIsStudentEditModalOpen(true);
  };

  const handleUpdateStudent = async () => {
    if (!editingStudent) return;
    try {
      await setDoc(doc(db, "students", editingStudent.id), editStudentData, { merge: true });
      setIsStudentEditModalOpen(false);
      setStatusMessage({ type: 'success', text: `Student ${editStudentData.name} updated successfully.` });
    } catch (err) {
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.UPDATE, `students/${editingStudent.id}`) });
    }
  };

  const handleEditSubject = (subject: Subject) => {
    setEditingSubject(subject);
    setEditSubjectData({
      programmeId: subject.programmeId,
      semester: subject.semester,
      code: subject.code,
      name: subject.name,
      type: subject.type,
      specialization: subject.specialization || "General"
    });
    setIsSubjectEditModalOpen(true);
  };

  const handleUpdateSubject = async () => {
    if (!editingSubject) return;
    try {
      await setDoc(doc(db, "subjects", editingSubject.id), editSubjectData, { merge: true });
      setIsSubjectEditModalOpen(false);
      setStatusMessage({ type: 'success', text: `Subject ${editSubjectData.name} updated successfully.` });
    } catch (err) {
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.UPDATE, `subjects/${editingSubject.id}`) });
    }
  };

  const confirmDeleteSemester = (progId: string, semester: number) => {
    const progName = programmes.find(p => p.id === progId)?.name || "Unknown Program";
    setDeleteSemesterTarget({ progId, semester, progName });
    setIsDeleteSemesterModalOpen(true);
  };

  const handleDeleteSemester = async () => {
    if (!deleteSemesterTarget) return;
    const { progId, semester } = deleteSemesterTarget;
    try {
      const studentsToDelete = students.filter(s => s.programmeId === progId && s.semester === semester);
      const deletePromises = studentsToDelete.map(s => deleteDoc(doc(db, "students", s.id)));
      await Promise.all(deletePromises);
      setIsDeleteSemesterModalOpen(false);
      setStatusMessage({ type: 'success', text: `All students from Semester ${semester} in ${deleteSemesterTarget.progName} deleted.` });
    } catch (err) {
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.DELETE, "students") });
    }
  };

  const handleClearAllStudents = async () => {
    try {
      const deletePromises = students.map(s => deleteDoc(doc(db, "students", s.id)));
      await Promise.all(deletePromises);
      setIsStudentClearModalOpen(false);
      setStatusMessage({ type: 'success', text: 'All student data cleared successfully.' });
    } catch (err) {
      setIsStudentClearModalOpen(false);
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.DELETE, "students") });
    }
  };

  const handleClearTimetable = async () => {
    try {
      const deletePromises = timetable.map(t => deleteDoc(doc(db, "timetable", t.id)));
      await Promise.all(deletePromises);
      setIsTimetableClearModalOpen(false);
      setStatusMessage({ type: 'success', text: "Timetable cleared successfully." });
    } catch (err) {
      setIsTimetableClearModalOpen(false);
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.DELETE, "timetable") });
    }
  };

  const handleClearSeating = async () => {
    try {
      const deletePromises = seatingPlans.map(p => deleteDoc(doc(db, "seating_plans", p.id)));
      await Promise.all(deletePromises);
      setIsSeatingClearModalOpen(false);
      setStatusMessage({ type: 'success', text: "Seating plans cleared successfully." });
    } catch (err) {
      setIsSeatingClearModalOpen(false);
      setStatusMessage({ type: 'error', text: handleFirestoreError(err, OperationType.DELETE, "seating_plans") });
    }
  };

  // --- Render Sections ---

  const renderDashboard = () => {
    // Ensure unique room count for accuracy
    const uniqueRoomCount = (rooms || []).reduce((acc, current) => {
      const exists = acc.some(r => r.roomNumber.toLowerCase().trim() === current.roomNumber.toLowerCase().trim());
      if (!exists) acc.push(current);
      return acc;
    }, [] as Room[]).length;

    const stats = [
      { label: "Total Students", value: (students || []).length, icon: Users, color: "bg-blue-500" },
      { label: "Programmes", value: (programmes || []).length, icon: BookOpen, color: "bg-indigo-500" },
      { label: "Subjects", value: (subjects || []).length, icon: FileText, color: "bg-emerald-500" },
      { label: "Rooms", value: uniqueRoomCount, icon: DoorOpen, color: "bg-amber-500" },
    ];

    const programmeData = (programmes || []).map(p => ({
      name: p.name,
      students: (students || []).filter(s => s.programmeId === p.id).length
    }));

    return (
      <div className="space-y-8">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
          {stats.map((s, i) => (
            <motion.div
              key={i}
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ delay: i * 0.1 }}
              className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4"
            >
              <div className={`${s.color} p-3 rounded-xl text-white`}>
                <s.icon size={24} />
              </div>
              <div>
                <p className="text-sm font-medium text-slate-500">{s.label}</p>
                <p className="text-2xl font-bold text-slate-800">{s.value}</p>
              </div>
            </motion.div>
          ))}
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
          <Card title="Students per Programme">
            <div className="h-[300px] w-full min-w-0">
              <ResponsiveContainer width="100%" height="100%" minWidth={0}>
                <BarChart data={programmeData}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} />
                  <XAxis dataKey="name" />
                  <YAxis />
                  <Tooltip />
                  <Bar dataKey="students" fill="#6366f1" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </Card>
          
          <Card title="Upcoming Exams">
            <div className="space-y-4">
              {timetable.slice(0, 5).map((t, i) => {
                const prog = programmes.find(p => p.id === t.programmeId);
                const subj = subjects.find(s => s.id === t.subjectId);
                return (
                  <div key={i} className="flex flex-col sm:flex-row sm:items-center justify-between p-4 bg-slate-50 rounded-xl gap-2">
                    <div>
                      <p className="font-semibold text-slate-800">{subj?.name}</p>
                      <p className="text-xs text-slate-500">{prog?.name} • Sem {t.semester}</p>
                    </div>
                    <div className="sm:text-right">
                      <p className="text-sm font-medium text-indigo-600">{t.date}</p>
                      <p className="text-xs text-slate-400">{t.timeSlot}</p>
                    </div>
                  </div>
                );
              })}
            </div>
          </Card>
        </div>
      </div>
    );
  };

  const renderStudents = () => {
    const filteredStudents = (students || []).filter(s => {
      const matchesSearch = (s.name || "").toLowerCase().includes(studentSearchTerm.toLowerCase()) || 
                           (s.studentId || "").toLowerCase().includes(studentSearchTerm.toLowerCase());
      const matchesProg = selectedStudentProg === "All" || s.programmeId === selectedStudentProg;
      const matchesSpec = selectedStudentSpec === "All" || (s.specialization || "General").trim().toLowerCase() === selectedStudentSpec.trim().toLowerCase();
      return matchesSearch && matchesProg && matchesSpec;
    });

    // Grouping for the navigation
    const progSpecMap: Record<string, Set<string>> = {};
    (students || []).forEach(s => {
      if (!progSpecMap[s.programmeId]) progSpecMap[s.programmeId] = new Set();
      const spec = (s.specialization || "General").trim();
      progSpecMap[s.programmeId].add(spec);
    });
    
        const sortedProgrammes = [...programmes].sort((a, b) => a.name.localeCompare(b.name));

    return (
      <div className="space-y-6">
        {/* Summary Card */}
        <Card title="Student Distribution Summary">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            <div className="p-4 bg-indigo-50 rounded-lg border border-indigo-100 flex flex-col justify-center">
              <p className="text-sm text-indigo-600 font-medium mb-1">Overall Total Students</p>
              <p className="text-3xl font-bold text-indigo-900">{studentStats.overallTotal}</p>
            </div>
            {Object.entries(studentStats.byProgram).sort(([a], [b]) => {
              const nameA = programmes.find(p => p.id === a)?.name || "";
              const nameB = programmes.find(p => p.id === b)?.name || "";
              return nameA.localeCompare(nameB);
            }).map(([progId, progData]) => {
              const progName = programmes.find(p => p.id === progId)?.name || "Unknown Program";
              return (
                <div key={progId} className="p-4 bg-white rounded-lg border border-slate-200">
                  <div className="flex justify-between items-start mb-3">
                    <h4 className="font-bold text-slate-800 truncate mr-2" title={progName}>{progName}</h4>
                    <span className="px-2 py-1 bg-slate-100 text-slate-600 text-[10px] font-bold rounded shrink-0">
                      Total: {progData.total}
                    </span>
                  </div>
                  <div className="space-y-3 max-h-48 overflow-y-auto pr-1 custom-scrollbar">
                    {Object.entries(progData.bySemester).sort(([a], [b]) => parseInt(a) - parseInt(b)).map(([sem, semData]) => (
                      <div key={sem} className="pl-3 border-l-2 border-slate-100">
                        <div className="flex justify-between items-center mb-1">
                          <p className="text-xs font-semibold text-slate-700">Semester {sem}</p>
                          <p className="text-[10px] font-medium text-slate-400">({semData.total} students)</p>
                        </div>
                        <div className="flex flex-wrap gap-1.5">
                          {Object.entries(semData.bySpecialization).sort().map(([spec, count]) => (
                            <span key={spec} className="text-[9px] px-1.5 py-0.5 bg-slate-50 text-slate-500 border border-slate-100 rounded">
                              {spec}: {count}
                            </span>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </Card>

        {/* Header & Search */}
        <div className="flex flex-col lg:flex-row gap-4 lg:items-center justify-between">
          <div className="relative w-full lg:w-96">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
            <input
              type="text"
              placeholder="Search students..."
              className="w-full pl-10 pr-4 py-2 bg-white border border-slate-200 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-transparent outline-none"
              value={studentSearchTerm}
              onChange={(e) => setStudentSearchTerm(e.target.value)}
            />
          </div>
          <div className="flex flex-wrap gap-2 sm:gap-3 w-full lg:w-auto justify-start sm:justify-end">
            <button
              onClick={() => setIsStudentClearModalOpen(true)}
              className="px-4 py-2 bg-rose-50 text-rose-600 border border-rose-100 rounded-lg hover:bg-rose-100 transition-colors flex items-center gap-2 text-sm font-semibold"
            >
              <Trash2 size={16} />
              <span className="hidden sm:inline">Clear All</span>
              <span className="sm:hidden">Clear</span>
            </button>
            <button
              onClick={() => setIsStudentUploadModalOpen(true)}
              className="px-4 py-2 bg-white border border-indigo-200 text-indigo-600 rounded-lg hover:bg-indigo-50 transition-colors flex items-center gap-2 text-sm font-semibold"
            >
              <Upload size={16} />
              <span className="hidden sm:inline">Upload Excel</span>
              <span className="sm:hidden">Upload</span>
            </button>
            <button
              onClick={() => exportToExcel(students, "Students_List")}
              className="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition-colors flex items-center gap-2 text-sm font-semibold shadow-md shadow-indigo-100"
            >
              <Download size={16} />
              Export
            </button>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
          {/* Navigation Sidebar */}
          <div className="lg:col-span-1 space-y-4">
            <Card title="Programmes">
              <div className="space-y-1">
                <button
                  onClick={() => { setSelectedStudentProg("All"); setSelectedStudentSpec("All"); }}
                  className={`w-full text-left px-3 py-2 rounded-md transition-colors text-sm font-medium ${
                    selectedStudentProg === "All" ? "bg-indigo-600 text-white" : "text-slate-600 hover:bg-slate-100"
                  }`}
                >
                  All Programmes
                </button>
                {sortedProgrammes.map(prog => (
                  <button
                    key={prog.id}
                    onClick={() => { setSelectedStudentProg(prog.id); setSelectedStudentSpec("All"); }}
                    className={`w-full text-left px-3 py-2 rounded-md transition-colors text-sm font-medium ${
                      selectedStudentProg === prog.id ? "bg-indigo-600 text-white" : "text-slate-600 hover:bg-slate-100"
                    }`}
                  >
                    {prog.name}
                  </button>
                ))}
              </div>
            </Card>

            {selectedStudentProg !== "All" && (
              <Card title="Specializations">
                <div className="space-y-1">
                  <button
                    onClick={() => setSelectedStudentSpec("All")}
                    className={`w-full text-left px-3 py-2 rounded-md transition-colors text-sm font-medium ${
                      selectedStudentSpec === "All" ? "bg-indigo-600 text-white" : "text-slate-600 hover:bg-slate-100"
                    }`}
                  >
                    All Specializations
                  </button>
                  {Array.from(progSpecMap[selectedStudentProg] || []).sort().map(spec => (
                    <button
                      key={spec}
                      onClick={() => setSelectedStudentSpec(spec)}
                      className={`w-full text-left px-3 py-2 rounded-md transition-colors text-sm font-medium ${
                        selectedStudentSpec === spec ? "bg-indigo-600 text-white" : "text-slate-600 hover:bg-slate-100"
                      }`}
                    >
                      {spec}
                    </button>
                  ))}
                </div>
              </Card>
            )}
          </div>

          {/* Student Details Display */}
          <div className="lg:col-span-3 space-y-6">
            <div className="flex items-center justify-between">
              <h2 className="text-xl font-bold text-slate-800">
                {selectedStudentProg === "All" 
                  ? "All Students" 
                  : `${programmes.find(p => p.id === selectedStudentProg)?.name}${selectedStudentSpec !== "All" ? ` - ${selectedStudentSpec}` : ""}`}
                <span className="ml-3 text-sm font-normal text-slate-400">({filteredStudents.length} students)</span>
              </h2>
            </div>

            <div className="max-h-[800px] overflow-y-auto pr-2 custom-scrollbar space-y-6">
              {filteredStudents.length === 0 ? (
                <Card>
                  <div className="text-center py-12">
                    <Users className="mx-auto text-slate-300 mb-4" size={48} />
                    <p className="text-slate-500">No students found for this selection.</p>
                  </div>
                </Card>
              ) : (
                <div className="space-y-6">
                  {Object.entries(
                    filteredStudents.reduce((acc, s) => {
                      const sem = s.semester || 1;
                      if (!acc[sem]) acc[sem] = [];
                      acc[sem].push(s);
                      return acc;
                    }, {} as Record<string, Student[]>)
                  ).sort(([a], [b]) => parseInt(a) - parseInt(b)).map(([sem, semStudents]) => (
                    <div key={sem}>
                      <Card title={
                        <div className="flex items-center justify-between w-full">
                          <span>Semester {sem} ({ (semStudents as Student[]).length } students)</span>
                          {selectedStudentProg !== "All" && (
                            <button
                              onClick={() => confirmDeleteSemester(selectedStudentProg, parseInt(sem))}
                              className="flex items-center gap-1.5 text-xs font-bold text-rose-500 hover:text-rose-700 transition-colors px-2 py-1 bg-rose-50 rounded-lg border border-rose-100"
                              title="Delete all students in this semester"
                            >
                              <Trash2 size={14} />
                              Delete Semester
                            </button>
                          )}
                        </div>
                      }>
                        <div className="overflow-x-auto">
                          <table className="w-full text-left">
                            <thead>
                              <tr className="border-b border-slate-100">
                                <th className="pb-4 font-semibold text-slate-600 text-sm">ID</th>
                                <th className="pb-4 font-semibold text-slate-600 text-sm">Name</th>
                                <th className="pb-4 font-semibold text-slate-600 text-sm">Programme</th>
                                <th className="pb-4 font-semibold text-slate-600 text-sm">Specialization</th>
                                <th className="pb-4 font-semibold text-slate-600 text-sm text-right">Actions</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-50">
                              {(semStudents as Student[]).map(student => (
                                <tr key={student.id} className="hover:bg-slate-50 transition-colors">
                                  <td className="py-4 text-sm font-medium text-slate-700">{student.studentId}</td>
                                  <td className="py-4 text-sm text-slate-600">{student.name}</td>
                                  <td className="py-4 text-sm text-slate-600">
                                    {programmes.find(p => p.id === student.programmeId)?.name || "Unknown"}
                                  </td>
                                  <td className="py-4 text-sm text-slate-600">
                                    <span className="px-2 py-1 bg-slate-100 rounded text-xs font-medium text-slate-600">
                                      {student.specialization || "General"}
                                    </span>
                                  </td>
                                  <td className="py-4 text-right">
                                    <div className="flex items-center justify-end gap-2">
                                      <button
                                        onClick={() => handleEditStudent(student)}
                                        className="text-slate-400 hover:text-indigo-600 transition-colors p-1"
                                        title="Edit Student"
                                      >
                                        <Edit2 size={16} />
                                      </button>
                                      <button
                                        onClick={() => confirmDelete("students", student.id, `Student ${student.name}`)}
                                        className="text-slate-400 hover:text-rose-600 transition-colors p-1"
                                        title="Delete Student"
                                      >
                                        <Trash2 size={16} />
                                      </button>
                                    </div>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </Card>
                    </div>
                  ))}
                  
                  {/* Overall Total at the end */}
                  <div className="p-6 bg-slate-800 text-white rounded-xl shadow-lg flex items-center justify-between">
                    <div className="flex items-center gap-4">
                      <div className="p-3 bg-slate-700 rounded-lg">
                        <Users size={24} />
                      </div>
                      <div>
                        <p className="text-slate-400 text-sm font-medium uppercase tracking-wider">Overall Total Students</p>
                        <p className="text-2xl font-bold">Across All Programs</p>
                      </div>
                    </div>
                    <div className="text-4xl font-black text-indigo-400">
                      {studentStats.overallTotal}
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderProgrammes = () => {
    // Group subjects by programme and then semester
    const groupedData: Record<string, Record<number, Subject[]>> = {};
    
    subjects.forEach(subject => {
      if (!groupedData[subject.programmeId]) {
        groupedData[subject.programmeId] = {};
      }
      if (!groupedData[subject.programmeId][subject.semester]) {
        groupedData[subject.programmeId][subject.semester] = [];
      }
      groupedData[subject.programmeId][subject.semester].push(subject);
    });

    const sortedProgrammes = [...programmes].sort((a, b) => a.name.localeCompare(b.name));

    return (
      <div className="space-y-8">
        {/* Hierarchical View */}
        <Card title="Programmes Structure">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {sortedProgrammes.map(prog => {
              const semesters = groupedData[prog.id] || {};
              const sortedSemesters = Object.keys(semesters).map(Number).sort((a, b) => a - b);

              return (
                <div key={prog.id} className="bg-slate-50 rounded-xl p-5 border border-slate-100 shadow-sm">
                  <div className="flex justify-between items-center mb-4 pb-2 border-b border-slate-200">
                    <h3 className="text-lg font-bold text-indigo-900">{prog.name}</h3>
                    <span className="px-2 py-1 bg-indigo-100 text-indigo-700 text-[10px] font-bold rounded uppercase tracking-wider">
                      {Object.values(semesters).flat().length} Subjects
                    </span>
                  </div>
                  
                  {sortedSemesters.length > 0 ? (
                    <div className="space-y-4 max-h-[300px] overflow-y-auto pr-2 custom-scrollbar">
                      {sortedSemesters.map(sem => (
                        <div key={sem} className="pl-2 border-l-2 border-slate-200">
                          <h4 className="text-sm font-bold text-slate-700 mb-2 flex items-center gap-2">
                            <div className="w-1.5 h-1.5 rounded-full bg-indigo-400"></div>
                            Semester {sem}
                          </h4>
                          <div className="space-y-1.5 ml-3">
                            {semesters[sem].sort((a, b) => a.code.localeCompare(b.code)).map(sub => (
                              <div key={sub.id} className="flex items-start gap-2 group">
                                <div className="mt-1.5 w-2 h-[1px] bg-slate-300"></div>
                                <div className="flex-1">
                                  <div className="flex justify-between items-start">
                                    <p className="text-xs font-medium text-slate-800 leading-tight">{sub.name}</p>
                                    <span className="text-[9px] font-mono text-slate-400 ml-2 shrink-0">{sub.code}</span>
                                  </div>
                                  <p className="text-[9px] text-slate-400 mt-0.5">{sub.type}</p>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  ) : (
                    <div className="py-8 text-center">
                      <BookOpen className="mx-auto text-slate-300 mb-2" size={24} />
                      <p className="text-xs text-slate-400 italic">No subjects added yet</p>
                    </div>
                  )}
                </div>
              );
            })}
            {programmes.length === 0 && (
              <div className="col-span-full py-12 text-center bg-slate-50 rounded-xl border-2 border-dashed border-slate-200">
                <BookOpen className="mx-auto text-slate-300 mb-3" size={48} />
                <p className="text-slate-500 font-medium">No programmes found</p>
                <p className="text-slate-400 text-sm">Add a programme below to get started</p>
              </div>
            )}
          </div>
        </Card>

        {/* Management Section */}
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Manage Programmes */}
          <div className="lg:col-span-1">
            <Card title="Manage Programmes">
              <div className="space-y-4">
                <div className="flex gap-2">
                  <input
                    type="text"
                    placeholder="New Programme Name"
                    className="flex-1 px-3 py-2 text-sm bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    value={newProgName}
                    onChange={(e) => setNewProgName(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleAddProgramme(newProgName)}
                  />
                  <button
                    onClick={() => handleAddProgramme(newProgName)}
                    className="bg-indigo-600 text-white px-4 py-2 rounded-lg text-sm font-semibold hover:bg-indigo-700 shrink-0"
                  >
                    Add
                  </button>
                </div>
                <div className="space-y-2">
                  {sortedProgrammes.map(p => (
                    <div key={p.id} className="p-3 bg-slate-50 rounded-lg flex items-center justify-between group border border-transparent hover:border-slate-200 transition-all">
                      <span className="text-sm font-medium text-slate-700">{p.name}</span>
                      <button 
                        onClick={() => confirmDelete("programmes", p.id, `Programme ${p.name}`)} 
                        className="text-slate-300 hover:text-rose-500 transition-colors"
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            </Card>
          </div>

          {/* Manage Subjects */}
          <div className="lg:col-span-2">
            <Card 
              title="Manage Subjects" 
              action={
                <div className="flex flex-wrap items-center gap-2">
                  <div className="relative hidden md:block lg:hidden xl:block">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
                    <input
                      type="text"
                      placeholder="Search subjects..."
                      className="pl-9 pr-4 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 w-28 sm:w-32"
                      value={subjectSearchTerm}
                      onChange={(e) => setSubjectSearchTerm(e.target.value)}
                    />
                  </div>
                  <div className="flex flex-wrap gap-1.5 sm:gap-2 items-center">
                    <select
                      className="px-2 py-1.5 text-[10px] sm:text-xs bg-slate-50 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 font-medium max-w-[80px] sm:max-w-none"
                      value={subjectProgFilter}
                      onChange={(e) => setSubjectProgFilter(e.target.value)}
                    >
                      <option value="All">All Progs</option>
                      {programmes.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                    <select
                      className="px-2 py-1.5 text-[10px] sm:text-xs bg-slate-50 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 font-medium"
                      value={subjectSemFilter}
                      onChange={(e) => setSubjectSemFilter(e.target.value === "All" ? "All" : Number(e.target.value))}
                    >
                      <option value="All">Sem All</option>
                      {[1, 2, 3, 4, 5, 6, 7, 8].map(sem => <option key={sem} value={sem}>Sem {sem}</option>)}
                    </select>
                    <select
                      className="px-2 py-1.5 text-[10px] sm:text-xs bg-slate-50 border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 font-medium max-w-[80px] sm:max-w-none"
                      value={subjectSpecFilter}
                      onChange={(e) => setSubjectSpecFilter(e.target.value)}
                    >
                      <option value="All">All Specs</option>
                      <option value="General">General</option>
                      <option value="Cyber Security">Cyber Security</option>
                      <option value="AIML">AIML</option>
                    </select>
                    <button
                      onClick={() => setIsSubjectUploadModalOpen(true)}
                      className="flex items-center gap-1 bg-indigo-50 text-indigo-600 px-2 py-1.5 rounded-lg hover:bg-indigo-100 transition-colors text-[9px] sm:text-[10px] font-bold uppercase tracking-wider"
                    >
                      <Upload size={12} />
                      <span className="hidden xs:inline">Upload</span>
                    </button>
                  </div>
                </div>
              }
            >
              <div className="space-y-4">
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-3 bg-slate-50 p-4 rounded-xl border border-slate-100">
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Programme</label>
                    <select 
                      className="w-full px-3 py-2 text-sm bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                      value={newSubjProg}
                      onChange={(e) => setNewSubjProg(e.target.value)}
                    >
                      <option value="">Select...</option>
                      {programmes.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Sem</label>
                    <input 
                      type="number" 
                      className="w-full px-3 py-2 text-sm bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                      value={newSubjSem}
                      onChange={(e) => setNewSubjSem(e.target.value)}
                      placeholder="Sem"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Code</label>
                    <input 
                      type="text" 
                      className="w-full px-3 py-2 text-sm bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                      value={newSubjCode}
                      onChange={(e) => setNewSubjCode(e.target.value)}
                      placeholder="Subj Code"
                    />
                  </div>
                  <div className="space-y-1 sm:col-span-2 md:col-span-1 lg:col-span-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Subject Name</label>
                    <input 
                      type="text" 
                      className="w-full px-3 py-2 text-sm bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                      value={newSubjName}
                      onChange={(e) => setNewSubjName(e.target.value)}
                      placeholder="Name"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Type</label>
                    <select
                      className="w-full px-3 py-2 text-sm bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                      value={newSubjType}
                      onChange={(e) => setNewSubjType(e.target.value as any)}
                    >
                      <option value="Theory">Theory</option>
                      <option value="Practical">Practical</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Specialization</label>
                    <select
                      className="w-full px-3 py-2 text-sm bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                      value={newSubjSpec}
                      onChange={(e) => setNewSubjSpec(e.target.value)}
                    >
                      <option value="General">General</option>
                      {[...new Set(students.filter(s => s.programmeId === newSubjProg).map(s => s.specialization))].filter(Boolean).map(spec => (
                        <option key={spec} value={spec}>{spec}</option>
                      ))}
                    </select>
                  </div>
                  <div className="flex items-end">
                    <button
                      onClick={() => {
                        const sem = Number(newSubjSem);
                        if (newSubjProg && sem > 0 && newSubjCode && newSubjName) {
                          handleAddSubject({ 
                            programmeId: newSubjProg, 
                            semester: sem, 
                            code: newSubjCode.trim(), 
                            name: newSubjName.trim(), 
                            type: newSubjType,
                            specialization: newSubjSpec
                          });
                          setNewSubjCode("");
                          setNewSubjName("");
                        } else {
                          setStatusMessage({ type: 'error', text: "Please fill all subject fields correctly." });
                        }
                      }}
                      className="w-full bg-indigo-600 text-white px-4 py-2 rounded-lg text-sm font-bold hover:bg-indigo-700 h-[38px] transition-all shadow-md shadow-indigo-100 flex items-center justify-center gap-1.5"
                    >
                      <Plus size={16} className="lg:hidden xl:block" />
                      <span className="lg:hidden xl:inline">Add Subject</span>
                      <span className="hidden lg:inline xl:hidden text-xs">Add</span>
                    </button>
                  </div>
                </div>
                <div className="max-h-[600px] overflow-auto border border-slate-100 rounded-lg custom-scrollbar">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-slate-50 sticky top-0 z-10">
                      <tr>
                        <th className="px-4 py-3 font-bold text-slate-600 uppercase text-[10px] tracking-wider">Code</th>
                        <th className="px-4 py-3 font-bold text-slate-600 uppercase text-[10px] tracking-wider">Name</th>
                        <th className="px-4 py-3 font-bold text-slate-600 uppercase text-[10px] tracking-wider">Programme</th>
                        <th className="px-4 py-3 font-bold text-slate-600 uppercase text-[10px] tracking-wider">Sem</th>
                        <th className="px-4 py-3 font-bold text-slate-600 uppercase text-[10px] tracking-wider">Specialization</th>
                        <th className="px-4 py-3 font-bold text-slate-600 uppercase text-[10px] tracking-wider text-right">Actions</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {subjects
                        .filter(s => {
                          const matchesSearch = s.name.toLowerCase().includes(subjectSearchTerm.toLowerCase()) || 
                                              s.code.toLowerCase().includes(subjectSearchTerm.toLowerCase());
                          const matchesSpec = subjectSpecFilter === "All" || (s.specialization || "General") === subjectSpecFilter;
                          const matchesProg = subjectProgFilter === "All" || s.programmeId === subjectProgFilter;
                          const matchesSem = subjectSemFilter === "All" || s.semester === subjectSemFilter;
                          return matchesSearch && matchesSpec && matchesProg && matchesSem;
                        })
                        .sort((a, b) => {
                          // Sort by programme first, then semester, then code
                          if (a.programmeId !== b.programmeId) {
                            const progA = programmes.find(p => p.id === a.programmeId)?.name || "";
                            const progB = programmes.find(p => p.id === b.programmeId)?.name || "";
                            return progA.localeCompare(progB);
                          }
                          if (a.semester !== b.semester) return a.semester - b.semester;
                          return a.code.localeCompare(b.code);
                        })
                        .map((s) => (
                        <tr key={s.id} className="hover:bg-slate-50/50 transition-colors">
                          <td className="px-4 py-3 text-slate-800 font-mono font-medium">{s.code}</td>
                          <td className="px-4 py-3 text-slate-600">{s.name}</td>
                          <td className="px-4 py-3 text-slate-600">{programmes.find(p => p.id === s.programmeId)?.name}</td>
                          <td className="px-4 py-3 text-slate-600">{s.semester}</td>
                          <td className="px-4 py-3 text-slate-600">
                            <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${s.specialization === 'General' || !s.specialization ? 'bg-slate-100 text-slate-600' : 'bg-indigo-100 text-indigo-600'}`}>
                              {s.specialization || 'General'}
                            </span>
                          </td>
                          <td className="px-4 py-3 text-right">
                            <div className="flex justify-end gap-2">
                              <button 
                                onClick={() => handleEditSubject(s)}
                                className="text-slate-300 hover:text-indigo-600 transition-colors"
                              >
                                <Edit2 size={16} />
                              </button>
                              <button onClick={() => confirmDelete("subjects", s.id, `Subject ${s.name}`)} className="text-slate-300 hover:text-rose-500 transition-colors">
                                <Trash2 size={16} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </Card>
          </div>
        </div>
      </div>
    );
  };

  const renderRooms = () => {
    // Filter out any potential duplicates in the local state for display safety
    const uniqueRooms = (rooms || []).reduce((acc, current) => {
      const x = acc.find(item => item.roomNumber.toLowerCase().trim() === current.roomNumber.toLowerCase().trim());
      if (!x) {
        return acc.concat([current]);
      } else {
        return acc;
      }
    }, [] as Room[]);

    const missingPredefined = PREDEFINED_ROOMS.filter(pr => 
      !uniqueRooms.some(r => r.roomNumber.toLowerCase().trim() === pr.roomNumber.toLowerCase().trim())
    );

    const totalCapacity = uniqueRooms.reduce((acc, r) => acc + (r.capacity || 0), 0);

    return (
      <div className="space-y-6">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
          <div className="lg:col-span-1">
            <Card>
              <div className="flex items-center gap-4 p-2">
                <div className="p-3 bg-indigo-100 text-indigo-600 rounded-xl">
                  <DoorOpen size={24} />
                </div>
                <div>
                  <p className="text-sm font-medium text-slate-500 uppercase tracking-wider">Total Rooms</p>
                  <p className="text-2xl font-bold text-slate-800">{uniqueRooms.length}</p>
                </div>
              </div>
            </Card>
          </div>
          <div className="lg:col-span-1">
            <Card>
              <div className="flex items-center gap-4 p-2">
                <div className="p-3 bg-emerald-100 text-emerald-600 rounded-xl">
                  <Users size={24} />
                </div>
                <div>
                  <p className="text-sm font-medium text-slate-500 uppercase tracking-wider">Total Seating Capacity</p>
                  <p className="text-2xl font-bold text-emerald-600">{totalCapacity}</p>
                </div>
              </div>
            </Card>
          </div>
          <div className="lg:col-span-1 flex items-center justify-end gap-3">
            {missingPredefined.length > 0 && (
              <button
                onClick={initializePredefinedRooms}
                disabled={isInitializingRooms}
                className="flex items-center gap-2 bg-emerald-600 text-white px-4 py-2 rounded-lg hover:bg-emerald-700 transition-colors disabled:opacity-50"
              >
                <Zap size={18} className={isInitializingRooms ? "animate-pulse" : ""} />
                {isInitializingRooms ? "Initializing..." : `Initialize ${missingPredefined.length} Missing Rooms`}
              </button>
            )}
            <button
              onClick={() => setIsRoomAddModalOpen(true)}
              className="flex items-center gap-2 bg-indigo-600 text-white px-4 py-2 rounded-lg hover:bg-indigo-700 transition-colors"
            >
              <Plus size={18} />
              Add Room
            </button>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {uniqueRooms.map((room) => (
            <div key={room.id}>
              <Card title={`Room ${room.roomNumber}`}>
                <div className="space-y-4">
                  <div className="grid grid-cols-2 gap-4">
                    <div className="bg-slate-50 p-3 rounded-lg text-center">
                      <p className="text-xs text-slate-500 uppercase font-bold tracking-wider">Columns</p>
                      <p className="text-xl font-bold text-slate-800">{room.columns || 0}</p>
                    </div>
                    <div className="bg-slate-50 p-3 rounded-lg text-center">
                      <p className="text-xs text-slate-500 uppercase font-bold tracking-wider">Capacity</p>
                      <p className="text-xl font-bold text-indigo-600">{room.capacity || 0}</p>
                    </div>
                  </div>
                  <div className="bg-slate-50 p-3 rounded-lg">
                    <p className="text-xs text-slate-500 uppercase font-bold tracking-wider mb-1">Seats Per Column</p>
                    <p className="text-sm text-slate-700 font-mono">
                      {room.seatsPerColumn?.join(", ") || "N/A"}
                    </p>
                  </div>
                  <div className="pt-4 border-t border-slate-100 flex justify-end">
                    <button onClick={() => confirmDelete("rooms", room.id, `Room ${room.roomNumber}`)} className="text-rose-500 hover:text-rose-700">
                      <Trash2 size={18} />
                    </button>
                  </div>
                </div>
              </Card>
            </div>
          ))}
        </div>
      </div>
    );
  };

  const renderSeating = () => {
    const filteredPlans = seatingPlans.filter(p => p.date === selectedSeatingDate && p.timeSlot === selectedSeatingSlot);

    return (
      <div className="space-y-6">
        <Card title="Generate Seating Arrangement">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 items-end">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Select Date</label>
              <input
                type="date"
                className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                value={selectedSeatingDate}
                min={new Date().toISOString().split('T')[0]}
                onChange={(e) => setSelectedSeatingDate(e.target.value)}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Select Time Slot</label>
              <select
                className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                value={selectedSeatingSlot}
                onChange={(e) => setSelectedSeatingSlot(e.target.value)}
              >
                <option value="">Choose Slot</option>
                {[...new Set(timetable.map(t => t.timeSlot))].map(slot => (
                  <option key={slot} value={slot}>{slot}</option>
                ))}
              </select>
            </div>
            <div className="flex flex-col sm:flex-row gap-3">
              <button
                onClick={() => handleGenerateSeating(selectedSeatingDate, selectedSeatingSlot)}
                disabled={!selectedSeatingDate || !selectedSeatingSlot || isGenerating}
                className="flex-1 bg-indigo-600 text-white px-6 py-2 rounded-lg font-semibold hover:bg-indigo-700 transition-colors disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {isGenerating ? (
                  <>
                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                    Generating...
                  </>
                ) : (
                  "Generate Plans"
                )}
              </button>
              {seatingPlans.length > 0 && (
                <button
                  onClick={() => setIsSeatingClearModalOpen(true)}
                  disabled={isGenerating}
                  className="flex items-center justify-center gap-2 bg-rose-50 text-rose-600 px-6 py-2 rounded-lg font-semibold hover:bg-rose-100 transition-colors border border-rose-100 disabled:opacity-50"
                >
                  <Trash2 size={18} />
                  Clear All
                </button>
              )}
            </div>
          </div>
        </Card>

        {filteredPlans.length > 0 ? (
          <div className="space-y-6">
            <div className="flex flex-col md:flex-row md:items-center justify-end gap-3">
              <button
                onClick={() => exportSeatingToExcel(filteredPlans, `Seating_Plan_${selectedSeatingDate}`)}
                className="flex items-center justify-center gap-2 bg-white border border-slate-200 text-slate-700 px-4 py-2 rounded-lg hover:bg-slate-50 w-full md:w-auto"
              >
                <Download size={18} />
                Export Excel
              </button>
              <button
                onClick={() => exportSeatingToPDF(filteredPlans, `Seating_Plan_${selectedSeatingDate}`, `Seating Plan - ${selectedSeatingDate} (${selectedSeatingSlot})`)}
                className="flex items-center justify-center gap-2 bg-indigo-50 text-indigo-600 border border-indigo-100 px-4 py-2 rounded-lg hover:bg-indigo-100 w-full md:w-auto"
              >
                <FileText size={18} />
                Download PDF
              </button>
            </div>

            <div className="grid grid-cols-1 gap-8">
              {rooms.filter(r => filteredPlans.some(p => p.roomId === r.id)).map(room => (
                <div key={room.id}>
                  <Card title={`Room ${room.roomNumber} Seating`}>
                    <div className="overflow-x-auto">
                      <div className="inline-grid gap-4" style={{ gridTemplateColumns: `repeat(${room.columns}, minmax(120px, 1fr))` }}>
                        {Array.from({ length: Math.max(...room.seatsPerColumn) }).map((_, rIdx) => (
                          <React.Fragment key={rIdx}>
                            {Array.from({ length: room.columns }).map((_, cIdx) => {
                              const plan = filteredPlans.find(p => p.roomId === room.id && p.row === rIdx + 1 && p.seat === cIdx + 1);
                              const student = students.find(s => s.id === plan?.studentId);
                              const isSeatValid = rIdx < room.seatsPerColumn[cIdx];

                              if (!isSeatValid) {
                                return <div key={cIdx} className="p-3"></div>; // Spacer for shorter columns
                              }

                              return (
                                <div key={cIdx} className={`p-3 rounded-lg border ${plan ? "bg-indigo-50 border-indigo-200" : "bg-slate-50 border-slate-100 opacity-40"}`}>
                                  {plan ? (
                                    <div className="text-center">
                                      <p className="text-[10px] font-bold text-indigo-400 uppercase">Col {plan.seat} • Seat {plan.row}</p>
                                      <p className="text-sm font-bold text-slate-800 truncate">{student?.name}</p>
                                      <p className="text-[10px] text-slate-500">{student?.studentId}</p>
                                      <p className="text-[10px] font-medium text-indigo-600">Sem {student?.semester}</p>
                                    </div>
                                  ) : (
                                    <div className="text-center py-4">
                                      <p className="text-xs text-slate-400">Empty</p>
                                    </div>
                                  )}
                                </div>
                              );
                            })}
                          </React.Fragment>
                        ))}
                      </div>
                    </div>
                  </Card>
                </div>
              ))}
            </div>
          </div>
        ) : (
          selectedSeatingDate && selectedSeatingSlot && !isDataLoading && (
            <div className="text-center py-12 bg-slate-50 rounded-2xl border-2 border-dashed border-slate-200">
              <div className="p-4 bg-white rounded-full w-16 h-16 flex items-center justify-center mx-auto mb-4 shadow-sm">
                <Grid className="text-slate-300" size={32} />
              </div>
              <h3 className="text-lg font-bold text-slate-800">No Seating Plan Found</h3>
              <p className="text-slate-500 max-w-xs mx-auto mt-2">
                Click "Generate Plans" to create a seating arrangement for this shift.
              </p>
            </div>
          )
        )}
        {isDataLoading && (
          <div className="flex flex-col items-center justify-center py-12">
            <div className="w-10 h-10 border-4 border-indigo-600 border-t-transparent rounded-full animate-spin mb-4" />
            <p className="text-slate-500 font-medium">Loading seating data...</p>
          </div>
        )}
      </div>
    );
  };

  const renderTimetable = () => {
    const groupedTimetable = timetable.reduce((acc, entry) => {
      const key = `${entry.programmeId}-${entry.semester}`;
      if (!acc[key]) acc[key] = [];
      acc[key].push(entry);
      return acc;
    }, {} as Record<string, TimetableEntry[]>);

    const handlePrint = (programmeId: string, semester: number) => {
      const prog = programmes.find(p => p.id === programmeId);
      const entries = groupedTimetable[`${programmeId}-${semester}`].sort((a, b) => a.date.localeCompare(b.date));
      
      const printWindow = window.open('', '_blank');
      if (!printWindow) return;

      const today = new Date().toLocaleDateString('en-GB');
      
      const content = `
        <html>
          <head>
            <title>Timetable - ${prog?.name} Sem ${semester}</title>
            <style>
              body { font-family: Arial, sans-serif; padding: 20px; }
              .header { text-align: center; margin-bottom: 5px; }
              .header h1 { font-size: 24px; margin: 0; }
              .issue-date { text-align: right; font-weight: bold; margin-bottom: 10px; }
              .subheader { text-align: center; border: 1px solid black; background: #f0f0f0; padding: 5px; font-weight: bold; margin-bottom: 0; }
              table { width: 100%; border-collapse: collapse; margin-top: 0; }
              th, td { border: 1px solid black; padding: 8px; text-align: center; }
              th { background: #fff; font-weight: bold; }
              .sr-no { width: 50px; }
              .paper-name { text-align: left; }
              .paper-code { width: 150px; }
              .date-day { width: 150px; }
              .time { width: 180px; }
            </style>
          </head>
          <body>
            <div class="header">
              <h1>Mid Term Theory Examination Time Table - March, 2026</h1>
            </div>
            <div class="issue-date">Date: ${today}</div>
            <div class="subheader">Semester-${prog?.name}-${semester}</div>
            <table>
              <thead>
                <tr>
                  <th class="sr-no">Sr. No.</th>
                  <th>Name of The Paper</th>
                  <th class="paper-code">Paper Code</th>
                  <th class="date-day">Date/ Day</th>
                  <th class="time">Time</th>
                </tr>
              </thead>
              <tbody>
                ${entries.map((t, idx) => {
                  const subj = subjects.find(s => s.id === t.subjectId);
                  const dateObj = new Date(t.date);
                  const day = dateObj.toLocaleDateString('en-GB', { weekday: 'long' });
                  return `
                    <tr>
                      <td>${idx + 1}</td>
                      <td class="paper-name">${subj?.name || 'Unknown'}</td>
                      <td>${subj?.code || 'N/A'}</td>
                      <td>${t.date}<br/>${day}</td>
                      <td>${t.timeSlot}</td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
            <script>window.print();</script>
          </body>
        </html>
      `;
      printWindow.document.write(content);
      printWindow.document.close();
    };

    return (
      <div className="space-y-8">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-center gap-2 overflow-x-auto pb-2 md:pb-0">
            {["All", 1, 2, 3, 4, 5, 6, 7, 8].map((sem) => (
              <button
                key={sem}
                onClick={() => setSelectedSemesterFilter(sem as any)}
                className={`px-4 py-2 rounded-lg text-sm font-medium transition-all whitespace-nowrap ${
                  selectedSemesterFilter === sem
                    ? "bg-indigo-600 text-white shadow-md shadow-indigo-100"
                    : "bg-white text-slate-600 border border-slate-200 hover:border-indigo-300 hover:text-indigo-600"
                }`}
              >
                {sem === "All" ? "All Semesters" : `Sem ${sem}`}
              </button>
            ))}
          </div>
          <div className="flex flex-wrap justify-end gap-3">
            <button
              onClick={() => {
                const data = timetable
                  .sort((a, b) => {
                    const progA = programmes.find(p => p.id === a.programmeId)?.name || "";
                    const progB = programmes.find(p => p.id === b.programmeId)?.name || "";
                    if (progA !== progB) return progA.localeCompare(progB);
                    if (a.semester !== b.semester) return a.semester - b.semester;
                    return a.date.localeCompare(b.date);
                  })
                  .map((t, idx) => {
                    const subj = subjects.find(s => s.id === t.subjectId);
                    const dateObj = new Date(t.date);
                    const day = dateObj.toLocaleDateString('en-GB', { weekday: 'long' });
                    const prog = programmes.find(p => p.id === t.programmeId);
                    return {
                      "Sr. No.": idx + 1,
                      "Programme": prog?.name || 'N/A',
                      "Semester": t.semester,
                      "Name of The Paper": subj?.name || 'N/A',
                      "Paper Code": subj?.code || 'N/A',
                      "Date/ Day": `${t.date} (${day})`,
                      "Time": t.timeSlot
                    };
                  });
                exportToExcel(data, "All_Examination_Timetables");
              }}
              className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-white border border-slate-200 text-indigo-600 px-4 py-2 rounded-lg hover:bg-indigo-50 transition-colors text-sm"
            >
              <Download size={16} />
              Export Excel
            </button>
            <button
              onClick={() => exportTimetableToPDF(timetable, "All_Examination_Timetables", "All Examination Timetables")}
              className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-white border border-slate-200 text-indigo-600 px-4 py-2 rounded-lg hover:bg-indigo-50 transition-colors text-sm"
            >
              <FileText size={16} />
              Export PDF
            </button>
            <button
              onClick={() => setIsTimetableClearModalOpen(true)}
              className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-white border border-slate-200 text-rose-600 px-4 py-2 rounded-lg hover:bg-rose-50 transition-colors text-sm"
            >
              <Trash2 size={16} />
              Clear All
            </button>
            <button
              onClick={() => setIsTimetableAutoModalOpen(true)}
              className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-indigo-600 text-white px-4 py-2 rounded-lg hover:bg-indigo-700 transition-colors text-sm"
            >
              <Zap size={16} />
              Auto Generate
            </button>
          </div>
        </div>

        {(Object.entries(groupedTimetable) as [string, TimetableEntry[]][])
          .filter(([key]) => {
            if (selectedSemesterFilter === "All") return true;
            const [_, sem] = key.split('-');
            return Number(sem) === selectedSemesterFilter;
          })
          .map(([key, entries]) => {
            const [progId, sem] = key.split('-');
            const prog = programmes.find(p => p.id === progId);
            const sortedEntries = [...entries].sort((a, b) => a.date.localeCompare(b.date));

            return (
              <div key={key} className="space-y-4">
                <div className="flex items-center justify-between px-2">
                  <h3 className="text-lg font-bold text-slate-800">
                    {prog?.name} - Semester {sem}
                  </h3>
                  <div className="flex items-center gap-4">
                    <button
                      onClick={() => {
                        const data = sortedEntries.map((t, idx) => {
                          const subj = subjects.find(s => s.id === t.subjectId);
                          const dateObj = new Date(t.date);
                          const day = dateObj.toLocaleDateString('en-GB', { weekday: 'long' });
                          return {
                            "Sr. No.": idx + 1,
                            "Name of The Paper": subj?.name || 'N/A',
                            "Paper Code": subj?.code || 'N/A',
                            "Date/ Day": `${t.date} (${day})`,
                            "Time": t.timeSlot
                          };
                        });
                        exportToExcel(data, `${prog?.name}_Sem_${sem}_Timetable`);
                      }}
                      className="flex items-center gap-2 text-indigo-600 hover:text-indigo-800 font-medium text-sm"
                    >
                      <Download size={16} />
                      Download Excel
                    </button>
                    <button
                      onClick={() => handlePrint(progId, Number(sem))}
                      className="flex items-center gap-2 text-indigo-600 hover:text-indigo-800 font-medium text-sm"
                    >
                      <Printer size={16} />
                      Print Formatted
                    </button>
                  </div>
                </div>
                <Card noPadding>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-slate-50 border-y border-slate-200">
                          <th className="px-4 py-3 font-bold text-slate-700 text-sm border-r border-slate-200 w-16 text-center">Sr. No.</th>
                          <th className="px-4 py-3 font-bold text-slate-700 text-sm border-r border-slate-200">Name of The Paper</th>
                          <th className="px-4 py-3 font-bold text-slate-700 text-sm border-r border-slate-200 w-40 text-center">Paper Code</th>
                          <th className="px-4 py-3 font-bold text-slate-700 text-sm border-r border-slate-200 w-48 text-center">Date/ Day</th>
                          <th className="px-4 py-3 font-bold text-slate-700 text-sm border-r border-slate-200 w-48 text-center">Time</th>
                          <th className="px-4 py-3 font-bold text-slate-700 text-sm text-center w-24">Actions</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {sortedEntries.map((t, idx) => {
                          const subj = subjects.find(s => s.id === t.subjectId);
                          const dateObj = new Date(t.date);
                          const day = dateObj.toLocaleDateString('en-GB', { weekday: 'long' });
                          return (
                            <tr key={t.id} className="hover:bg-slate-50/50 transition-colors">
                              <td className="px-4 py-4 text-sm text-slate-600 border-r border-slate-100 text-center">{idx + 1}</td>
                              <td className="px-4 py-4 text-sm font-medium text-slate-800 border-r border-slate-100">{subj?.name}</td>
                              <td className="px-4 py-4 text-sm text-slate-600 border-r border-slate-100 text-center">{subj?.code}</td>
                              <td className="px-4 py-4 text-sm text-slate-600 border-r border-slate-100 text-center">
                                <div>{t.date}</div>
                                <div className="text-[10px] font-bold text-slate-400 uppercase">{day}</div>
                              </td>
                              <td className="px-4 py-4 text-sm text-slate-600 border-r border-slate-100 text-center">{t.timeSlot}</td>
                              <td className="px-4 py-4 text-center">
                                <div className="flex items-center justify-center gap-2">
                                  <button 
                                    onClick={() => {
                                      const subj = subjects.find(s => s.id === t.subjectId);
                                      confirmDelete("timetable", t.id, `Exam for ${subj?.name || 'Unknown'}`);
                                    }} 
                                    className="text-rose-600 hover:text-rose-800 p-1"
                                    title="Delete"
                                  >
                                    <Trash2 size={16} />
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </Card>
              </div>
            );
          })}
      </div>
    );
  };

  // --- Main Layout --- ---

  if (loading) return <div className="h-screen flex items-center justify-center bg-slate-50">
    <div className="animate-spin rounded-full h-12 w-12 border-4 border-indigo-600 border-t-transparent"></div>
  </div>;

  if (!user) {
    return (
      <ErrorBoundary>
        <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-10 rounded-3xl shadow-xl border border-slate-100 max-w-md w-full text-center space-y-8"
        >
          <div className="bg-indigo-600 w-16 h-16 rounded-2xl flex items-center justify-center text-white mx-auto shadow-lg shadow-indigo-200">
            <Grid size={32} />
          </div>
          <div>
            <h1 className="text-3xl font-bold text-slate-800">SmartExam</h1>
            <p className="text-slate-500 mt-2">Advanced Exam Management System</p>
          </div>
          <button
            onClick={handleLogin}
            className="w-full flex items-center justify-center gap-3 bg-indigo-600 text-white py-4 rounded-2xl font-bold text-lg hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100 active:scale-95"
          >
            <LogIn size={24} />
            Login with Google
          </button>
          <p className="text-xs text-slate-400">
            By logging in, you agree to the terms of service and privacy policy.
          </p>
        </motion.div>
      </div>
      </ErrorBoundary>
    );
  }

  return (
    <ErrorBoundary>
      <div className="h-screen bg-slate-50 flex flex-col md:flex-row overflow-hidden">
      {/* Mobile Header */}
      <header className="md:hidden bg-white border-b border-slate-200 p-4 flex items-center justify-between sticky top-0 z-40">
        <div className="flex items-center gap-3">
          <div className="bg-indigo-600 p-2 rounded-xl text-white">
            <Grid size={20} />
          </div>
          <h1 className="text-lg font-bold text-slate-800 tracking-tight">SmartExam</h1>
        </div>
        <button 
          onClick={() => setIsSidebarOpen(!isSidebarOpen)}
          className="p-2 text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
        >
          {isSidebarOpen ? <X size={24} /> : <Menu size={24} />}
        </button>
      </header>

      {/* Sidebar Overlay */}
      <AnimatePresence>
        {isSidebarOpen && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={() => setIsSidebarOpen(false)}
            className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-40 md:hidden"
          />
        )}
      </AnimatePresence>

      {/* Sidebar */}
      <aside className={`
        fixed inset-y-0 left-0 z-50 bg-white border-r border-slate-200 flex flex-col transition-all duration-300 ease-in-out overflow-x-hidden
        ${isSidebarCollapsed ? "md:w-20" : "md:w-72"}
        md:sticky md:top-0 md:h-screen
        ${isSidebarOpen ? "w-72 translate-x-0" : "w-72 -translate-x-full md:translate-x-0"}
      `}>
        <div className={`flex items-center justify-between p-6 ${isSidebarCollapsed ? "flex-col gap-4 px-2" : ""}`}>
          <div className="flex items-center gap-3 overflow-hidden">
            <div className="bg-indigo-600 p-2 rounded-xl text-white shrink-0">
              <Grid size={24} />
            </div>
            {!isSidebarCollapsed && (
              <h1 className="text-xl font-bold text-slate-800 tracking-tight truncate">SmartExam</h1>
            )}
          </div>
          <button 
            onClick={() => setIsSidebarCollapsed(!isSidebarCollapsed)}
            className="hidden md:flex p-1.5 text-slate-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition-all"
          >
            {isSidebarCollapsed ? <ChevronRight size={20} /> : <ChevronLeft size={20} />}
          </button>
          <button 
            onClick={() => setIsSidebarOpen(false)}
            className="md:hidden p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
          >
            <X size={24} />
          </button>
        </div>

        <nav className={`flex-1 overflow-y-auto overflow-x-hidden px-4 space-y-2 no-scrollbar ${isSidebarCollapsed ? "px-2" : ""}`}>
          <SidebarItem icon={LayoutDashboard} label="Dashboard" active={activeTab === "dashboard"} onClick={() => { setActiveTab("dashboard"); setIsSidebarOpen(false); }} collapsed={isSidebarCollapsed} />
          <SidebarItem icon={BookOpen} label="Programmes" active={activeTab === "programmes"} onClick={() => { setActiveTab("programmes"); setIsSidebarOpen(false); }} collapsed={isSidebarCollapsed} />
          <SidebarItem icon={Users} label="Students" active={activeTab === "students"} onClick={() => { setActiveTab("students"); setIsSidebarOpen(false); }} collapsed={isSidebarCollapsed} />
          <SidebarItem icon={DoorOpen} label="Rooms" active={activeTab === "rooms"} onClick={() => { setActiveTab("rooms"); setIsSidebarOpen(false); }} collapsed={isSidebarCollapsed} />
          <SidebarItem icon={Calendar} label="Timetable" active={activeTab === "timetable"} onClick={() => { setActiveTab("timetable"); setIsSidebarOpen(false); }} collapsed={isSidebarCollapsed} />
          <SidebarItem icon={Grid} label="Seating Plan" active={activeTab === "seating"} onClick={() => { setActiveTab("seating"); setIsSidebarOpen(false); }} collapsed={isSidebarCollapsed} />
        </nav>

        <div className={`p-4 mt-auto border-t border-slate-100 ${isSidebarCollapsed ? "px-2" : ""}`}>
          <div className={`p-3 bg-indigo-50 rounded-2xl space-y-3 ${isSidebarCollapsed ? "p-1.5 bg-transparent" : ""}`}>
            <div className={`flex items-center gap-3 overflow-hidden ${isSidebarCollapsed ? "justify-center" : ""}`}>
              <img src={user.photoURL || ""} alt="" className="w-10 h-10 rounded-full border-2 border-white shadow-sm shrink-0" />
              {!isSidebarCollapsed && (
                <div className="overflow-hidden">
                  <p className="text-sm font-bold text-slate-700 truncate">{user.displayName}</p>
                  <p className="text-xs text-slate-500 truncate">{user.email}</p>
                </div>
              )}
            </div>
            {!isSidebarCollapsed ? (
              <button 
                onClick={handleLogout}
                className="w-full flex items-center justify-center gap-2 py-2 text-xs font-bold text-rose-600 bg-white border border-rose-100 rounded-xl hover:bg-rose-50 transition-colors"
              >
                <LogOut size={14} />
                Logout
              </button>
            ) : (
              <button 
                onClick={handleLogout}
                className="w-full flex items-center justify-center p-2 text-rose-600 hover:bg-rose-50 rounded-xl transition-colors"
              >
                <LogOut size={20} />
              </button>
            )}
          </div>
        </div>
      </aside>

      {/* Content */}
      <main className="flex-1 p-4 md:p-10 overflow-y-auto w-full">
        {connectionError && (
          <div className="mb-6 animate-in fade-in slide-in-from-top-4 duration-300">
            <div className="bg-rose-50 border border-rose-200 p-4 rounded-xl shadow-lg flex gap-3">
              <AlertCircle className="text-rose-600 shrink-0" size={20} />
              <div>
                <p className="text-sm font-bold text-rose-800">Connection Error</p>
                <p className="text-xs text-rose-700 mt-1">{connectionError}</p>
              </div>
            </div>
          </div>
        )}

        {statusMessage && (
          <motion.div
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            className={`mb-6 p-4 rounded-lg flex items-center justify-between ${
              statusMessage.type === 'success' ? 'bg-emerald-50 text-emerald-800 border border-emerald-100' :
              statusMessage.type === 'error' ? 'bg-rose-50 text-rose-800 border border-rose-100' :
              'bg-blue-50 text-blue-800 border border-blue-100'
            }`}
          >
            <div className="flex items-center gap-2">
              <AlertCircle size={18} />
              <p className="text-sm font-medium">{statusMessage.text}</p>
            </div>
            <button onClick={() => setStatusMessage(null)} className="text-slate-400 hover:text-slate-600">
              <X size={18} />
            </button>
          </motion.div>
        )}

        <header className="mb-6 md:mb-10 flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h2 className="text-2xl md:text-3xl font-bold text-slate-800 capitalize">{activeTab.replace("-", " ")}</h2>
            <p className="text-sm md:text-base text-slate-500 mt-1">Manage your college examination process efficiently.</p>
          </div>
        </header>

        <AnimatePresence mode="wait">
          <motion.div
            key={activeTab}
            initial={{ opacity: 0, x: 10 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: -10 }}
            transition={{ duration: 0.2 }}
          >
            {activeTab === "dashboard" && renderDashboard()}
            {activeTab === "students" && renderStudents()}
            {activeTab === "rooms" && renderRooms()}
            {activeTab === "seating" && renderSeating()}
            {activeTab === "timetable" && renderTimetable()}
            
            {activeTab === "programmes" && renderProgrammes()}
          </motion.div>
        </AnimatePresence>

        {/* Global Modals */}
        <Modal isOpen={isStudentUploadModalOpen} onClose={() => setIsStudentUploadModalOpen(false)} title="Upload Students Excel">
          <div className="space-y-6">
            <div className="flex justify-between items-center">
              <p className="text-sm text-slate-600">Download the template to ensure correct format:</p>
              <button
                onClick={downloadStudentTemplate}
                className="flex items-center gap-2 text-indigo-600 hover:text-indigo-800 font-medium text-sm"
              >
                <Download size={16} />
                Download Template
              </button>
            </div>
            <div className="p-8 border-2 border-dashed border-slate-200 rounded-2xl flex flex-col items-center justify-center gap-4 bg-slate-50 relative">
              <div className="p-4 bg-indigo-100 text-indigo-600 rounded-full">
                <Upload size={32} />
              </div>
              <div className="text-center">
                <p className="font-semibold text-slate-800">Click to upload or drag and drop</p>
                <p className="text-sm text-slate-500">Excel files only (.xlsx, .xls)</p>
              </div>
              <input
                type="file"
                accept=".xlsx, .xls"
                className="absolute inset-0 opacity-0 cursor-pointer"
                onChange={(e) => {
                  if (e.target.files?.[0]) {
                    handleExcelUpload(e.target.files[0], "students");
                    setIsStudentUploadModalOpen(false);
                  }
                }}
              />
            </div>
            <div className="bg-amber-50 p-4 rounded-xl border border-amber-100 flex gap-3">
              <AlertCircle className="text-amber-600 shrink-0" size={20} />
              <div className="text-sm text-amber-800">
                <p className="font-bold mb-1">Template Requirements:</p>
                <ul className="list-disc list-inside space-y-1">
                  <li>Columns: <strong>AdmissionNo, StudentName, Program, Semester</strong></li>
                  <li>In <strong>Program</strong>, you can include specialization in brackets: <code>B.Tech (CSE)</code></li>
                  <li>Or use a separate <strong>Specialization</strong> column</li>
                </ul>
              </div>
            </div>
          </div>
        </Modal>

        <Modal isOpen={isStudentEditModalOpen} onClose={() => setIsStudentEditModalOpen(false)} title="Edit Student Record">
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Student ID</label>
              <input
                type="text"
                className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                value={editStudentData.studentId}
                onChange={(e) => setEditStudentData({ ...editStudentData, studentId: e.target.value })}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Full Name</label>
              <input
                type="text"
                className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                value={editStudentData.name}
                onChange={(e) => setEditStudentData({ ...editStudentData, name: e.target.value })}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Programme</label>
              <select
                className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                value={editStudentData.programmeId}
                onChange={(e) => setEditStudentData({ ...editStudentData, programmeId: e.target.value })}
              >
                {programmes.map(p => (
                  <option key={p.id} value={p.id}>{p.name}</option>
                ))}
              </select>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Semester</label>
                <input
                  type="number"
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={editStudentData.semester}
                  onChange={(e) => setEditStudentData({ ...editStudentData, semester: parseInt(e.target.value) })}
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Specialization</label>
                <input
                  type="text"
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={editStudentData.specialization}
                  onChange={(e) => setEditStudentData({ ...editStudentData, specialization: e.target.value })}
                />
              </div>
            </div>
            <div className="flex gap-3 pt-4">
              <button
                onClick={() => setIsStudentEditModalOpen(false)}
                className="flex-1 px-4 py-2 border border-slate-200 text-slate-600 rounded-lg hover:bg-slate-50 transition-colors font-medium"
              >
                Cancel
              </button>
              <button
                onClick={handleUpdateStudent}
                className="flex-1 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 transition-colors font-medium"
              >
                Update Record
              </button>
            </div>
          </div>
        </Modal>

        <Modal isOpen={isSubjectEditModalOpen} onClose={() => setIsSubjectEditModalOpen(false)} title="Edit Subject Details">
          <div className="space-y-4">
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Subject Code</label>
                <input
                  type="text"
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={editSubjectData.code}
                  onChange={(e) => setEditSubjectData({ ...editSubjectData, code: e.target.value })}
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Subject Name</label>
                <input
                  type="text"
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={editSubjectData.name}
                  onChange={(e) => setEditSubjectData({ ...editSubjectData, name: e.target.value })}
                />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Programme</label>
                <select
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={editSubjectData.programmeId}
                  onChange={(e) => setEditSubjectData({ ...editSubjectData, programmeId: e.target.value })}
                >
                  {programmes.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Semester</label>
                <input
                  type="number"
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={editSubjectData.semester}
                  onChange={(e) => setEditSubjectData({ ...editSubjectData, semester: Number(e.target.value) })}
                />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Type</label>
                <select
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={editSubjectData.type}
                  onChange={(e) => setEditSubjectData({ ...editSubjectData, type: e.target.value as any })}
                >
                  <option value="Theory">Theory</option>
                  <option value="Practical">Practical</option>
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Specialization</label>
                <select
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={editSubjectData.specialization}
                  onChange={(e) => setEditSubjectData({ ...editSubjectData, specialization: e.target.value })}
                >
                  <option value="General">General (Core)</option>
                  <option value="Cyber Security">Cyber Security</option>
                  <option value="AIML">AIML</option>
                </select>
              </div>
            </div>
            <div className="flex gap-3 pt-4">
              <button
                onClick={() => setIsSubjectEditModalOpen(false)}
                className="flex-1 px-4 py-2 border border-slate-200 text-slate-600 rounded-lg hover:bg-slate-50 transition-colors font-medium"
              >
                Cancel
              </button>
              <button
                onClick={handleUpdateSubject}
                className="flex-1 px-4 py-2 bg-indigo-600 text-white rounded-lg font-semibold hover:bg-indigo-700 transition-colors"
              >
                Save Changes
              </button>
            </div>
          </div>
        </Modal>

        <Modal isOpen={isDeleteSemesterModalOpen} onClose={() => setIsDeleteSemesterModalOpen(false)} title="Delete Semester Students">
          <div className="space-y-6">
            <div className="flex items-center gap-4 p-4 bg-rose-50 border border-rose-100 rounded-xl text-rose-800">
              <AlertCircle size={24} className="shrink-0" />
              <div className="text-sm font-medium">
                <p>Are you sure you want to delete all students from:</p>
                <p className="mt-1 font-bold">
                  {deleteSemesterTarget?.progName} - Semester {deleteSemesterTarget?.semester}
                </p>
                <p className="mt-2 text-xs opacity-80">This action will remove {(students.filter(s => s.programmeId === deleteSemesterTarget?.progId && s.semester === deleteSemesterTarget?.semester)).length} student records.</p>
              </div>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => setIsDeleteSemesterModalOpen(false)}
                className="flex-1 px-4 py-2 border border-slate-200 text-slate-600 rounded-lg hover:bg-slate-50 transition-colors font-medium"
              >
                Cancel
              </button>
              <button
                onClick={handleDeleteSemester}
                className="flex-1 px-4 py-2 bg-rose-600 text-white rounded-lg hover:bg-rose-700 transition-colors font-medium"
              >
                Delete Semester
              </button>
            </div>
          </div>
        </Modal>

        <Modal isOpen={isStudentClearModalOpen} onClose={() => setIsStudentClearModalOpen(false)} title="Clear All Students">
          <div className="space-y-6">
            <div className="flex items-center gap-4 p-4 bg-rose-50 border border-rose-100 rounded-xl text-rose-800">
              <AlertCircle size={24} className="shrink-0" />
              <p className="text-sm font-medium">
                Are you sure you want to delete <strong>ALL</strong> student data? This action cannot be undone and will remove all student records from the system.
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => setIsStudentClearModalOpen(false)}
                className="flex-1 px-4 py-2 border border-slate-200 text-slate-600 rounded-lg hover:bg-slate-50 transition-colors font-medium"
              >
                Cancel
              </button>
              <button
                onClick={handleClearAllStudents}
                className="flex-1 px-4 py-2 bg-rose-600 text-white rounded-lg hover:bg-rose-700 transition-colors font-medium"
              >
                Yes, Clear All
              </button>
            </div>
          </div>
        </Modal>

        <Modal isOpen={isRoomAddModalOpen} onClose={() => setIsRoomAddModalOpen(false)} title="Add New Room">
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Room Number</label>
              <input
                type="text"
                className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                value={newRoom.roomNumber}
                onChange={(e) => setNewRoom({ ...newRoom, roomNumber: e.target.value })}
                placeholder="e.g. PG-FC1"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Seats Per Column (comma separated)</label>
              <input
                type="text"
                className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                value={seatsPerColumnInput}
                onChange={(e) => {
                  const val = e.target.value;
                  setSeatsPerColumnInput(val);
                  
                  // Only update state if it ends with a number or is empty (allows typing commas/spaces)
                  if (val.trim() === "" || /[0-9]\s*$/.test(val)) {
                    const seats = val.split(",").map(s => parseInt(s.trim())).filter(s => !isNaN(s));
                    setNewRoom({ 
                      ...newRoom, 
                      seatsPerColumn: seats,
                      columns: seats.length,
                      capacity: seats.reduce((acc, curr) => acc + curr, 0)
                    });
                  }
                }}
                placeholder="e.g. 9, 9, 9, 9, 9, 9"
              />
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div className="bg-slate-50 p-3 rounded-lg">
                <p className="text-xs text-slate-500 uppercase font-bold">Columns</p>
                <p className="text-lg font-bold text-slate-800">{newRoom.columns}</p>
              </div>
              <div className="bg-slate-50 p-3 rounded-lg">
                <p className="text-xs text-slate-500 uppercase font-bold">Total Capacity</p>
                <p className="text-lg font-bold text-indigo-600">{newRoom.capacity}</p>
              </div>
            </div>
            <button
              onClick={async () => {
                if (!newRoom.roomNumber.trim()) {
                  setStatusMessage({ type: 'error', text: "Room number is required" });
                  return;
                }
                if (rooms.some(r => r.roomNumber.toLowerCase().trim() === newRoom.roomNumber.toLowerCase().trim())) {
                  setStatusMessage({ type: 'error', text: "Room name must be unique" });
                  return;
                }
                if (newRoom.capacity === 0) {
                  setStatusMessage({ type: 'error', text: "Room must have at least one seat" });
                  return;
                }
                await handleAddRoom(newRoom);
                setIsRoomAddModalOpen(false);
                setNewRoom({ roomNumber: "", capacity: 0, columns: 0, seatsPerColumn: [] });
                setSeatsPerColumnInput("");
                setStatusMessage({ type: 'success', text: "Room added successfully!" });
              }}
              className="w-full bg-indigo-600 text-white py-2 rounded-lg font-semibold hover:bg-indigo-700 transition-colors"
            >
              Save Room
            </button>
          </div>
        </Modal>


        <Modal isOpen={isTimetableClearModalOpen} onClose={() => setIsTimetableClearModalOpen(false)} title="Clear Timetable">
          <div className="space-y-4">
            <div className="p-4 bg-rose-50 border border-rose-100 rounded-lg">
              <p className="text-sm text-rose-800 font-medium">
                Warning: This will permanently delete all scheduled exams from the timetable. This action cannot be undone.
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => setIsTimetableClearModalOpen(false)}
                className="flex-1 px-4 py-2 border border-slate-200 text-slate-700 rounded-lg hover:bg-slate-50 transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={handleClearTimetable}
                className="flex-1 px-4 py-2 bg-rose-600 text-white rounded-lg font-semibold hover:bg-rose-700 transition-colors"
              >
                Clear Everything
              </button>
            </div>
          </div>
        </Modal>

        <Modal isOpen={isSeatingClearModalOpen} onClose={() => setIsSeatingClearModalOpen(false)} title="Clear Seating Plans">
          <div className="space-y-4">
            <div className="p-4 bg-rose-50 border border-rose-100 rounded-lg">
              <p className="text-sm text-rose-800 font-medium">
                Warning: This will permanently delete all generated seating arrangements. This action cannot be undone.
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => setIsSeatingClearModalOpen(false)}
                className="flex-1 px-4 py-2 border border-slate-200 text-slate-700 rounded-lg hover:bg-slate-50 transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={handleClearSeating}
                className="flex-1 px-4 py-2 bg-rose-600 text-white rounded-lg font-semibold hover:bg-rose-700 transition-colors"
              >
                Clear All
              </button>
            </div>
          </div>
        </Modal>

        <Modal isOpen={isDeleteConfirmModalOpen} onClose={() => setIsDeleteConfirmModalOpen(false)} title="Confirm Deletion">
          <div className="space-y-6">
            <div className="flex items-center gap-4 p-4 bg-rose-50 border border-rose-100 rounded-xl text-rose-800">
              <AlertCircle size={24} className="shrink-0" />
              <p className="text-sm font-medium">
                Are you sure you want to delete <strong>{deleteTarget?.label}</strong>? This action cannot be undone.
              </p>
            </div>
            <div className="flex gap-3">
              <button
                onClick={() => setIsDeleteConfirmModalOpen(false)}
                className="flex-1 px-4 py-2 border border-slate-200 text-slate-600 rounded-lg hover:bg-slate-50 transition-colors font-medium"
              >
                Cancel
              </button>
              <button
                onClick={() => deleteTarget && handleDeleteDoc(deleteTarget.collection, deleteTarget.id)}
                className="flex-1 px-4 py-2 bg-rose-600 text-white rounded-lg hover:bg-rose-700 transition-colors font-medium"
              >
                Delete
              </button>
            </div>
          </div>
        </Modal>

        <Modal isOpen={isTimetableAutoModalOpen} onClose={() => setIsTimetableAutoModalOpen(false)} title="Advanced Timetable Generation">
          <div className="space-y-4">
            <div className="p-4 bg-amber-50 border border-amber-100 rounded-lg">
              <p className="text-sm text-amber-800">
                Generate a sequential timetable based on shift capacity and date constraints.
              </p>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Start Date</label>
                <input
                  type="date"
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={autoGenConfig.startDate}
                  min={new Date().toISOString().split('T')[0]}
                  onChange={(e) => setAutoGenConfig({ ...autoGenConfig, startDate: e.target.value })}
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">
                  Preferred End Date (Min: {calculatedMinEndDate || "N/A"})
                </label>
                <input
                  type="date"
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={autoGenConfig.preferredEndDate}
                  min={calculatedMinEndDate}
                  onChange={(e) => setAutoGenConfig({ ...autoGenConfig, preferredEndDate: e.target.value })}
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Mixed Pattern (e.g., 2,1,2)</label>
                <input
                  type="text"
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={autoGenConfig.mixedPattern}
                  onChange={(e) => setAutoGenConfig({ ...autoGenConfig, mixedPattern: e.target.value })}
                  placeholder="2,1,2"
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Public Holidays (YYYY-MM-DD, comma separated)</label>
                <input
                  type="text"
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={autoGenConfig.publicHolidays}
                  onChange={(e) => setAutoGenConfig({ ...autoGenConfig, publicHolidays: e.target.value })}
                  placeholder="2026-03-25, 2026-03-26"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Target Semester</label>
                <select
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={autoGenConfig.semester}
                  onChange={(e) => setAutoGenConfig({ ...autoGenConfig, semester: Number(e.target.value) })}
                >
                  <option value={0}>All Semesters</option>
                  {[1, 2, 3, 4, 5, 6, 7, 8].map(s => <option key={s} value={s}>Semester {s}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-slate-700 mb-1">Exams to Schedule (0 = All)</label>
                <input
                  type="number"
                  className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  value={autoGenConfig.totalExamsToSchedule}
                  onChange={(e) => setAutoGenConfig({ ...autoGenConfig, totalExamsToSchedule: Number(e.target.value) })}
                />
              </div>
            </div>

            <div className="flex items-center justify-between p-3 bg-indigo-50 rounded-lg border border-indigo-100">
              <div className="flex items-center gap-2 text-indigo-700">
                <Users size={16} />
                <span className="text-sm font-bold">Total Seating Capacity</span>
              </div>
              <span className="text-lg font-bold text-indigo-700">
                {(rooms || []).reduce((acc, r) => acc + (r.capacity || 0), 0)}
              </span>
            </div>

            <div className="p-3 bg-slate-50 rounded-lg border border-slate-200">
              <p className="text-xs text-slate-500">
                <strong>Note:</strong> Sundays are automatically excluded from scheduling. Saturdays are included unless listed as a public holiday.
              </p>
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">Time Slots (comma separated)</label>
              <input
                type="text"
                className="w-full px-4 py-2 bg-white border border-slate-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                value={autoGenConfig.slots?.join(", ") || ""}
                onChange={(e) => setAutoGenConfig({ ...autoGenConfig, slots: e.target.value.split(",").map(s => s.trim()).filter(s => s) })}
                placeholder="Morning, Afternoon"
              />
            </div>
            <button
              onClick={handleAutoGenerateTimetable}
              className="w-full bg-indigo-600 text-white py-3 rounded-lg font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100"
            >
              Generate Timetable
            </button>
          </div>
        </Modal>


        <Modal isOpen={isSubjectUploadModalOpen} onClose={() => setIsSubjectUploadModalOpen(false)} title="Upload Subjects Excel">
          <div className="space-y-6">
            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200">
              <h4 className="text-sm font-bold text-slate-700 mb-3 flex items-center gap-2">
                <FileText size={16} className="text-indigo-600" />
                Excel Template Format
              </h4>
              <div className="overflow-x-auto">
                <table className="w-full text-[10px] text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-200">
                      <th className="p-1.5 border border-slate-300">Programme</th>
                      <th className="p-1.5 border border-slate-300">Semester</th>
                      <th className="p-1.5 border border-slate-300">Subject Code</th>
                      <th className="p-1.5 border border-slate-300">Subject Name</th>
                      <th className="p-1.5 border border-slate-300">Subject Type</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className="p-1.5 border border-slate-300">BCA</td>
                      <td className="p-1.5 border border-slate-300">1</td>
                      <td className="p-1.5 border border-slate-300">BCA101</td>
                      <td className="p-1.5 border border-slate-300">Programming in C</td>
                      <td className="p-1.5 border border-slate-300">Theory</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>

            <div className="p-8 border-2 border-dashed border-slate-200 rounded-2xl flex flex-col items-center justify-center gap-4 bg-slate-50 relative">
              <div className="p-4 bg-indigo-100 text-indigo-600 rounded-full">
                <Upload size={32} />
              </div>
              <div className="text-center">
                <p className="font-semibold text-slate-800">Click to upload or drag and drop</p>
                <p className="text-sm text-slate-500">Excel files only (.xlsx, .xls)</p>
              </div>
              <input
                type="file"
                accept=".xlsx, .xls"
                className="absolute inset-0 opacity-0 cursor-pointer"
                onChange={(e) => {
                  if (e.target.files?.[0]) {
                    handleExcelUpload(e.target.files[0], "subjects");
                    setIsSubjectUploadModalOpen(false);
                  }
                }}
              />
            </div>
            
            <div className="flex justify-between items-center">
              <button
                onClick={() => {
                  const templateData = [
                    { "Programme": "BCA", "Semester": 1, "Subject Code": "BCA101", "Subject Name": "Programming in C", "Subject Type": "Theory" }
                  ];
                  exportToExcel(templateData, "Subject_Upload_Template");
                }}
                className="text-indigo-600 text-sm font-semibold flex items-center gap-1 hover:underline"
              >
                <Download size={14} />
                Download Template
              </button>
            </div>
          </div>
        </Modal>
      </main>
    </div>
    </ErrorBoundary>
  );
};
