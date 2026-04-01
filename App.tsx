import React, { useState, useEffect, useMemo, Component } from "react";
import {
  LayoutDashboard,
  Database,
  PlusCircle,
  FileText,
  ShieldAlert,
  Activity,
  CheckCircle2,
  Clock,
  Trash2,
  Edit,
  Download,
  Upload,
  Search,
  BrainCircuit,
  Loader2,
  X,
  Calendar,
  ChevronDown,
  LogOut,
  LogIn,
  FileSpreadsheet,
  AlertCircle,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  Legend,
} from "recharts";
import { dbService } from "./db";
import { analyzeCase } from "./services/geminiService";
import {
  auth,
  db,
  collection,
  doc,
  setDoc,
  deleteDoc,
  onSnapshot,
  query,
  orderBy,
  signInWithPopup,
  signOut,
  googleProvider,
  OperationType,
  handleFirestoreError,
  writeBatch
} from "./firebase";
import { onAuthStateChanged, User } from "firebase/auth";
import {
  Page,
  CaseRecord,
  CaseStatus,
  Student,
  Teacher,
  DatabaseState,
  Attachment,
} from "./types";

class ErrorBoundary extends Component<any, any> {
  state = { hasError: false, error: null };

  constructor(props: any) {
    super(props);
  }

  static getDerivedStateFromError(error: any) {
    return { hasError: true, error };
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen flex items-center justify-center bg-slate-50 p-4">
          <div className="bg-white p-8 rounded-3xl shadow-xl max-w-lg w-full border border-slate-100">
            <h2 className="text-2xl font-bold text-red-600 mb-4">Terjadi Kesalahan</h2>
            <p className="text-slate-600 mb-6">
              Aplikasi mengalami kendala teknis. Silakan muat ulang halaman atau hubungi admin.
            </p>
            <div className="bg-slate-50 p-4 rounded-xl text-xs font-mono text-slate-500 overflow-auto max-h-40 mb-6">
              {(this.state.error as any)?.message}
            </div>
            <button
              onClick={() => window.location.reload()}
              className="w-full bg-indigo-600 text-white py-3 rounded-xl font-semibold hover:bg-indigo-700 transition"
            >
              Muat Ulang Halaman
            </button>
          </div>
        </div>
      );
    }
    return (this as any).props.children;
  }
}

// Components
const Navbar: React.FC<{
  currentPage: Page;
  onPageChange: (p: Page) => void;
  isLoggedIn: boolean;
  onLogout: () => void;
}> = ({ currentPage, onPageChange, isLoggedIn, onLogout }) => (
  <nav className="sticky top-0 z-50 bg-white/80 backdrop-blur-xl border-b border-slate-200 px-6 py-4 flex justify-between items-center shadow-sm">
    <div className="flex items-center gap-4">
      <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white font-bold text-xl shadow-md shadow-indigo-200">
        BK
      </div>
      <h1 className="font-bold text-lg text-slate-800 leading-none uppercase hidden md:block">
        BK Peduli Siswa
      </h1>
    </div>
    <div className="flex gap-2 md:gap-4 text-xs md:text-sm font-semibold overflow-x-auto">
      <button
        onClick={() => onPageChange("dashboard")}
        className={`px-3 py-2 rounded-xl transition flex items-center gap-2 ${currentPage === "dashboard" ? "bg-slate-100 text-indigo-600" : "hover:bg-slate-50 text-slate-600"}`}
      >
        <LayoutDashboard size={18} />{" "}
        <span className="hidden sm:inline">Beranda</span>
      </button>
      <button
        onClick={() => onPageChange("master")}
        className={`px-3 py-2 rounded-xl transition flex items-center gap-2 ${currentPage === "master" ? "bg-slate-100 text-indigo-600" : "hover:bg-slate-50 text-slate-600"}`}
      >
        <Database size={18} /> <span className="hidden sm:inline">Master</span>
      </button>
      <button
        onClick={() => onPageChange("input")}
        className={`px-4 py-2 rounded-xl transition flex items-center gap-2 whitespace-nowrap ${currentPage === "input" ? "bg-indigo-700 text-white shadow-md shadow-indigo-200" : "bg-indigo-600 text-white shadow-md shadow-indigo-100 hover:bg-indigo-700"}`}
      >
        <PlusCircle size={18} /> <span>Input Kasus</span>
      </button>
      <button
        onClick={() => onPageChange("laporan")}
        className={`px-3 py-2 rounded-xl transition flex items-center gap-2 ${currentPage === "laporan" ? "bg-slate-100 text-indigo-600" : "hover:bg-slate-50 text-slate-600"}`}
      >
        <FileText size={18} /> <span className="hidden sm:inline">Laporan</span>
      </button>
      <button
        onClick={() => onPageChange("rekap")}
        className={`px-3 py-2 rounded-xl transition flex items-center gap-2 ${currentPage === "rekap" ? "bg-slate-100 text-indigo-600" : "hover:bg-slate-50 text-slate-600"}`}
      >
        <Activity size={18} /> <span className="hidden sm:inline">Rekap</span>
      </button>
      {isLoggedIn ? (
        <button
          onClick={onLogout}
          className="px-3 py-2 rounded-xl text-red-500 hover:bg-red-50 transition flex items-center gap-2"
        >
          <LogOut size={18} /> <span className="hidden sm:inline">Keluar</span>
        </button>
      ) : (
        <button
          onClick={() => onPageChange("input")}
          className="px-3 py-2 rounded-xl text-indigo-600 hover:bg-indigo-50 transition flex items-center gap-2"
        >
          <LogIn size={18} /> <span className="hidden sm:inline">Login</span>
        </button>
      )}
    </div>
  </nav>
);

const App: React.FC = () => {
  return (
    <ErrorBoundary>
      <AppContent />
    </ErrorBoundary>
  );
};

const AppContent: React.FC = () => {
  const [user, setUser] = useState<User | null>(null);
  const [isAdmin, setIsAdmin] = useState(false);
  const [currentPage, setCurrentPage] = useState<Page>("dashboard");
  const [dbState, setDbState] = useState<DatabaseState>({
    siswa: [],
    wali_kelas: [],
    guru_bk: [],
    kasus: [],
  });
  const [editingCase, setEditingCase] = useState<CaseRecord | null>(null);
  const [loginError, setLoginError] = useState<string | null>(null);
  const [isAuthLoading, setIsAuthLoading] = useState(true);
  const [aiAnalysis, setAiAnalysis] = useState<any>(null);
  const [isAiLoading, setIsAiLoading] = useState(false);
  const [isRestoring, setIsRestoring] = useState(false);

  // Utility to compress image
  const compressImage = (base64Str: string, maxWidth = 800, maxHeight = 800, quality = 0.7): Promise<string> => {
    return new Promise((resolve) => {
      const img = new Image();
      img.src = base64Str;
      img.onload = () => {
        const canvas = document.createElement("canvas");
        let width = img.width;
        let height = img.height;

        if (width > height) {
          if (width > maxWidth) {
            height *= maxWidth / width;
            width = maxWidth;
          }
        } else {
          if (height > maxHeight) {
            width *= maxHeight / height;
            height = maxHeight;
          }
        }

        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx?.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
    });
  };

  // Initialize data and auth
  useEffect(() => {
    const unsubAuth = onAuthStateChanged(auth, (u) => {
      setUser(u);
      if (u && u.email === "wiwikismiati61@guru.smp.belajar.id") {
        setIsAdmin(true);
      } else {
        setIsAdmin(false);
      }
      setIsAuthLoading(false);
    });

    const unsubSiswa = onSnapshot(collection(db, "siswa"), (snapshot) => {
      const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as any));
      setDbState(prev => ({ ...prev, siswa: data }));
    }, (err) => handleFirestoreError(err, OperationType.LIST, "siswa"));

    const unsubWali = onSnapshot(collection(db, "wali_kelas"), (snapshot) => {
      const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as any));
      setDbState(prev => ({ ...prev, wali_kelas: data }));
    }, (err) => handleFirestoreError(err, OperationType.LIST, "wali_kelas"));

    const unsubBK = onSnapshot(collection(db, "guru_bk"), (snapshot) => {
      const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as any));
      setDbState(prev => ({ ...prev, guru_bk: data }));
    }, (err) => handleFirestoreError(err, OperationType.LIST, "guru_bk"));

    const unsubKasus = onSnapshot(query(collection(db, "kasus"), orderBy("created_at", "desc")), (snapshot) => {
      const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as any));
      setDbState(prev => ({ ...prev, kasus: data }));
    }, (err) => handleFirestoreError(err, OperationType.LIST, "kasus"));

    return () => {
      unsubAuth();
      unsubSiswa();
      unsubWali();
      unsubBK();
      unsubKasus();
    };
  }, []);

  const handleLogin = async () => {
    try {
      setLoginError(null);
      await signInWithPopup(auth, googleProvider);
    } catch (error: any) {
      console.error("Login error:", error);
      if (error.code === 'auth/unauthorized-domain') {
        setLoginError("Domain ini belum terdaftar di Authorized Domains Firebase Console.");
      } else {
        setLoginError(error.message || "Gagal login.");
      }
    }
  };

  const handleLogout = async () => {
    if (confirm("Yakin ingin keluar?")) {
      await signOut(auth);
      setCurrentPage("dashboard");
    }
  };

  const handleCaseSubmit = async (record: Omit<CaseRecord, "created_at">) => {
    const id = editingCase?.id ? String(editingCase.id) : `case_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const newRecord = {
      ...record,
      created_at: editingCase ? editingCase.created_at : Date.now(),
      uid: user?.uid || "anonymous"
    };

    // Check size
    const size = JSON.stringify(newRecord).length;
    if (size > 1000000) {
      alert("Gagal menyimpan: Ukuran data (termasuk lampiran) melebihi batas 1MB. Silakan kurangi jumlah atau ukuran lampiran.");
      return;
    }

    try {
      await setDoc(doc(db, "kasus", id), newRecord);
      setEditingCase(null);
      setCurrentPage("laporan");
      alert("Data berhasil disimpan!");
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, `kasus/${id}`);
    }
  };

  const handleDeleteCase = async (id: string | number) => {
    if (confirm("Apakah Anda yakin ingin menghapus data ini?")) {
      try {
        await deleteDoc(doc(db, "kasus", String(id)));
      } catch (err) {
        handleFirestoreError(err, OperationType.DELETE, `kasus/${id}`);
      }
    }
  };

  const handleEditCase = (record: CaseRecord) => {
    setEditingCase(record);
    setCurrentPage("input");
  };

  const handleStatusUpdate = async (id: string | number, status: CaseStatus) => {
    const item = dbState.kasus.find((k) => k.id === id);
    if (item) {
      try {
        await setDoc(doc(db, "kasus", String(id)), { ...item, status });
      } catch (err) {
        handleFirestoreError(err, OperationType.UPDATE, `kasus/${id}`);
      }
    }
  };

  const handleAiAnalysis = async (kronologi: string, kategori: string) => {
    setIsAiLoading(true);
    setAiAnalysis(null);
    try {
      const result = await analyzeCase(kronologi, kategori);
      setAiAnalysis(result);
    } catch (error: any) {
      if (error.message.includes("GEMINI_API_KEY")) {
        alert("KONFIGURASI ERROR: GEMINI_API_KEY tidak ditemukan. Silakan atur di Environment Variables pada pengaturan proyek Vercel Anda.");
      } else {
        alert("Terjadi kesalahan saat menghubungi layanan AI. Coba lagi nanti.");
        console.error("AI Analysis error:", error);
      }
    } finally {
      setIsAiLoading(false);
    }
  };

  // Views
  if (isAuthLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <div className="text-center">
          <Loader2 className="w-10 h-10 text-indigo-600 animate-spin mx-auto mb-4" />
          <p className="text-slate-500 font-medium">Memuat sistem...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen pb-20">
      <Navbar
        currentPage={currentPage}
        onPageChange={(p) => {
          if (p === "input" && !editingCase) {
            setEditingCase(null);
          }
          setCurrentPage(p);
        }}
        isLoggedIn={!!user}
        onLogout={handleLogout}
      />

      <main className="max-w-7xl mx-auto p-4 md:p-8">
        {currentPage === "dashboard" && <DashboardView dbState={dbState} />}
        
        {(currentPage === "master" || currentPage === "input") && !isAdmin && (
          <div className="flex items-center justify-center p-4 mt-10">
            <div className="bg-white w-full max-w-md rounded-3xl shadow-xl p-10 border border-slate-100">
              <div className="text-center mb-8">
                <div className="w-16 h-16 bg-indigo-600 rounded-2xl mx-auto mb-6 flex items-center justify-center text-white text-2xl font-bold shadow-lg shadow-indigo-200">
                  BK
                </div>
                <h2 className="text-2xl font-bold text-slate-900 tracking-tight">
                  Login Admin
                </h2>
                <p className="text-slate-500 text-sm mt-2">
                  Silakan login dengan akun Google Admin
                </p>
              </div>
              <div className="space-y-5">
                {loginError && (
                  <div className="text-red-500 text-xs font-medium text-center bg-red-50 p-3 rounded-lg border border-red-100">
                    {loginError}
                  </div>
                )}
                {user && !isAdmin && (
                  <div className="text-amber-600 text-xs font-medium text-center bg-amber-50 p-3 rounded-lg border border-amber-100">
                    Anda masuk sebagai <span className="font-bold">{user.email}</span>, tetapi akun ini bukan Admin.
                  </div>
                )}
                <button
                  onClick={handleLogin}
                  className="w-full bg-indigo-600 text-white py-3.5 rounded-xl font-semibold hover:bg-indigo-700 transition duration-200 mt-2 shadow-md shadow-indigo-200 flex items-center justify-center gap-2"
                >
                  <ShieldAlert size={18} /> Masuk dengan Google
                </button>
                {user && (
                  <button
                    onClick={handleLogout}
                    className="w-full bg-slate-100 text-slate-600 py-3.5 rounded-xl font-semibold hover:bg-slate-200 transition duration-200 flex items-center justify-center gap-2"
                  >
                    <LogOut size={18} /> Keluar
                  </button>
                )}
              </div>
            </div>
          </div>
        )}

        {currentPage === "master" && isAdmin && (
          <MasterView dbState={dbState} onRefresh={() => {}} isRestoring={isRestoring} setIsRestoring={setIsRestoring} />
        )}
        {currentPage === "input" && isAdmin && (
          <InputView
            dbState={dbState}
            editingCase={editingCase}
            onSubmit={handleCaseSubmit}
            onCancel={() => {
              setEditingCase(null);
              setCurrentPage("laporan");
            }}
            onAiAnalysis={handleAiAnalysis}
            onClearAiAnalysis={() => setAiAnalysis(null)}
            aiAnalysis={aiAnalysis}
            isAiLoading={isAiLoading}
            compressImage={compressImage}
          />
        )}
        {currentPage === "laporan" && (
          <ReportView
            dbState={dbState}
            isLoggedIn={isAdmin}
            onEdit={handleEditCase}
            onDelete={handleDeleteCase}
            onStatusUpdate={handleStatusUpdate}
          />
        )}
        {currentPage === "rekap" && (
          <RekapView dbState={dbState} />
        )}
      </main>
    </div>
  );
};

// Sub-Views
const DashboardView: React.FC<{ dbState: DatabaseState }> = ({ dbState }) => {
  const [filterNama, setFilterNama] = useState("");
  const [filterKelas, setFilterKelas] = useState("");
  const [filterStatus, setFilterStatus] = useState("");

  const stats = useMemo(
    () => ({
      total: dbState.kasus.length,
      baru: dbState.kasus.filter((k) => k.status === CaseStatus.BARU).length,
      proses: dbState.kasus.filter((k) => k.status === CaseStatus.PROSES)
        .length,
      selesai: dbState.kasus.filter((k) => k.status === CaseStatus.SELESAI)
        .length,
    }),
    [dbState.kasus],
  );

  const filteredKasus = useMemo(() => {
    return dbState.kasus.filter((k) => {
      const matchNama = k.nama_siswa
        .toLowerCase()
        .includes(filterNama.toLowerCase());
      const matchKelas = k.kelas
        .toLowerCase()
        .includes(filterKelas.toLowerCase());
      const matchStatus = filterStatus ? k.status === filterStatus : true;
      return matchNama && matchKelas && matchStatus;
    });
  }, [dbState.kasus, filterNama, filterKelas, filterStatus]);

  const groupedData = useMemo(() => {
    const groups: Record<string, any> = {};
    filteredKasus.forEach((k) => {
      if (!groups[k.nama_siswa]) groups[k.nama_siswa] = {};
      if (!groups[k.nama_siswa][k.kelas]) groups[k.nama_siswa][k.kelas] = {};
      if (!groups[k.nama_siswa][k.kelas][k.status])
        groups[k.nama_siswa][k.kelas][k.status] = {};
      if (!groups[k.nama_siswa][k.kelas][k.status][k.kategori_kasus])
        groups[k.nama_siswa][k.kelas][k.status][k.kategori_kasus] = [];

      groups[k.nama_siswa][k.kelas][k.status][k.kategori_kasus].push({
        kronologi: k.kronologi,
        tindak_lanjut: k.tindak_lanjut,
        tanggal: k.tanggal,
      });
    });
    return groups;
  }, [filteredKasus]);

  const chartData = [
    { name: "Baru", count: stats.baru, color: "#3b82f6" },
    { name: "Proses", count: stats.proses, color: "#f59e0b" },
    { name: "Selesai", count: stats.selesai, color: "#10b981" },
  ];

  const categoryData = useMemo(() => {
    const counts: Record<string, number> = {};
    dbState.kasus.forEach((k) => {
      counts[k.kategori_kasus] = (counts[k.kategori_kasus] || 0) + 1;
    });
    return Object.entries(counts).map(([name, value]) => ({ name, value }));
  }, [dbState.kasus]);

  const COLORS = [
    "#6366f1",
    "#a855f7",
    "#ec4899",
    "#f43f5e",
    "#10b981",
    "#f59e0b",
  ];

  return (
    <div className="space-y-6 md:space-y-8 animate-in fade-in duration-500">
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 md:gap-6">
        <div className="glass-card p-5 md:p-6 flex justify-between items-center group hover:shadow-md transition-all">
          <div>
            <p className="text-[10px] md:text-xs font-semibold text-slate-500 uppercase tracking-wider">
              Total Kasus
            </p>
            <h2 className="text-3xl md:text-4xl font-light mt-1 md:mt-2 text-slate-800">
              {stats.total}
            </h2>
          </div>
          <div className="w-10 h-10 md:w-12 md:h-12 rounded-full bg-slate-50 flex items-center justify-center text-slate-500">
            <Database size={20} className="md:w-6 md:h-6" />
          </div>
        </div>
        <div className="glass-card p-5 md:p-6 flex justify-between items-center group hover:shadow-md transition-all">
          <div>
            <p className="text-[10px] md:text-xs font-semibold text-slate-500 uppercase tracking-wider">
              Antrean Kasus
            </p>
            <h2 className="text-3xl md:text-4xl font-light mt-1 md:mt-2 text-slate-800">
              {stats.baru}
            </h2>
          </div>
          <div className="w-10 h-10 md:w-12 md:h-12 rounded-full bg-blue-50 flex items-center justify-center text-blue-500">
            <Clock size={20} className="md:w-6 md:h-6" />
          </div>
        </div>
        <div className="glass-card p-5 md:p-6 flex justify-between items-center group hover:shadow-md transition-all">
          <div>
            <p className="text-[10px] md:text-xs font-semibold text-slate-500 uppercase tracking-wider">
              Dalam Proses
            </p>
            <h2 className="text-3xl md:text-4xl font-light mt-1 md:mt-2 text-slate-800">
              {stats.proses}
            </h2>
          </div>
          <div className="w-10 h-10 md:w-12 md:h-12 rounded-full bg-amber-50 flex items-center justify-center text-amber-500">
            <Activity size={20} className="md:w-6 md:h-6" />
          </div>
        </div>
        <div className="glass-card p-5 md:p-6 flex justify-between items-center group hover:shadow-md transition-all">
          <div>
            <p className="text-[10px] md:text-xs font-semibold text-slate-500 uppercase tracking-wider">
              Selesai
            </p>
            <h2 className="text-3xl md:text-4xl font-light mt-1 md:mt-2 text-slate-800">
              {stats.selesai}
            </h2>
          </div>
          <div className="w-10 h-10 md:w-12 md:h-12 rounded-full bg-emerald-50 flex items-center justify-center text-emerald-500">
            <CheckCircle2 size={20} className="md:w-6 md:h-6" />
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 md:gap-8">
        <div className="glass-card p-5 md:p-8 min-h-[350px] md:min-h-[400px]">
          <h3 className="text-sm md:text-base font-semibold mb-6 md:mb-8 text-slate-800 flex items-center gap-2">
            <LayoutDashboard size={18} className="text-indigo-500" /> Statistik
            Penanganan
          </h3>
          <div className="h-[250px] md:h-[300px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData}>
                <CartesianGrid
                  strokeDasharray="3 3"
                  vertical={false}
                  stroke="#f1f5f9"
                />
                <XAxis
                  dataKey="name"
                  axisLine={false}
                  tickLine={false}
                  tick={{ fill: "#64748b", fontSize: 12 }}
                />
                <YAxis
                  axisLine={false}
                  tickLine={false}
                  tick={{ fill: "#64748b", fontSize: 12 }}
                />
                <Tooltip
                  cursor={{ fill: "#f8fafc" }}
                  contentStyle={{
                    borderRadius: "12px",
                    border: "none",
                    boxShadow: "0 10px 15px -3px rgba(0,0,0,0.1)",
                  }}
                />
                <Bar dataKey="count" radius={[8, 8, 0, 0]}>
                  {chartData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="glass-card p-5 md:p-8 min-h-[350px] md:min-h-[400px]">
          <h3 className="text-sm md:text-base font-semibold mb-6 md:mb-8 text-slate-800 flex items-center gap-2">
            <ShieldAlert size={18} className="text-indigo-500" /> Distribusi
            Kategori Kasus
          </h3>
          <div className="h-[250px] md:h-[300px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie
                  data={categoryData}
                  innerRadius={60}
                  outerRadius={80}
                  paddingAngle={5}
                  dataKey="value"
                >
                  {categoryData.map((entry, index) => (
                    <Cell
                      key={`cell-${index}`}
                      fill={COLORS[index % COLORS.length]}
                    />
                  ))}
                </Pie>
                <Tooltip
                  contentStyle={{
                    borderRadius: "12px",
                    border: "none",
                    boxShadow: "0 10px 15px -3px rgba(0,0,0,0.1)",
                  }}
                />
                <Legend
                  layout="vertical"
                  align="right"
                  verticalAlign="middle"
                  iconType="circle"
                />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="glass-card p-5 md:p-8">
        <h3 className="text-lg font-bold text-slate-900 mb-6">
          Report Tindak Lanjut Siswa Bermasalah
        </h3>

        <div className="flex flex-col sm:flex-row gap-4 mb-8">
          <div className="relative flex-1">
            <Search
              size={18}
              className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
            />
            <input
              type="text"
              placeholder="Cari Nama Siswa..."
              className="w-full pl-10 pr-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 focus:ring-2 focus:ring-indigo-500 outline-none text-sm transition-all"
              value={filterNama}
              onChange={(e) => setFilterNama(e.target.value)}
            />
          </div>
          <div className="relative w-full sm:w-48">
            <input
              type="text"
              placeholder="Filter Kelas..."
              className="w-full px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-200 focus:ring-2 focus:ring-indigo-500 outline-none text-sm transition-all"
              value={filterKelas}
              onChange={(e) => setFilterKelas(e.target.value)}
            />
          </div>
          <select
            className="bg-slate-50 px-4 py-2.5 rounded-xl text-sm font-medium border border-slate-200 outline-none text-slate-700 focus:ring-2 focus:ring-indigo-500 transition-all w-full sm:w-48"
            value={filterStatus}
            onChange={(e) => setFilterStatus(e.target.value)}
          >
            <option value="">Semua Status</option>
            {[CaseStatus.BARU, CaseStatus.PROSES, CaseStatus.SELESAI].map(
              (s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ),
            )}
          </select>
        </div>

        <div className="border border-slate-200 rounded-2xl overflow-hidden">
          <div className="bg-[#86a361] text-white p-3 font-bold text-sm">
            Kasus Siswa
          </div>
          <div className="divide-y divide-slate-100 max-h-[400px] overflow-y-auto">
            {Object.keys(groupedData).length === 0 ? (
              <div className="p-10 text-center text-slate-400 italic text-sm">
                Tidak ada data ditemukan
              </div>
            ) : (
              Object.entries(groupedData).map(([nama, kelasGroup]) => (
                <div key={nama} className="bg-white">
                  <div className="bg-[#c5d9a8] p-2 pl-4 text-white font-bold text-sm flex items-center gap-2">
                    <span className="w-4 h-4 border border-white/50 flex items-center justify-center text-[10px] leading-none">
                      -
                    </span>
                    {nama}
                  </div>
                  {Object.entries(kelasGroup).map(([kelas, statusGroup]) => (
                    <div key={kelas}>
                      <div className="bg-[#e2efd9] p-2 pl-8 text-slate-800 font-bold text-xs flex items-center gap-2">
                        <span className="w-3.5 h-3.5 border border-slate-400 flex items-center justify-center text-[8px] leading-none">
                          -
                        </span>
                        {kelas}
                      </div>
                      {Object.entries(statusGroup).map(
                        ([status, kategoriGroup]) => (
                          <div key={status}>
                            <div className="p-2 pl-12 text-slate-700 font-bold text-xs flex items-center gap-2">
                              <span className="w-3 h-3 border border-slate-400 flex items-center justify-center text-[8px] leading-none">
                                -
                              </span>
                              {status}
                            </div>
                            {Object.entries(kategoriGroup).map(
                              ([kategori, items]) => (
                                <div key={kategori}>
                                  <div className="p-2 pl-16 text-slate-600 font-bold text-xs flex items-center gap-2">
                                    <span className="w-3 h-3 border border-slate-400 flex items-center justify-center text-[8px] leading-none">
                                      -
                                    </span>
                                    {kategori}
                                  </div>
                                  {(items as any[]).map((item, idx) => (
                                    <div key={idx} className="space-y-0">
                                      <div className="p-2 pl-20 text-slate-600 text-xs flex items-start gap-2">
                                        <span className="w-3 h-3 mt-0.5 border border-slate-400 flex items-center justify-center text-[8px] leading-none shrink-0">
                                          -
                                        </span>
                                        <span className="italic">
                                          {item.kronologi}
                                        </span>
                                      </div>
                                      <div className="p-2 pl-24 text-slate-600 text-xs flex items-start gap-2">
                                        <span className="w-3 h-3 mt-0.5 border border-slate-400 flex items-center justify-center text-[8px] leading-none shrink-0">
                                          -
                                        </span>
                                        <span className="font-medium">
                                          {item.tindak_lanjut}
                                        </span>
                                      </div>
                                      <div className="p-2 pl-28 text-slate-500 text-[10px] font-mono">
                                        {item.tanggal}
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              ),
                            )}
                          </div>
                        ),
                      )}
                    </div>
                  ))}
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

const MasterView: React.FC<{
  dbState: DatabaseState;
  onRefresh: () => void;
  isRestoring: boolean;
  setIsRestoring: (v: boolean) => void;
}> = ({ dbState, onRefresh, isRestoring, setIsRestoring }) => {
  const [restorePreview, setRestorePreview] = useState<{
    counts: { siswa: number; wali_kelas: number; guru_bk: number; kasus: number };
    total: number;
    data: DatabaseState;
  } | null>(null);
  const [importPreview, setImportPreview] = useState<{
    counts: { siswa: number; wali_kelas: number; guru_bk: number };
    total: number;
    wb: any;
  } | null>(null);
  const [statusMsg, setStatusMsg] = useState<{ type: 'success' | 'error', text: string } | null>(null);

  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: "binary" });
      
      const counts = {
        siswa: wb.Sheets["Siswa"] ? XLSX.utils.sheet_to_json(wb.Sheets["Siswa"]).length : 0,
        wali_kelas: wb.Sheets["WaliKelas"] ? XLSX.utils.sheet_to_json(wb.Sheets["WaliKelas"]).length : 0,
        guru_bk: wb.Sheets["GuruBK"] ? XLSX.utils.sheet_to_json(wb.Sheets["GuruBK"]).length : 0,
      };
      const total = counts.siswa + counts.wali_kelas + counts.guru_bk;

      setImportPreview({ counts, total, wb });
    };
    reader.readAsBinaryString(file);
  };

  const executeImport = async () => {
    if (!importPreview) return;
    const { wb } = importPreview;
    setImportPreview(null);
    setIsRestoring(true);

    const sheets = {
      Siswa: "siswa",
      WaliKelas: "wali_kelas",
      GuruBK: "guru_bk",
    };

    try {
      for (const [sheetName, storeName] of Object.entries(sheets)) {
        const ws = wb.Sheets[sheetName];
        if (ws) {
          const data = XLSX.utils.sheet_to_json(ws);
          let batch = writeBatch(db);
          let count = 0;
          
          for (const item of data as any[]) {
            const id = `master_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
            batch.set(doc(db, storeName, id), item);
            count++;
            
            if (count === 500) {
              await batch.commit();
              batch = writeBatch(db);
              count = 0;
            }
          }
          if (count > 0) await batch.commit();
        }
      }
      setStatusMsg({ type: 'success', text: "Data Master berhasil diimpor!" });
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "master_import");
    } finally {
      setIsRestoring(false);
    }
  };

  const handleBackup = () => {
    const data = JSON.stringify(dbState);
    const blob = new window.Blob([data], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Backup_BK_${new Date().toISOString().split("T")[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleRestore = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data: DatabaseState = JSON.parse(evt.target?.result as string);
        
        const counts = {
          siswa: (data.siswa || []).length,
          wali_kelas: (data.wali_kelas || []).length,
          guru_bk: (data.guru_bk || []).length,
          kasus: (data.kasus || []).length
        };
        const total = counts.siswa + counts.wali_kelas + counts.guru_bk + counts.kasus;

        setRestorePreview({ counts, total, data });
      } catch (err) {
        setStatusMsg({ type: 'error', text: "Gagal membaca file backup. Pastikan format file benar." });
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const executeRestore = async () => {
    if (!restorePreview) return;
    const { data } = restorePreview;
    setRestorePreview(null);
    setIsRestoring(true);

    try {
      const stores = ["siswa", "wali_kelas", "guru_bk", "kasus"];
      let skippedAttachments = 0;
      let totalRestored = 0;

      for (const store of stores) {
        const items = (data as any)[store] || [];
        let batch = writeBatch(db);
        let count = 0;

        for (const item of items as any[]) {
          const id = item.id ? String(item.id) : `restored_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
          
          const size = JSON.stringify(item).length;
          if (size > 1000000) {
            if (item.lampiran && item.lampiran.length > 0) {
              item.lampiran = [];
              skippedAttachments++;
            }
          }
          
          batch.set(doc(db, store, id), item);
          count++;
          totalRestored++;

          if (count === 500) {
            await batch.commit();
            batch = writeBatch(db);
            count = 0;
          }
        }
        if (count > 0) await batch.commit();
      }
      
      let msg = `Database berhasil direstore (${totalRestored} dokumen)!`;
      if (skippedAttachments > 0) {
        msg += `\n\nCatatan: ${skippedAttachments} dokumen memiliki lampiran yang terlalu besar dan telah dikosongkan.`;
      }
      setStatusMsg({ type: 'success', text: msg });
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "restore");
    } finally {
      setIsRestoring(false);
    }
  };

  return (
    <div className="relative">
      {/* Loading Overlay */}
      {isRestoring && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[100] flex flex-col items-center justify-center p-6 text-center">
          <div className="w-16 h-16 border-4 border-indigo-500 border-t-transparent rounded-full animate-spin mb-6"></div>
          <h3 className="text-xl font-bold text-white mb-2">Sedang Memproses Data...</h3>
          <p className="text-slate-300 text-sm max-w-xs">Mohon tunggu sebentar, sistem sedang melakukan sinkronisasi dengan database.</p>
        </div>
      )}

      {/* Restore Preview Modal */}
      {restorePreview && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[90] flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl max-w-md w-full overflow-hidden animate-in zoom-in-95 duration-300">
            <div className="p-6 md:p-8">
              <div className="w-16 h-16 bg-indigo-50 text-indigo-600 rounded-2xl flex items-center justify-center mb-6 mx-auto">
                <Database size={32} />
              </div>
              <h3 className="text-xl font-bold text-slate-900 text-center mb-2">Konfirmasi Restore</h3>
              <p className="text-slate-500 text-center text-sm mb-6">File backup terdeteksi dengan rincian berikut:</p>
              
              <div className="bg-slate-50 rounded-2xl p-4 mb-6 space-y-2">
                <div className="flex justify-between text-sm">
                  <span className="text-slate-500">Data Siswa</span>
                  <span className="font-bold text-slate-800">{restorePreview.counts.siswa}</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-slate-500">Data Wali Kelas</span>
                  <span className="font-bold text-slate-800">{restorePreview.counts.wali_kelas}</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-slate-500">Data Guru BK</span>
                  <span className="font-bold text-slate-800">{restorePreview.counts.guru_bk}</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-slate-500">Data Kasus</span>
                  <span className="font-bold text-slate-800">{restorePreview.counts.kasus}</span>
                </div>
                <div className="pt-2 mt-2 border-t border-slate-200 flex justify-between font-bold text-indigo-600">
                  <span>Total Data</span>
                  <span>{restorePreview.total}</span>
                </div>
              </div>

              <p className="text-[10px] text-amber-600 bg-amber-50 p-3 rounded-xl border border-amber-100 mb-6 italic">
                * Data dengan ID yang sama akan diperbarui. Proses ini tidak dapat dibatalkan.
              </p>

              <div className="flex gap-3">
                <button 
                  onClick={() => setRestorePreview(null)}
                  className="flex-1 py-3 rounded-xl bg-slate-100 text-slate-600 font-bold hover:bg-slate-200 transition-colors"
                >
                  Batal
                </button>
                <button 
                  onClick={executeRestore}
                  className="flex-1 py-3 rounded-xl bg-indigo-600 text-white font-bold hover:bg-indigo-700 transition-colors shadow-lg shadow-indigo-100"
                >
                  Lanjutkan
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Import Preview Modal */}
      {importPreview && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[90] flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl max-w-md w-full overflow-hidden animate-in zoom-in-95 duration-300">
            <div className="p-6 md:p-8">
              <div className="w-16 h-16 bg-emerald-50 text-emerald-600 rounded-2xl flex items-center justify-center mb-6 mx-auto">
                <FileSpreadsheet size={32} />
              </div>
              <h3 className="text-xl font-bold text-slate-900 text-center mb-2">Konfirmasi Import Excel</h3>
              <p className="text-slate-500 text-center text-sm mb-6">Data master terdeteksi dalam file:</p>
              
              <div className="bg-slate-50 rounded-2xl p-4 mb-6 space-y-2">
                <div className="flex justify-between text-sm">
                  <span className="text-slate-500">Siswa</span>
                  <span className="font-bold text-slate-800">{importPreview.counts.siswa}</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-slate-500">Wali Kelas</span>
                  <span className="font-bold text-slate-800">{importPreview.counts.wali_kelas}</span>
                </div>
                <div className="flex justify-between text-sm">
                  <span className="text-slate-500">Guru BK</span>
                  <span className="font-bold text-slate-800">{importPreview.counts.guru_bk}</span>
                </div>
                <div className="pt-2 mt-2 border-t border-slate-200 flex justify-between font-bold text-emerald-600">
                  <span>Total Data Master</span>
                  <span>{importPreview.total}</span>
                </div>
              </div>

              <div className="flex gap-3">
                <button 
                  onClick={() => setImportPreview(null)}
                  className="flex-1 py-3 rounded-xl bg-slate-100 text-slate-600 font-bold hover:bg-slate-200 transition-colors"
                >
                  Batal
                </button>
                <button 
                  onClick={executeImport}
                  className="flex-1 py-3 rounded-xl bg-emerald-600 text-white font-bold hover:bg-emerald-700 transition-colors shadow-lg shadow-emerald-100"
                >
                  Impor Sekarang
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Status Message Modal */}
      {statusMsg && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-[110] flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl max-w-sm w-full overflow-hidden animate-in zoom-in-95 duration-300">
            <div className="p-6 md:p-8 text-center">
              <div className={`w-16 h-16 rounded-2xl flex items-center justify-center mb-6 mx-auto ${statusMsg.type === 'success' ? 'bg-emerald-50 text-emerald-600' : 'bg-red-50 text-red-600'}`}>
                {statusMsg.type === 'success' ? <CheckCircle2 size={32} /> : <AlertCircle size={32} />}
              </div>
              <h3 className="text-lg font-bold text-slate-900 mb-2">
                {statusMsg.type === 'success' ? 'Berhasil!' : 'Terjadi Kesalahan'}
              </h3>
              <p className="text-slate-500 text-sm mb-8 whitespace-pre-wrap">{statusMsg.text}</p>
              <button 
                onClick={() => setStatusMsg(null)}
                className={`w-full py-3 rounded-xl font-bold text-white transition-colors ${statusMsg.type === 'success' ? 'bg-emerald-600 hover:bg-emerald-700' : 'bg-red-600 hover:bg-red-700'}`}
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}
      <div className="space-y-6 animate-in slide-in-from-bottom-4 duration-500">
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 md:gap-6">
        <div className="glass-card p-5 md:p-8">
          <h3 className="text-base md:text-lg font-semibold text-slate-800 mb-2 flex items-center gap-2">
            <Upload size={20} className="text-indigo-500" /> Import Data Master
            (Excel)
          </h3>
          <p className="text-[10px] md:text-xs text-slate-500 mb-4 bg-slate-50 p-3 md:p-4 rounded-xl border border-slate-100">
            Gunakan file Excel dengan Sheet: <b>Siswa</b>, <b>WaliKelas</b>,{" "}
            <b>GuruBK</b>.<br />
            Kolom Siswa: Nama, Kelas. Kolom lainnya: Nama.
            <span className="text-red-500 font-medium mt-1 md:mt-2 block">
              *Data lama akan ditimpa!
            </span>
          </p>
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleExcelImport}
            className="block w-full text-xs md:text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 md:file:py-2.5 md:file:px-6 file:rounded-xl file:border-0 file:text-[10px] md:file:text-xs file:font-semibold file:bg-indigo-50 file:text-indigo-600 cursor-pointer hover:file:bg-indigo-100 transition-all"
          />
        </div>
        <div className="glass-card p-5 md:p-8">
          <h3 className="text-base md:text-lg font-semibold text-slate-800 mb-2 flex items-center gap-2">
            <Database size={20} className="text-indigo-500" /> Backup & Restore
          </h3>
          <p className="text-[10px] md:text-xs text-slate-500 mb-4">
            Amankan seluruh data kasus dan data master ke file JSON.
          </p>
          <div className="flex flex-col sm:flex-row gap-3 md:gap-4 mt-2">
            <button
              onClick={handleBackup}
              className="flex-1 bg-slate-800 text-white py-2.5 md:py-3 rounded-xl text-[10px] md:text-xs font-semibold hover:bg-slate-700 transition flex items-center justify-center gap-2"
            >
              <Download size={16} /> Backup DB
            </button>
            <label className="flex-1 bg-white border border-slate-200 text-slate-700 py-2.5 md:py-3 rounded-xl text-[10px] md:text-xs font-semibold text-center cursor-pointer hover:bg-slate-50 transition flex items-center justify-center gap-2">
              <Upload size={16} /> Restore{" "}
              <input
                type="file"
                accept=".json"
                className="hidden"
                onChange={handleRestore}
              />
            </label>
          </div>
        </div>
      </div>

      <div className="glass-card p-5 md:p-8">
        <h4 className="font-semibold text-xs md:text-sm mb-4 md:mb-6 text-slate-500">
          Preview Ketersediaan Data
        </h4>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 md:gap-4 text-center">
          <div className="bg-slate-50 p-4 md:p-6 rounded-2xl border border-slate-100 group hover:bg-slate-100 transition-colors">
            <span className="block text-2xl md:text-3xl font-light text-slate-800 mb-1">
              {dbState.siswa.length}
            </span>
            <span className="text-[10px] md:text-xs font-medium text-slate-500">
              Siswa Terdaftar
            </span>
          </div>
          <div className="bg-slate-50 p-4 md:p-6 rounded-2xl border border-slate-100 group hover:bg-slate-100 transition-colors">
            <span className="block text-2xl md:text-3xl font-light text-slate-800 mb-1">
              {dbState.wali_kelas.length}
            </span>
            <span className="text-[10px] md:text-xs font-medium text-slate-500">
              Wali Kelas
            </span>
          </div>
          <div className="bg-slate-50 p-4 md:p-6 rounded-2xl border border-slate-100 group hover:bg-slate-100 transition-colors">
            <span className="block text-2xl md:text-3xl font-light text-slate-800 mb-1">
              {dbState.guru_bk.length}
            </span>
            <span className="text-[10px] md:text-xs font-medium text-slate-500">Guru BK</span>
          </div>
        </div>
      </div>
    </div>
  </div>
  );
};

const KATEGORI_OPTIONS = [
  "Kedisiplinan",
  "Etika",
  "Akademik",
  "Bullying",
  "Perkelahian",
  "Merokok",
  "Narkoba",
];

const TINDAK_LANJUT_OPTIONS = [
  "Konseling Individu",
  "Konseling Kelompok",
  "Panggilan Orang Tua",
  "Mediasi",
  "Home Visit",
  "Skorsing",
];

const InputView: React.FC<{
  dbState: DatabaseState;
  editingCase: CaseRecord | null;
  onSubmit: (r: any) => void;
  onCancel: () => void;
  onAiAnalysis: (kron: string, kat: string) => void;
  onClearAiAnalysis: () => void;
  aiAnalysis: any;
  isAiLoading: boolean;
  compressImage: (base64: string) => Promise<string>;
}> = ({
  dbState,
  editingCase,
  onSubmit,
  onCancel,
  onAiAnalysis,
  onClearAiAnalysis,
  aiAnalysis,
  isAiLoading,
  compressImage,
}) => {
  const [formData, setFormData] = useState<Partial<CaseRecord>>({
    tanggal: new Date().toISOString().split("T")[0],
    kategori_kasus: "",
    kelas: "",
    nama_siswa: "",
    guru_kelas: "",
    guru_bk: "",
    kronologi: "",
    tindak_lanjut: "Konseling Individu",
    status: CaseStatus.BARU,
    lampiran: [],
  });

  const [isKategoriLainnya, setIsKategoriLainnya] = useState(false);
  const [isTindakLanjutLainnya, setIsTindakLanjutLainnya] = useState(false);

  useEffect(() => {
    if (editingCase) {
      setFormData(editingCase);
      if (editingCase.kategori_kasus && !KATEGORI_OPTIONS.includes(editingCase.kategori_kasus)) {
        setIsKategoriLainnya(true);
      } else {
        setIsKategoriLainnya(false);
      }
      if (editingCase.tindak_lanjut && !TINDAK_LANJUT_OPTIONS.includes(editingCase.tindak_lanjut)) {
        setIsTindakLanjutLainnya(true);
      } else {
        setIsTindakLanjutLainnya(false);
      }
    } else {
      setIsKategoriLainnya(false);
      setIsTindakLanjutLainnya(false);
    }
  }, [editingCase]);

  const CLASSES = useMemo(() => {
    const list: string[] = [];
    ["7", "8", "9"].forEach((grade) => {
      ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((section) => {
        list.push(grade + section);
      });
    });
    return list;
  }, []);

  // Updated to handle case-insensitive Excel properties (Nama/nama, Kelas/kelas)
  const filteredStudents = useMemo(() => {
    if (!formData.kelas) return [];
    return dbState.siswa.filter((s) => {
      const sKelas = (s as any).Kelas || (s as any).kelas;
      return (
        sKelas &&
        sKelas.toString().toUpperCase() === formData.kelas?.toUpperCase()
      );
    });
  }, [formData.kelas, dbState.siswa]);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    files.forEach((f: any) => {
      const reader = new FileReader();
      reader.onload = async (ev) => {
        let data = ev.target?.result as string;
        
        // Compress if it's an image
        if (f.type.startsWith("image/")) {
          try {
            data = await compressImage(data);
          } catch (err) {
            console.error("Compression failed", err);
          }
        }

        setFormData((prev) => ({
          ...prev,
          lampiran: [
            ...(prev.lampiran || []),
            {
              name: f.name,
              type: f.type,
              data: data,
            },
          ],
        }));
      };
      reader.readAsDataURL(f);
    });
  };

  const removeFile = (idx: number) => {
    setFormData((prev) => ({
      ...prev,
      lampiran: prev.lampiran?.filter((_, i) => i !== idx),
    }));
  };

  const inputClass =
    "w-full p-3 md:p-4 rounded-xl border border-slate-200 bg-slate-50 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-400 text-slate-800 appearance-none text-sm md:text-base";

  return (
    <div className="max-w-5xl mx-auto animate-in zoom-in-95 duration-300">
      <form
        onSubmit={(e) => {
          e.preventDefault();
          onSubmit(formData);
        }}
        className="glass-card p-5 md:p-8 lg:p-12 relative overflow-hidden"
      >
        <div className="absolute top-0 left-0 w-full h-1.5 bg-indigo-500"></div>

        {editingCase && (
          <div className="absolute top-5 right-5 md:top-8 md:right-12 bg-amber-50 text-amber-700 px-3 py-1 md:px-4 md:py-1.5 rounded-full text-[10px] md:text-xs font-semibold uppercase tracking-wider border border-amber-200 animate-pulse">
            Edit Mode
          </div>
        )}

        <div className="mb-6 md:mb-10">
          <h3 className="text-2xl md:text-3xl font-bold text-slate-900 tracking-tight">
            Input Detail Kasus
          </h3>
          <p className="text-slate-500 text-sm md:text-base mt-1 md:mt-2">
            Lengkapi data laporan siswa secara objektif dan akurat.
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 md:gap-x-12 gap-y-6 md:gap-y-12">
          {/* Row 1 */}
          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              TANGGAL KEJADIAN
            </label>
            <div className="relative group">
              <input
                type="date"
                required
                className={inputClass + " pr-10 md:pr-12"}
                value={formData.tanggal}
                onChange={(e) =>
                  setFormData({ ...formData, tanggal: e.target.value })
                }
              />
              <Calendar
                className="absolute right-4 md:right-5 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-indigo-500 transition-colors pointer-events-none w-4 h-4 md:w-5 md:h-5"
              />
            </div>
          </div>

          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              KATEGORI KASUS
            </label>
            <div className="relative">
              <select
                required
                className={inputClass}
                value={isKategoriLainnya ? "Lainnya" : formData.kategori_kasus}
                onChange={(e) => {
                  if (e.target.value === "Lainnya") {
                    setIsKategoriLainnya(true);
                    setFormData({ ...formData, kategori_kasus: "" });
                  } else {
                    setIsKategoriLainnya(false);
                    setFormData({ ...formData, kategori_kasus: e.target.value });
                  }
                }}
              >
                <option value="">-- Pilih Jenis Kasus --</option>
                {[...KATEGORI_OPTIONS, "Lainnya"].map((k) => (
                  <option key={k} value={k}>
                    {k}
                  </option>
                ))}
              </select>
              <ChevronDown
                className="absolute right-4 md:right-5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none w-4 h-4 md:w-5 md:h-5"
              />
            </div>
            {isKategoriLainnya && (
              <input
                type="text"
                required
                placeholder="Masukkan kategori kasus lainnya"
                className={`${inputClass} mt-2`}
                value={formData.kategori_kasus}
                onChange={(e) =>
                  setFormData({ ...formData, kategori_kasus: e.target.value })
                }
              />
            )}
          </div>

          {/* Row 2 */}
          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              KELAS
            </label>
            <div className="relative">
              <select
                required
                className={inputClass}
                value={formData.kelas}
                onChange={(e) =>
                  setFormData({
                    ...formData,
                    kelas: e.target.value,
                    nama_siswa: "",
                  })
                }
              >
                <option value="">-- Pilih Kelas --</option>
                {CLASSES.map((k) => (
                  <option key={k} value={k}>
                    {k}
                  </option>
                ))}
              </select>
              <ChevronDown
                className="absolute right-4 md:right-5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none w-4 h-4 md:w-5 md:h-5"
              />
            </div>
          </div>

          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              NAMA SISWA
            </label>
            <div className="relative">
              {/* Changed from Input list to Select for better visibility of data */}
              <select
                required
                className={inputClass}
                value={formData.nama_siswa}
                onChange={(e) =>
                  setFormData({ ...formData, nama_siswa: e.target.value })
                }
                disabled={!formData.kelas}
              >
                {!formData.kelas ? (
                  <option value="">Pilih kelas terlebih dahulu...</option>
                ) : filteredStudents.length === 0 ? (
                  <option value="">
                    Tidak ada data siswa di kelas ini (Cek Master)
                  </option>
                ) : (
                  <>
                    <option value="">-- Pilih Nama Siswa --</option>
                    {filteredStudents.map((s, i) => {
                      const sNama = (s as any).Nama || (s as any).nama;
                      return (
                        <option key={i} value={sNama}>
                          {sNama}
                        </option>
                      );
                    })}
                  </>
                )}
              </select>
              <ChevronDown
                className="absolute right-4 md:right-5 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none w-4 h-4 md:w-5 md:h-5"
              />
            </div>
          </div>

          {/* Row 3 */}
          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              WALI KELAS
            </label>
            <div className="relative">
              <input
                list="wali-list"
                className={inputClass}
                placeholder="Masukkan nama Wali Kelas"
                value={formData.guru_kelas}
                onChange={(e) =>
                  setFormData({ ...formData, guru_kelas: e.target.value })
                }
              />
              <datalist id="wali-list">
                {dbState.wali_kelas.map((t, i) => (
                  <option key={i} value={t.Nama || (t as any).nama} />
                ))}
              </datalist>
            </div>
          </div>

          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              GURU BK
            </label>
            <div className="relative">
              <input
                list="bk-list"
                required
                className={inputClass}
                placeholder="Masukkan nama Guru BK"
                value={formData.guru_bk}
                onChange={(e) =>
                  setFormData({ ...formData, guru_bk: e.target.value })
                }
              />
              <datalist id="bk-list">
                {dbState.guru_bk.map((t, i) => (
                  <option key={i} value={t.Nama || (t as any).nama} />
                ))}
              </datalist>
            </div>
          </div>
        </div>

        {/* Kronologi */}
        <div className="space-y-2 md:space-y-4 mt-8 md:mt-12">
          <div className="flex flex-col sm:flex-row sm:justify-between sm:items-end mb-2 gap-2">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              KRONOLOGI KEJADIAN
            </label>
            <button
              type="button"
              onClick={() =>
                onAiAnalysis(
                  formData.kronologi || "",
                  formData.kategori_kasus || "",
                )
              }
              disabled={!formData.kronologi || isAiLoading}
              className="text-[10px] md:text-xs font-semibold uppercase tracking-wider bg-indigo-50 text-indigo-600 px-4 md:px-5 py-2 md:py-2.5 rounded-xl flex items-center justify-center gap-2 hover:bg-indigo-100 transition-all disabled:opacity-50 border border-indigo-100 w-full sm:w-auto"
            >
              {isAiLoading ? (
                <Loader2 size={16} className="animate-spin" />
              ) : (
                <BrainCircuit size={16} />
              )}
              Dapatkan Analisis AI
            </button>
          </div>
          <textarea
            rows={6}
            required
            placeholder="Tuliskan urutan kejadian secara lengkap dan objektif..."
            className="w-full p-4 md:p-6 rounded-2xl border border-slate-200 bg-slate-50 focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none resize-none text-sm md:text-base leading-relaxed transition-all placeholder:text-slate-400 text-slate-800"
            value={formData.kronologi}
            onChange={(e) =>
              setFormData({ ...formData, kronologi: e.target.value })
            }
          ></textarea>
        </div>

        {aiAnalysis && (
          <div className="p-5 md:p-8 bg-indigo-50/50 rounded-2xl border border-indigo-100 animate-in fade-in slide-in-from-top-4 duration-500 relative mt-6 md:mt-8">
            <button
              type="button"
              onClick={onClearAiAnalysis}
              className="absolute top-4 right-4 md:top-6 md:right-6 text-indigo-400 hover:text-indigo-600 transition-colors"
            >
              <X size={20} className="md:w-6 md:h-6" />
            </button>
            <h4 className="text-lg md:text-xl font-bold text-indigo-900 mb-4 md:mb-6 flex items-center gap-2 md:gap-3">
              <BrainCircuit size={20} className="text-indigo-600 md:w-6 md:h-6" /> Rekomendasi
              Psikologis Gemini AI
            </h4>
            <div className="space-y-4 md:space-y-6 text-xs md:text-sm leading-relaxed text-slate-700">
              <div>
                <span className="font-semibold text-indigo-600 uppercase text-[10px] md:text-[11px] tracking-wider block mb-1">
                  Analisis Kasus:
                </span>{" "}
                {aiAnalysis.analisis}
              </div>
              <div>
                <span className="font-semibold text-indigo-600 uppercase text-[10px] md:text-[11px] tracking-wider block mb-1">
                  Pendekatan Disarankan:
                </span>{" "}
                {aiAnalysis.pendekatan}
              </div>
              <div>
                <span className="font-semibold text-indigo-600 uppercase text-[10px] md:text-[11px] tracking-wider block mb-1">
                  Langkah Strategis:
                </span>
                <ul className="list-disc ml-4 md:ml-6 mt-1 md:mt-2 space-y-1 md:space-y-2">
                  {aiAnalysis.saran?.map((s: string, i: number) => (
                    <li key={i}>{s}</li>
                  ))}
                </ul>
              </div>
            </div>
          </div>
        )}

        <div className="grid grid-cols-1 md:grid-cols-2 gap-x-6 md:gap-x-8 gap-y-6 md:gap-y-8 mt-8 md:mt-12">
          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              TINDAK LANJUT UTAMA
            </label>
            <div className="relative">
              <select
                className={inputClass}
                value={isTindakLanjutLainnya ? "Lainnya" : formData.tindak_lanjut}
                onChange={(e) => {
                  if (e.target.value === "Lainnya") {
                    setIsTindakLanjutLainnya(true);
                    setFormData({ ...formData, tindak_lanjut: "" });
                  } else {
                    setIsTindakLanjutLainnya(false);
                    setFormData({ ...formData, tindak_lanjut: e.target.value });
                  }
                }}
              >
                {[...TINDAK_LANJUT_OPTIONS, "Lainnya"].map((k) => (
                  <option key={k} value={k}>
                    {k}
                  </option>
                ))}
              </select>
              <ChevronDown
                className="absolute right-4 md:right-5 top-1/2 -translate-y-1/2 text-slate-300 pointer-events-none w-4 h-4 md:w-5 md:h-5"
              />
            </div>
            {isTindakLanjutLainnya && (
              <input
                type="text"
                required
                placeholder="Masukkan tindak lanjut lainnya"
                className={`${inputClass} mt-2`}
                value={formData.tindak_lanjut}
                onChange={(e) =>
                  setFormData({ ...formData, tindak_lanjut: e.target.value })
                }
              />
            )}
          </div>
          <div className="space-y-2 md:space-y-4">
            <label className="text-[10px] md:text-xs font-semibold uppercase text-slate-500 tracking-wider block ml-1">
              STATUS LAPORAN
            </label>
            <div className="flex flex-col sm:flex-row gap-2 p-1.5 bg-slate-100/50 rounded-xl border border-slate-200">
              {[CaseStatus.BARU, CaseStatus.PROSES, CaseStatus.SELESAI].map(
                (s) => (
                  <button
                    key={s}
                    type="button"
                    onClick={() => setFormData({ ...formData, status: s })}
                    className={`flex-1 py-2.5 md:py-3 rounded-lg text-[10px] md:text-xs font-semibold uppercase tracking-wider transition-all ${formData.status === s ? "bg-white shadow-sm text-indigo-600" : "text-slate-500 hover:text-slate-700 hover:bg-slate-50"}`}
                  >
                    {s}
                  </button>
                ),
              )}
            </div>
          </div>
        </div>

        <div className="p-5 md:p-8 bg-slate-50/50 rounded-2xl border border-dashed border-slate-300 group hover:border-indigo-300 transition-all mt-8 md:mt-12">
          <label className="block text-[10px] md:text-xs font-semibold mb-2 md:mb-4 uppercase text-slate-500 tracking-wider ml-1">
            LAMPIRAN BUKTI FISIK
          </label>
          <input
            type="file"
            multiple
            onChange={handleFileChange}
            className="block w-full text-xs md:text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 md:file:py-2.5 md:file:px-6 file:rounded-xl file:border-0 file:bg-white file:text-indigo-600 file:font-semibold file:text-[10px] md:file:text-xs file:uppercase file:tracking-wider file:shadow-sm hover:file:bg-indigo-50 transition-all cursor-pointer"
          />
          <div className="flex flex-wrap gap-3 md:gap-4 mt-4 md:mt-6">
            {formData.lampiran?.map((f, i) => (
              <div
                key={i}
                className="bg-white px-3 py-2 md:px-4 md:py-3 rounded-xl text-[10px] md:text-xs font-medium flex gap-2 md:gap-3 items-center border border-slate-200 shadow-sm animate-in fade-in zoom-in-90 group-hover:bg-slate-50 transition-all"
              >
                <FileText size={16} className="text-indigo-500 w-4 h-4" />
                <span className="truncate max-w-[100px] md:max-w-[150px] text-slate-700">
                  {f.name}
                </span>
                <button
                  type="button"
                  onClick={() => removeFile(i)}
                  className="text-red-400 hover:text-red-600 transition-colors ml-1 md:ml-2"
                >
                  <X size={16} className="w-4 h-4" />
                </button>
              </div>
            ))}
            {(!formData.lampiran || formData.lampiran.length === 0) && (
              <p className="text-slate-400 text-[10px] md:text-xs italic py-1 ml-1">
                Belum ada file yang diunggah.
              </p>
            )}
          </div>
        </div>

        <div className="pt-8 md:pt-12 border-t border-slate-100 flex flex-col-reverse sm:flex-row justify-end gap-3 md:gap-6 mt-8 md:mt-12">
          <button
            type="button"
            onClick={onCancel}
            className="bg-slate-100 text-slate-600 px-6 py-3 md:px-10 md:py-4 rounded-xl font-semibold hover:bg-slate-200 transition-all text-xs md:text-sm w-full sm:w-auto"
          >
            Batal
          </button>
          <button
            type="submit"
            className={`px-8 py-3 md:px-12 md:py-4 rounded-xl font-bold text-xs md:text-sm flex items-center justify-center gap-2 md:gap-3 transition-all shadow-md ${editingCase ? "bg-amber-500 hover:bg-amber-600 shadow-amber-200" : "bg-indigo-600 hover:bg-indigo-700 shadow-indigo-200"} text-white w-full sm:w-auto`}
          >
            {editingCase ? <Edit size={18} className="w-4 h-4 md:w-5 md:h-5" /> : <CheckCircle2 size={18} className="w-4 h-4 md:w-5 md:h-5" />}{" "}
            {editingCase ? "Update Laporan" : "Simpan Laporan"}
          </button>
        </div>
      </form>
    </div>
  );
};

const ReportView: React.FC<{
  dbState: DatabaseState;
  isLoggedIn: boolean;
  onEdit: (r: CaseRecord) => void;
  onDelete: (id: string | number) => void;
  onStatusUpdate: (id: string | number, s: CaseStatus) => void;
}> = ({ dbState, isLoggedIn, onEdit, onDelete, onStatusUpdate }) => {
  const [search, setSearch] = useState("");
  const [filterKategori, setFilterKategori] = useState("");
  const [filterStatus, setFilterStatus] = useState("");

  const filtered = useMemo(() => {
    return dbState.kasus
      .filter(
        (k) =>
          (k.nama_siswa.toLowerCase().includes(search.toLowerCase()) ||
            k.kelas.toLowerCase().includes(search.toLowerCase())) &&
          (filterKategori
            ? filterKategori === "Lainnya"
              ? !KATEGORI_OPTIONS.includes(k.kategori_kasus)
              : k.kategori_kasus === filterKategori
            : true) &&
          (filterStatus ? k.status === filterStatus : true),
      )
      .sort((a, b) => b.created_at - a.created_at);
  }, [dbState.kasus, search, filterKategori, filterStatus]);

  const handleExport = () => {
    if (filtered.length === 0) return alert("Tidak ada data untuk diekspor");
    const ws = XLSX.utils.json_to_sheet(
      filtered.map(({ lampiran, ...rest }) => ({
        ...rest,
        jum_file: lampiran.length,
      })),
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap_Kasus");
    XLSX.writeFile(
      wb,
      `Laporan_BK_Siswa_${new Date().toISOString().split("T")[0]}.xlsx`,
    );
  };

  return (
    <div className="glass-card p-4 md:p-6 lg:p-10 animate-in fade-in duration-500">
      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center mb-6 md:mb-8 gap-4 md:gap-6">
        <div>
          <h3 className="text-xl md:text-2xl font-bold text-slate-900">
            Rekapitulasi Kasus
          </h3>
          <p className="text-xs md:text-sm text-slate-500 mt-1">
            Database lengkap penanganan bimbingan konseling siswa.
          </p>
        </div>
        <div className="flex flex-col sm:flex-row items-center gap-3 w-full lg:w-auto">
          <div className="relative w-full sm:flex-1 lg:min-w-[250px]">
            <Search
              size={18}
              className="absolute left-3 md:left-4 top-1/2 -translate-y-1/2 text-slate-400 w-4 h-4 md:w-5 md:h-5"
            />
            <input
              type="text"
              placeholder="Cari Siswa/Kelas..."
              className="w-full pl-10 md:pl-12 pr-4 py-2 md:py-2.5 rounded-xl bg-slate-50 border border-slate-200 focus:ring-2 focus:ring-indigo-500 outline-none text-xs md:text-sm transition-all"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
          <button
            onClick={handleExport}
            className="bg-emerald-600 text-white px-4 md:px-5 py-2 md:py-2.5 rounded-xl text-xs md:text-sm font-semibold hover:bg-emerald-700 transition flex items-center justify-center gap-2 shadow-sm shadow-emerald-100 w-full sm:w-auto"
          >
            <Download size={16} className="w-4 h-4" /> Excel
          </button>
        </div>
      </div>

      <div className="flex flex-col sm:flex-row flex-wrap gap-2 md:gap-3 mb-4 md:mb-6">
        <select
          className="bg-slate-50 px-3 md:px-4 py-2 rounded-xl text-xs md:text-sm font-medium border border-slate-200 outline-none text-slate-700 focus:ring-2 focus:ring-indigo-500 transition-all w-full sm:w-auto"
          value={filterKategori}
          onChange={(e) => setFilterKategori(e.target.value)}
        >
          <option value="">Semua Kategori</option>
          {[...KATEGORI_OPTIONS, "Lainnya"].map((k) => (
            <option key={k} value={k}>
              {k}
            </option>
          ))}
        </select>
        <select
          className="bg-slate-50 px-3 md:px-4 py-2 rounded-xl text-xs md:text-sm font-medium border border-slate-200 outline-none text-slate-700 focus:ring-2 focus:ring-indigo-500 transition-all w-full sm:w-auto"
          value={filterStatus}
          onChange={(e) => setFilterStatus(e.target.value)}
        >
          <option value="">Semua Status</option>
          {[CaseStatus.BARU, CaseStatus.PROSES, CaseStatus.SELESAI].map((s) => (
            <option key={s} value={s}>
              {s}
            </option>
          ))}
        </select>
      </div>

      <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white">
        <table className="w-full text-left border-collapse min-w-[800px] md:min-w-[1000px]">
          <thead>
            <tr className="text-[10px] md:text-xs uppercase font-semibold text-slate-500 border-b border-slate-200 bg-slate-50">
              <th className="p-3 md:p-4">Tanggal</th>
              <th className="p-3 md:p-4">Informasi Siswa</th>
              <th className="p-3 md:p-4">Kategori</th>
              <th className="p-3 md:p-4">Tindak Lanjut</th>
              <th className="p-3 md:p-4 text-center">Berkas</th>
              <th className="p-3 md:p-4 text-center">Status</th>
              {isLoggedIn && <th className="p-3 md:p-4 text-center">Aksi</th>}
            </tr>
          </thead>
          <tbody className="text-xs md:text-sm divide-y divide-slate-100">
            {filtered.length === 0 ? (
              <tr>
                <td colSpan={isLoggedIn ? 7 : 6} className="p-16 text-center">
                  <div className="flex flex-col items-center opacity-40 grayscale">
                    <Database size={48} className="mb-3 text-slate-400" />
                    <p className="text-slate-500 font-medium text-sm">
                      Belum ada data ditemukan
                    </p>
                  </div>
                </td>
              </tr>
            ) : (
              filtered.map((k) => (
                <tr
                  key={k.id}
                  className="hover:bg-slate-50 transition-colors group"
                >
                  <td className="p-3 md:p-4 whitespace-nowrap text-slate-600">
                    {k.tanggal}
                  </td>
                  <td className="p-3 md:p-4">
                    <div className="font-semibold text-slate-900">
                      {k.nama_siswa}
                    </div>
                    <div className="text-[10px] md:text-xs text-slate-500 mt-0.5">
                      Kelas {k.kelas}
                    </div>
                  </td>
                  <td className="p-3 md:p-4">
                    <span className="bg-indigo-50 text-indigo-700 px-2 md:px-2.5 py-1 rounded-md text-[10px] md:text-xs font-medium">
                      {k.kategori_kasus}
                    </span>
                  </td>
                  <td className="p-3 md:p-4">
                    <div
                      className="max-w-[150px] md:max-w-[200px] truncate font-medium text-slate-700"
                      title={k.tindak_lanjut}
                    >
                      {k.tindak_lanjut}
                    </div>
                    <div className="text-[10px] md:text-xs text-slate-500 mt-0.5">
                      Oleh: {k.guru_bk}
                    </div>
                  </td>
                  <td className="p-3 md:p-4 text-center">
                    {k.lampiran.length > 0 ? (
                      <button
                        onClick={() =>
                          alert(`Total ${k.lampiran.length} file terlampir`)
                        }
                        className="bg-slate-100 text-slate-600 px-2 md:px-2.5 py-1 rounded-md text-[10px] md:text-xs font-medium hover:bg-slate-200 transition-all flex items-center gap-1 md:gap-1.5 mx-auto"
                      >
                        <FileText size={12} className="md:w-[14px] md:h-[14px]" /> {k.lampiran.length}
                      </button>
                    ) : (
                      <span className="text-slate-300">-</span>
                    )}
                  </td>
                  <td className="p-3 md:p-4 text-center">
                    <select
                      value={k.status}
                      disabled={!isLoggedIn}
                      onChange={(e) =>
                        onStatusUpdate(k.id!, e.target.value as CaseStatus)
                      }
                      className={`text-[10px] md:text-xs font-semibold px-2 md:px-2.5 py-1 md:py-1.5 rounded-md border-none outline-none appearance-none text-center transition-colors ${!isLoggedIn ? "cursor-default" : "cursor-pointer"} ${
                        k.status === CaseStatus.BARU
                          ? "bg-blue-50 text-blue-700"
                          : k.status === CaseStatus.PROSES
                            ? "bg-amber-50 text-amber-700"
                            : "bg-emerald-50 text-emerald-700"
                      }`}
                    >
                      <option value={CaseStatus.BARU}>Baru</option>
                      <option value={CaseStatus.PROSES}>Proses</option>
                      <option value={CaseStatus.SELESAI}>Selesai</option>
                    </select>
                  </td>
                  {isLoggedIn && (
                    <td className="p-3 md:p-4 text-center">
                      <div className="flex justify-center gap-1 md:gap-2 opacity-100 md:opacity-0 md:group-hover:opacity-100 transition-opacity">
                        <button
                          onClick={() => onEdit(k)}
                          className="p-1 md:p-1.5 bg-white border border-slate-200 text-slate-600 rounded-lg hover:bg-amber-50 hover:text-amber-600 hover:border-amber-200 transition-colors shadow-sm"
                        >
                          <Edit size={14} className="md:w-4 md:h-4" />
                        </button>
                        <button
                          onClick={() => onDelete(k.id!)}
                          className="p-1 md:p-1.5 bg-white border border-slate-200 text-slate-600 rounded-lg hover:bg-red-50 hover:text-red-600 hover:border-red-200 transition-colors shadow-sm"
                        >
                          <Trash2 size={14} className="md:w-4 md:h-4" />
                        </button>
                      </div>
                    </td>
                  )}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

const RekapView: React.FC<{ dbState: DatabaseState }> = ({ dbState }) => {
  const [filterNama, setFilterNama] = useState("");
  const [filterKategori, setFilterKategori] = useState("");

  const categories = useMemo(() => {
    const cats = new Set<string>();
    dbState.kasus.forEach((k) => cats.add(k.kategori_kasus));
    return Array.from(cats).sort();
  }, [dbState.kasus]);

  const pivotData = useMemo(() => {
    const data: Record<string, Record<string, number>> = {};
    const filteredKasus = dbState.kasus.filter((k) => {
      const matchNama = k.nama_siswa.toLowerCase().includes(filterNama.toLowerCase());
      const matchKategori = filterKategori ? k.kategori_kasus === filterKategori : true;
      return matchNama && matchKategori;
    });

    filteredKasus.forEach((k) => {
      if (!data[k.nama_siswa]) data[k.nama_siswa] = {};
      data[k.nama_siswa][k.kategori_kasus] = (data[k.nama_siswa][k.kategori_kasus] || 0) + 1;
    });

    return data;
  }, [dbState.kasus, filterNama, filterKategori]);

  const students = useMemo(() => Object.keys(pivotData).sort(), [pivotData]);

  const handleExport = () => {
    if (students.length === 0) return alert("Tidak ada data untuk diekspor");
    const exportData = students.map((s) => {
      const row: any = { "Nama Siswa": s };
      categories.forEach((c) => {
        row[c] = pivotData[s][c] || 0;
      });
      row["Total Panggilan"] = (Object.values(pivotData[s]) as number[]).reduce((a: number, b: number) => a + b, 0);
      return row;
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap_Penanganan");
    XLSX.writeFile(wb, `Rekap_Penanganan_Siswa_${new Date().toISOString().split("T")[0]}.xlsx`);
  };

  return (
    <div className="glass-card p-4 md:p-6 lg:p-10 animate-in fade-in duration-500">
      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-center mb-6 md:mb-8 gap-4 md:gap-6">
        <div>
          <h3 className="text-xl md:text-2xl font-bold text-slate-900">
            Rekap Penanganan Siswa
          </h3>
          <p className="text-xs md:text-sm text-slate-500 mt-1">
            Jumlah Panggilan Orang Tua per Siswa dan Kategori Kasus.
          </p>
        </div>
        <button
          onClick={handleExport}
          className="bg-emerald-600 text-white px-4 md:px-5 py-2 md:py-2.5 rounded-xl text-xs md:text-sm font-semibold hover:bg-emerald-700 transition flex items-center justify-center gap-2 shadow-sm shadow-emerald-100 w-full lg:w-auto"
        >
          <Download size={16} className="w-4 h-4" /> Excel
        </button>
      </div>

      <div className="flex flex-col sm:flex-row flex-wrap gap-2 md:gap-3 mb-6">
        <div className="relative w-full sm:w-64">
          <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            type="text"
            placeholder="Filter Nama Siswa..."
            className="w-full pl-10 pr-4 py-2 rounded-xl bg-slate-50 border border-slate-200 focus:ring-2 focus:ring-indigo-500 outline-none text-sm transition-all"
            value={filterNama}
            onChange={(e) => setFilterNama(e.target.value)}
          />
        </div>
        <select
          className="bg-slate-50 px-4 py-2 rounded-xl text-sm font-medium border border-slate-200 outline-none text-slate-700 focus:ring-2 focus:ring-indigo-500 transition-all w-full sm:w-auto"
          value={filterKategori}
          onChange={(e) => setFilterKategori(e.target.value)}
        >
          <option value="">Semua Kategori</option>
          {categories.map((c) => (
            <option key={c} value={c}>
              {c}
            </option>
          ))}
        </select>
      </div>

      <div className="overflow-x-auto rounded-2xl border border-slate-200 bg-white">
        <table className="w-full text-left border-collapse min-w-[600px]">
          <thead>
            <tr className="text-[10px] md:text-xs uppercase font-semibold text-slate-500 border-b border-slate-200 bg-slate-50">
              <th className="p-4 sticky left-0 bg-slate-50 z-10">Nama Siswa</th>
              {categories.map((c) => (
                <th key={c} className="p-4 text-center">{c}</th>
              ))}
              <th className="p-4 text-center bg-indigo-50 text-indigo-700">Total</th>
            </tr>
          </thead>
          <tbody className="text-xs md:text-sm divide-y divide-slate-100">
            {students.length === 0 ? (
              <tr>
                <td colSpan={categories.length + 2} className="p-16 text-center">
                  <div className="flex flex-col items-center opacity-40 grayscale">
                    <Activity size={48} className="mb-3 text-slate-400" />
                    <p className="text-slate-500 font-medium text-sm">
                      Tidak ada data panggilan orang tua ditemukan
                    </p>
                  </div>
                </td>
              </tr>
            ) : (
              students.map((s) => {
                const total = (Object.values(pivotData[s]) as number[]).reduce((a: number, b: number) => a + b, 0);
                return (
                  <tr key={s} className="hover:bg-slate-50 transition-colors">
                    <td className="p-4 font-semibold text-slate-900 sticky left-0 bg-white group-hover:bg-slate-50 z-10 border-r border-slate-100">
                      {s}
                    </td>
                    {categories.map((c) => (
                      <td key={c} className="p-4 text-center text-slate-600">
                        {pivotData[s][c] || 0}
                      </td>
                    ))}
                    <td className="p-4 text-center font-bold text-indigo-600 bg-indigo-50/30">
                      {total}
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default App;
