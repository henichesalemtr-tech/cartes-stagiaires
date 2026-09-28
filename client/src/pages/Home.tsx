import {
  useEffect,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
} from "react";
import * as XLSX from "xlsx";
import QRCode from "qrcode";
import {
  Archive,
  ArrowDownToLine,
  ArrowUpFromLine,
  BadgeCheck,
  BarChart3,
  BookOpen,
  Check,
  ChevronLeft,
  ChevronRight,
  ClipboardList,
  Download,
  FileSpreadsheet,
  FileText,
  ImagePlus,
  Languages,
  LayoutDashboard,
  Menu,
  Pencil,
  Plus,
  Printer,
  QrCode,
  RotateCcw,
  Save,
  Search,
  Settings2,
  ShieldCheck,
  Trash2,
  Upload,
  UserRound,
  UsersRound,
  X,
} from "lucide-react";
import { toast } from "sonner";

type View =
  | "dashboard"
  | "trainees"
  | "import"
  | "design"
  | "print"
  | "settings";
type Lang = "ar" | "fr";
type ScaleKey =
  | "orgScale"
  | "line2Scale"
  | "line3Scale"
  | "nameScale"
  | "registrationScale"
  | "passportScale"
  | "specialtyScale"
  | "footerScale"
  | "studyYearScale"
  | "flagsScale"
  | "ministryLogoScale";

type Trainee = {
  id: string;
  firstName: string;
  lastName: string;
  registration: string;
  passport: string;
  studyYear: string;
  specialty: string;
  level: string;
  photo?: string;
  fullName?: string;
};

type Settings = {
  orgAr: string;
  orgFr: string;
  line2Ar: string;
  line2Fr: string;
  line3Ar: string;
  line3Fr: string;
  ministryAr: string;
  ministryFr: string;
  cardLanguage: Lang;
  showFlags: boolean;
  showMinistry: boolean;
  showPassport: boolean;
  showSpecialty: boolean;
  showPhoto: boolean;
  showStudyYear: boolean;
  textScale: number;
  logoScale: number;
  orgScale: number;
  line2Scale: number;
  line3Scale: number;
  nameScale: number;
  registrationScale: number;
  passportScale: number;
  specialtyScale: number;
  footerScale: number;
  studyYearScale: number;
  flagsScale: number;
  ministryLogoScale: number;
  accent: "green" | "blue" | "amber";
};

const STORAGE_KEY = "hakani-cards-state-v1";
const OFFICIAL_LOGO = "/logo-mfep.png";
const ALGERIA_FLAG = "/flag-algeria.svg";
const NIGER_FLAG = "/flag-niger.svg";

const defaultSettings: Settings = {
  orgAr: "المعهد الوطني المتخصص في التكوين المهني",
  orgFr: "REPUBLIQUE ALGERIENNE DEMOCRATIQUE ET POPULAIRE",
  line2Ar: "Hassani Abdelkrim",
  line2Fr: "Ministere de la formation et de l'enseingnement professionnels",
  line3Ar: "Adrar • Algérie",
  line3Fr:
    "Institut National Spécialisé de Formation Professionnelle Hassani Abdelkrim",
  ministryAr: "",
  ministryFr: "",
  cardLanguage: "fr",
  showFlags: true,
  showMinistry: true,
  showPassport: true,
  showSpecialty: true,
  showPhoto: true,
  showStudyYear: true,
  textScale: 100,
  logoScale: 100,
  orgScale: 100,
  line2Scale: 100,
  line3Scale: 100,
  nameScale: 100,
  registrationScale: 100,
  passportScale: 100,
  specialtyScale: 100,
  footerScale: 100,
  studyYearScale: 100,
  flagsScale: 100,
  ministryLogoScale: 100,
  accent: "green",
};

const demoTrainees: Trainee[] = [
  {
    id: "tr-01",
    firstName: "Mohamed",
    lastName: "ALI",
    registration: "2025/00147",
    passport: "AA384921",
    studyYear: "2025 — 2026",
    specialty: "Informatique",
    level: "Technicien supérieur",
  },
  {
    id: "tr-02",
    firstName: "Meriem",
    lastName: "NOUR",
    registration: "2025/00148",
    passport: "AA384922",
    studyYear: "2025 — 2026",
    specialty: "Comptabilité et gestion",
    level: "Technicien supérieur",
  },
  {
    id: "tr-03",
    firstName: "Salma",
    lastName: "TEST",
    registration: "2025/00149",
    passport: "AA384923",
    studyYear: "2025 — 2026",
    specialty: "Électricité industrielle",
    level: "Technicien",
  },
  {
    id: "tr-04",
    firstName: "Amir",
    lastName: "DEEP",
    registration: "2025/00150",
    passport: "AA384924",
    studyYear: "2025 — 2026",
    specialty: "Marketing digital",
    level: "Technicien supérieur",
  },
  {
    id: "tr-05",
    firstName: "Salem",
    lastName: "HADAD",
    registration: "2025/00151",
    passport: "AA384925",
    studyYear: "2025 — 2026",
    specialty: "Énergies renouvelables",
    level: "Technicien",
  },
];

function traineeName(trainee: Trainee) {
  return (
    trainee.fullName?.trim() ||
    [trainee.firstName, trainee.lastName].filter(Boolean).join(" ") ||
    "—"
  );
}
function uid() {
  return `tr-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { trainees: demoTrainees, settings: defaultSettings };
    const parsed = JSON.parse(raw);
    const migratedSettings =
      parsed.version === 2
        ? parsed.settings || {}
        : { ...(parsed.settings || {}), cardLanguage: "fr" };
    return {
      trainees: Array.isArray(parsed.trainees) ? parsed.trainees : demoTrainees,
      settings: { ...defaultSettings, ...migratedSettings },
    };
  } catch {
    return { trainees: demoTrainees, settings: defaultSettings };
  }
}

function saveState(trainees: Trainee[], settings: Settings) {
  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({ version: 2, trainees, settings })
  );
}

function downloadBlob(
  content: string,
  filename: string,
  type = "application/json"
) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function Flags() {
  return (
    <span className="flags" aria-label="Drapeaux Algérie et Niger">
      <img className="flag-image" src={ALGERIA_FLAG} alt="Algérie" />
      <img className="flag-image" src={NIGER_FLAG} alt="Niger" />
    </span>
  );
}

function MinistryMark({ small = false }: { small?: boolean }) {
  return (
    <span
      className={`ministry-mark ${small ? "ministry-mark-small" : ""}`}
      aria-label="شعار وزارة التكوين والتعليم المهنيين"
    >
      <img
        src={OFFICIAL_LOGO}
        alt="Ministère de la Formation et de l'Enseignement Professionnels"
      />
    </span>
  );
}

function QRImage({
  trainee,
  className = "",
}: {
  trainee: Trainee;
  className?: string;
}) {
  const [src, setSrc] = useState("");
  useEffect(() => {
    let alive = true;
    QRCode.toDataURL(
      JSON.stringify({
        registration: trainee.registration,
        fullName: traineeName(trainee),
        passport: trainee.passport,
      }),
      {
        width: 180,
        margin: 0,
        color: { dark: "#112a20", light: "#ffffff" },
        errorCorrectionLevel: "M",
      }
    ).then(url => alive && setSrc(url));
    return () => {
      alive = false;
    };
  }, [trainee]);
  return src ? (
    <img className={className} src={src} alt="QR" />
  ) : (
    <span className={`qr-placeholder ${className}`}>
      <QrCode size={22} />
    </span>
  );
}

function CardPreview({
  trainee,
  settings,
  compact = false,
  language,
  onEdit,
}: {
  trainee: Trainee;
  settings: Settings;
  compact?: boolean;
  language?: Lang;
  onEdit?: () => void;
}) {
  const lang = language || settings.cardLanguage;
  const isFr = lang === "fr";
  return (
    <div
      className={`id-card ${compact ? "id-card-compact" : ""} accent-${settings.accent}`}
      dir={isFr ? "ltr" : "rtl"}
      style={
        {
          "--card-text-scale": String(settings.textScale / 100),
          "--card-logo-scale": String(settings.logoScale / 100),
          "--card-org-scale": String(settings.orgScale / 100),
          "--card-line2-scale": String(settings.line2Scale / 100),
          "--card-line3-scale": String(settings.line3Scale / 100),
          "--card-name-scale": String(settings.nameScale / 100),
          "--card-registration-scale": String(settings.registrationScale / 100),
          "--card-passport-scale": String(settings.passportScale / 100),
          "--card-specialty-scale": String(settings.specialtyScale / 100),
          "--card-footer-scale": String(settings.footerScale / 100),
          "--card-study-year-scale": String(settings.studyYearScale / 100),
          "--card-flags-scale": String(settings.flagsScale / 100),
          "--card-ministry-logo-scale": String(settings.ministryLogoScale / 100),
        } as CSSProperties
      }
    >
      <div className="card-wave card-wave-one" />
      <div className="card-wave card-wave-two" />
      <div className="zellige zellige-tl" />
      <div className="zellige zellige-br" />
      <div className="card-crescent">☪</div>
      <div className="card-header">
        {settings.showFlags && <Flags />}
        <div className="card-titles">
          <span className="card-org-text">{isFr ? settings.orgFr : settings.orgAr}</span>
          <strong className="card-line2-text">{isFr ? settings.line2Fr : settings.line2Ar}</strong>
          <small className="card-line3-text">{isFr ? settings.line3Fr : settings.line3Ar}</small>
        </div>
        {settings.showMinistry && <MinistryMark small />}
      </div>
      <div className="card-rule" />
      <div className="card-body">
        {settings.showPhoto && (
          <div className="photo-box">
            {trainee.photo ? (
              <img src={trainee.photo} alt="صورة Stagiaire" />
            ) : (
              <UserRound size={compact ? 21 : 28} strokeWidth={1.3} />
            )}
            <span>PHOTO</span>
          </div>
        )}
        <div className="trainee-info">
          <div className="trainee-name card-matricule-text">
            <span>Matricule</span>
            <b>{trainee.registration || "—"}</b>
          </div>
          <div className="trainee-name card-name-text">
            <span>{isFr ? "NOM ET PRÉNOM" : "الاسم واللقب"}</span>
            <b>{traineeName(trainee)}</b>
          </div>
          {settings.showPassport && (
            <div className="info-row">
              <span>Passeport</span>
              <b className="card-passport-value">{trainee.passport || "—"}</b>
            </div>
          )}
          {settings.showSpecialty && (
            <div className="info-row">
              <span>Spécialité</span>
              <b className="card-specialty-value">{trainee.specialty || "—"}</b>
            </div>
          )}
        </div>
        <div className="qr-box">
          <QRImage trainee={trainee} />
          <span>{isFr ? "SCAN" : "مسح"}</span>
        </div>
      </div>
      <div className="card-footer">
        <span className="card-footer-text">{isFr ? settings.ministryFr : settings.ministryAr}</span>
        {settings.showStudyYear && (
          <span className="valid-chip card-study-year-text">
            {trainee.studyYear || "2025 — 2026"}
          </span>
        )}
      </div>
      {onEdit && (
        <button
          className="card-edit"
          onClick={onEdit}
          aria-label="تعديل Stagiaire"
        >
          <Pencil size={13} />
        </button>
      )}
    </div>
  );
}

function Sidebar({
  activeView,
  setActiveView,
  traineesCount,
  onClose,
}: {
  activeView: View;
  setActiveView: (view: View) => void;
  traineesCount: number;
  onClose?: () => void;
}) {
  const items: {
    id: View;
    label: string;
    icon: typeof LayoutDashboard;
    badge?: string;
  }[] = [
    { id: "dashboard", label: "Tableau de bord", icon: LayoutDashboard },
    {
      id: "trainees",
      label: "Stagiaires",
      icon: UsersRound,
      badge: String(traineesCount),
    },
    { id: "import", label: "Import des données", icon: FileSpreadsheet },
    { id: "design", label: "Conception de la carte", icon: BadgeCheck },
    { id: "print", label: "Impression & export", icon: Printer },
  ];
  return (
    <aside className="sidebar">
      <div className="brand-lockup">
        <div className="brand-icon">
          <MinistryMark />
        </div>
        <div>
          <strong>CARTE PRO</strong>
          <span>Gestion des cartes stagiaires</span>
        </div>
      </div>
      <div className="sidebar-label">ESPACE DE TRAVAIL</div>
      <nav>
        {items.map(({ id, label, icon: Icon, badge }) => (
          <button
            key={id}
            className={`nav-item ${activeView === id ? "active" : ""}`}
            onClick={() => {
              setActiveView(id);
              onClose?.();
            }}
          >
            <Icon size={18} />
            <span>{label}</span>
            {badge && <em>{badge}</em>}
          </button>
        ))}
      </nav>
      <div className="sidebar-bottom">
        <div className="storage-card">
          <div className="storage-icon">
            <Archive size={16} />
          </div>
          <div>
            <strong>Stockage local</strong>
            <span>Données sur cet appareil</span>
          </div>
          <Check size={15} />
        </div>
        <button
          className={`nav-item ${activeView === "settings" ? "active" : ""}`}
          onClick={() => {
            setActiveView("settings");
            onClose?.();
          }}
        >
          <Settings2 size={18} />
          <span>Paramètres</span>
        </button>
        <div className="user-mini">
          <div className="avatar">م</div>
          <div>
            <strong>Administrateur de l’institut</strong>
            <span>Compte local</span>
          </div>
          <ChevronLeft size={16} />
        </div>
      </div>
    </aside>
  );
}

function Topbar({
  activeView,
  lang,
  setLang,
  onMenu,
}: {
  activeView: View;
  lang: Lang;
  setLang: (lang: Lang) => void;
  onMenu: () => void;
}) {
  const titles: Record<View, [string, string]> = {
    dashboard: ["Tableau de bord", "Résumé de l’activité des cartes"],
    trainees: ["Stagiaires", "إدارة السجل وإضافة الصور"],
    import: ["Import des données", "إضافة دفعة sur ملف Excel"],
    design: ["Conception de la carte", "خصّص هوية البطاقة ومحتواها"],
    print: ["Impression & export", "جهّز بطاقاتك للطباعة على A4"],
    settings: ["Paramètres", "النسخ الاحتياطي والتهيئة المحلية"],
  };
  const [ar, sub] = titles[activeView];
  return (
    <header className="topbar">
      <button className="mobile-menu" onClick={onMenu}>
        <Menu size={20} />
      </button>
      <div>
        <h1>{ar}</h1>
        <p>{sub}</p>
      </div>
      <div className="topbar-actions">
        <div className="online-pill">
          <span /> Fonctionne en local
        </div>
        <div className="lang-toggle">
          <button
            className={lang === "ar" ? "selected" : ""}
            onClick={() => setLang("ar")}
          >
            ع
          </button>
          <button
            className={lang === "fr" ? "selected" : ""}
            onClick={() => setLang("fr")}
          >
            FR
          </button>
          <Languages size={15} />
        </div>
      </div>
    </header>
  );
}

function Dashboard({
  trainees,
  settings,
  setView,
  onEdit,
}: {
  trainees: Trainee[];
  settings: Settings;
  setView: (v: View) => void;
  onEdit: (t: Trainee) => void;
}) {
  const withPhoto = trainees.filter(t => t.photo).length;
  const stats = [
    {
      label: "Total des stagiaires",
      value: trainees.length,
      note: "Dans le registre local",
      icon: UsersRound,
      tone: "green",
    },
    {
      label: "Cartes prêtes",
      value: withPhoto,
      note: `${Math.round((withPhoto / Math.max(trainees.length, 1)) * 100)}% du registre`,
      icon: BadgeCheck,
      tone: "amber",
    },
    {
      label: "Photos manquantes",
      value: trainees.length - withPhoto,
      note: "Ajoutez-les depuis le registre",
      icon: ImagePlus,
      tone: "blue",
    },
  ];
  const recent = trainees.slice(0, 4);
  return (
    <div className="page-content">
      <section className="hero-panel">
        <div className="hero-copy">
          <div className="eyebrow">
            <ShieldCheck size={14} /> Centre d’émission des cartes
          </div>
          <h2>
            Chaque carte commence par <em>des données précises.</em>
          </h2>
          <p>
            Créez des cartes uniformes, importez vos données en masse et
            imprimez-les en quelques minutes.
          </p>
          <div className="hero-actions">
            <button
              className="button button-primary"
              onClick={() => setView("trainees")}
            >
              <Plus size={17} /> Ajouter un stagiaire
            </button>
            <button
              className="button button-ghost"
              onClick={() => setView("import")}
            >
              <Upload size={16} /> Importer Excel
            </button>
          </div>
        </div>
        <div className="hero-visual">
          <div className="floating-card-main">
            <CardPreview
              trainee={trainees[0] || demoTrainees[0]}
              settings={settings}
              compact
            />
          </div>
          <div className="hero-stamp">
            <span>CR80</span>
            <small>STANDARD</small>
          </div>
        </div>
      </section>
      <div className="stats-grid">
        {stats.map(({ label, value, note, icon: Icon, tone }) => (
          <div className="stat-card" key={label}>
            <div className={`stat-icon ${tone}`}>
              <Icon size={19} />
            </div>
            <div>
              <span>{label}</span>
              <strong>{value}</strong>
              <small>{note}</small>
            </div>
            <BarChart3 className="stat-chart" size={31} />
          </div>
        ))}
      </div>
      <div className="dashboard-grid">
        <section className="surface recent-card">
          <div className="section-heading">
            <div>
              <span className="section-kicker">Ajouts récents</span>
              <h3>Stagiaires ajoutés récemment</h3>
            </div>
            <button className="text-button" onClick={() => setView("trainees")}>
              Voir tout <ChevronLeft size={15} />
            </button>
          </div>
          <div className="recent-list">
            {recent.map((t, index) => (
              <button
                className="recent-row"
                key={t.id}
                onClick={() => onEdit(t)}
              >
                <div className="row-avatar">
                  {t.photo ? (
                    <img src={t.photo} alt="" />
                  ) : (
                    traineeName(t).charAt(0)
                  )}
                </div>
                <div className="row-name">
                  <strong>{traineeName(t)}</strong>
                  <span>{t.specialty}</span>
                </div>
                <span className="row-code">{t.registration}</span>
                <span className={`status-dot ${t.photo ? "ready" : "pending"}`}>
                  {t.photo ? "Prête" : "Photo manquante"}
                </span>
                <ChevronLeft size={15} className="row-chevron" />
              </button>
            ))}
          </div>
        </section>
        <section className="surface quick-card">
          <div className="section-heading">
            <div>
              <span className="section-kicker">Raccourcis</span>
              <h3>Commencer ici</h3>
            </div>
          </div>
          <button className="quick-action" onClick={() => setView("import")}>
            <div className="quick-icon import">
              <FileSpreadsheet size={19} />
            </div>
            <div>
              <strong>Importer un fichier Excel</strong>
              <span>Ajoutez plusieurs stagiaires en une fois</span>
            </div>
            <ChevronLeft size={16} />
          </button>
          <button className="quick-action" onClick={() => setView("design")}>
            <div className="quick-icon design">
              <BadgeCheck size={19} />
            </div>
            <div>
              <strong>Personnaliser la carte</strong>
              <span>Titres, langue et couleurs</span>
            </div>
            <ChevronLeft size={16} />
          </button>
          <button className="quick-action" onClick={() => setView("print")}>
            <div className="quick-icon print">
              <Printer size={19} />
            </div>
            <div>
              <strong>Préparer l’impression</strong>
              <span>9 cartes sur une feuille A4</span>
            </div>
            <ChevronLeft size={16} />
          </button>
        </section>
      </div>
    </div>
  );
}

function TraineesView({
  trainees,
  setTrainees,
  settings,
  openEditor,
}: {
  trainees: Trainee[];
  setTrainees: (t: Trainee[]) => void;
  settings: Settings;
  openEditor: (t?: Trainee) => void;
}) {
  const [query, setQuery] = useState("");
  const [selected, setSelected] = useState<string[]>([]);
  const filtered = useMemo(
    () =>
      trainees.filter(t =>
        `${traineeName(t)} ${t.registration} ${t.passport} ${t.specialty}`
          .toLowerCase()
          .includes(query.toLowerCase())
      ),
    [trainees, query]
  );
  const remove = (id: string) => {
    if (confirm("حذف هذا Stagiaire sur Registre local؟")) {
      setTrainees(trainees.filter(t => t.id !== id));
      toast.success("تم حذف Stagiaire");
    }
  };
  const removeSelected = () => {
    if (!selected.length) return;
    if (confirm(`حذف ${selected.length} متربصين من السجل المحلي؟`)) {
      setTrainees(trainees.filter(t => !selected.includes(t.id)));
      setSelected([]);
      toast.success(`تم حذف ${selected.length} سجلاً`);
    }
  };
  const toggle = (id: string) =>
    setSelected(
      selected.includes(id) ? selected.filter(x => x !== id) : [...selected, id]
    );
  const allVisibleSelected =
    filtered.length > 0 && filtered.every(t => selected.includes(t.id));
  return (
    <div className="page-content">
      <div className="view-toolbar">
        <div>
          <span className="section-kicker">Registre local</span>
          <h2 className="page-title">
            Liste des stagiaires <span>{trainees.length}</span>
          </h2>
        </div>
        <div className="toolbar-actions">
          <div className="search-box">
            <Search size={17} />
            <input
              value={query}
              onChange={e => setQuery(e.target.value)}
              placeholder="ابحث بالاسم أو Matricule..."
            />
          </div>
          <button
            className="button button-primary"
            onClick={() => openEditor()}
          >
            <Plus size={17} /> Nouveau stagiaire
          </button>
        </div>
      </div>
      <section className="surface table-surface">
        <div className="table-top">
          <div className="table-meta">
            {selected.length ? (
              <span className="selected-count">تم تحديد {selected.length}</span>
            ) : (
              <span>Les enregistrements ajoutés apparaissent ici</span>
            )}
            {selected.length > 0 && (
              <>
                <button className="link-button" onClick={() => setSelected([])}>
                  Annuler التحديد
                </button>
                <button
                  className="button button-danger"
                  onClick={removeSelected}
                >
                  <Trash2 size={14} /> حذف المحدد
                </button>
              </>
            )}
          </div>
          <div className="table-tools">
            <span className="filter-chip">
              <span className="filter-dot" /> Tous <b>{filtered.length}</b>
            </span>
          </div>
        </div>
        <div className="trainees-table">
          <div className="table-row table-head">
            <div>
              <input
                type="checkbox"
                checked={allVisibleSelected}
                onChange={() =>
                  setSelected(
                    allVisibleSelected
                      ? selected.filter(id => !filtered.some(t => t.id === id))
                      : Array.from(
                          new Set([...selected, ...filtered.map(t => t.id)])
                        )
                  )
                }
              />
            </div>
            <span>Stagiaire</span>
            <span>Matricule</span>
            <span>Passeport</span>
            <span>Spécialité / niveau</span>
            <span>État de la carte</span>
            <span></span>
          </div>
          {filtered.map(t => (
            <div className="table-row" key={t.id}>
              <div>
                <input
                  type="checkbox"
                  checked={selected.includes(t.id)}
                  onChange={() => toggle(t.id)}
                />
              </div>
              <div className="trainee-cell">
                <div className="table-avatar">
                  {t.photo ? (
                    <img src={t.photo} alt="" />
                  ) : (
                    traineeName(t).charAt(0)
                  )}
                </div>
                <div>
                  <strong>{traineeName(t)}</strong>
                  <small>Ajouté localement</small>
                </div>
              </div>
              <span className="mono-text">{t.registration}</span>
              <span className="mono-text muted">{t.passport || "—"}</span>
              <div className="specialty-cell">
                <strong>{t.specialty || "—"}</strong>
                <small>{t.level || "—"}</small>
              </div>
              <span className={`table-status ${t.photo ? "ready" : "pending"}`}>
                <i /> {t.photo ? "Prête à imprimer" : "Photo requise"}
              </span>
              <div className="row-actions">
                <button onClick={() => openEditor(t)} aria-label="تعديل">
                  <Pencil size={15} />
                </button>
                <button onClick={() => remove(t.id)} aria-label="حذف">
                  <Trash2 size={15} />
                </button>
              </div>
            </div>
          ))}
          {filtered.length === 0 && (
            <div className="empty-state">
              <UsersRound size={32} />
              <strong>Aucun résultat</strong>
              <span>Essayez une autre recherche ou ajoutez un stagiaire</span>
            </div>
          )}
        </div>
      </section>
    </div>
  );
}

function ImportView({
  onImport,
  onExport,
  trainees,
}: {
  onImport: (file: File) => void;
  onExport: () => void;
  trainees: Trainee[];
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const jsonRef = useRef<HTMLInputElement>(null);
  const downloadTemplate = () => {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      [
        "Nom et prénom",
        "Matricule",
        "رقم Passeport",
        "Année scolaire",
        "Spécialité",
        "Niveau",
      ],
      [
        "Ali MOHAMMED",
        "2025/00147",
        "AA384921",
        "2025 — 2026",
        "Informatique",
        "Technicien supérieur",
      ],
    ]);
    XLSX.utils.book_append_sheet(workbook, sheet, "Stagiaires");
    const content = XLSX.write(workbook, { bookType: "xls", type: "array" });
    downloadBlob(content, "قالب-Stagiaires.xls", "application/vnd.ms-excel");
  };
  return (
    <div className="page-content">
      <div className="view-toolbar">
        <div>
          <span className="section-kicker">Import sécurisé</span>
          <h2 className="page-title">Ajoutez vos données en une fois</h2>
          <p className="page-lead">
            Les fichiers Excel XLS ou XLSX sont pris en charge. Vos données
            restent sur cet appareil.
          </p>
        </div>
        <button className="button button-ghost" onClick={onExport}>
          <ArrowDownToLine size={16} /> Exporter une sauvegarde
        </button>
      </div>
      <div className="import-layout">
        <section
          className="surface import-drop"
          onClick={() => inputRef.current?.click()}
        >
          <input
            ref={inputRef}
            type="file"
            hidden
            accept=".xlsx,.xls"
            onChange={e => e.target.files?.[0] && onImport(e.target.files[0])}
          />
          <div className="upload-orb">
            <Upload size={25} />
          </div>
          <h3>Déposez votre fichier Excel ici</h3>
          <p>ou cliquez pour choisir un fichier</p>
          <span className="file-types">
            <FileSpreadsheet size={15} /> XLS / XLSX
          </span>
          <button className="button button-primary">
            Choisir un fichier <ArrowUpFromLine size={16} />
          </button>
        </section>
        <section className="surface import-guide">
          <div className="guide-head">
            <div className="guide-icon">
              <ClipboardList size={19} />
            </div>
            <div>
              <h3>Format des colonnes</h3>
              <p>Utilisez ces en-têtes pour faire correspondre les données</p>
            </div>
          </div>
          <div className="columns-list">
            <span>
              Nom et prénom <b>Requis</b>
            </span>
            <span>
              Matricule <b>Requis</b>
            </span>
            <span>
              رقم Passeport <i>Facultatif</i>
            </span>
            <span>
              Spécialité <i>Facultatif</i>
            </span>
            <span>
              Niveau <i>Facultatif</i>
            </span>
          </div>
          <button className="text-button wide" onClick={downloadTemplate}>
            <Download size={15} /> Télécharger le modèle
          </button>
        </section>
      </div>
      <section className="surface backup-strip">
        <div className="backup-icon">
          <Archive size={18} />
        </div>
        <div>
          <strong>Sauvegarde locale</strong>
          <span>
            صدّر {trainees.length} سجلاً إلى ملف JSON، واستعده لاحقاً على نفس
            الجهاز أو جهاز آخر.
          </span>
        </div>
        <div className="backup-actions">
          <button className="button button-ghost" onClick={onExport}>
            <Download size={15} /> تصدير JSON
          </button>
          <input
            ref={jsonRef}
            type="file"
            hidden
            accept=".json"
            onChange={e => e.target.files?.[0] && onImport(e.target.files[0])}
          />
          <button
            className="button button-ghost"
            onClick={() => jsonRef.current?.click()}
          >
            <RotateCcw size={15} /> Restaurer JSON
          </button>
        </div>
      </section>
    </div>
  );
}

function ScaleSlider({
  label,
  value,
  min = 70,
  max = 140,
  onChange,
}: {
  label: string;
  value: number;
  min?: number;
  max?: number;
  onChange: (value: number) => void;
}) {
  return (
    <label className="scale-control">
      <span>
        <b>{label}</b>
        <output>{value}%</output>
      </span>
      <input
        type="range"
        min={min}
        max={max}
        step="5"
        value={value}
        onChange={e => onChange(Number(e.target.value))}
      />
    </label>
  );
}

function DesignView({
  settings,
  setSettings,
  trainee,
}: {
  settings: Settings;
  setSettings: (s: Settings) => void;
  trainee: Trainee;
}) {
  const update = (key: keyof Settings, value: string | boolean | number) =>
    setSettings({ ...settings, [key]: value });
  return (
    <div className="page-content">
      <div className="view-toolbar">
        <div>
          <span className="section-kicker">Aperçu en direct</span>
          <h2 className="page-title">Concevez la carte de l’institut</h2>
          <p className="page-lead">
            Chaque modification apparaît immédiatement sur le modèle CR80.
          </p>
        </div>
        <div className="design-actions">
          <button
            className={`seg-button ${settings.cardLanguage === "ar" ? "active" : ""}`}
            onClick={() => update("cardLanguage", "ar")}
          >
            Arabe
          </button>
          <button
            className={`seg-button ${settings.cardLanguage === "fr" ? "active" : ""}`}
            onClick={() => update("cardLanguage", "fr")}
          >
            Français
          </button>
        </div>
      </div>
      <div className="design-layout">
        <section className="surface design-form">
          <div className="form-section">
            <div className="form-section-title">
              <span>01</span>
              <div>
                <h3>عناوين البطاقة</h3>
                <p>النصوص التي تظهر أعلى surتصف البطاقة</p>
              </div>
            </div>
            <div className="field-grid">
              <label>
                <span>العنوان الرئيسي بArabe</span>
                <input
                  value={settings.orgAr}
                  onChange={e => update("orgAr", e.target.value)}
                />
              </label>
              <label>
                <span>Titre principal en français</span>
                <input
                  dir="ltr"
                  value={settings.orgFr}
                  onChange={e => update("orgFr", e.target.value)}
                />
              </label>
              <label>
                <span>اسم المعهد بArabe</span>
                <input
                  value={settings.line2Ar}
                  onChange={e => update("line2Ar", e.target.value)}
                />
              </label>
              <label>
                <span>Nom de l'institut</span>
                <input
                  dir="ltr"
                  value={settings.line2Fr}
                  onChange={e => update("line2Fr", e.target.value)}
                />
              </label>
              <label>
                <span>السطر الثالث بArabe</span>
                <input
                  value={settings.line3Ar}
                  onChange={e => update("line3Ar", e.target.value)}
                />
              </label>
              <label>
                <span>Troisième ligne</span>
                <input
                  dir="ltr"
                  value={settings.line3Fr}
                  onChange={e => update("line3Fr", e.target.value)}
                />
              </label>
            </div>
          </div>
          <div className="form-section">
            <div className="form-section-title">
              <span>02</span>
              <div>
                <h3>Éléments visibles</h3>
                <p>Contrôlez les logos et l’identité visuelle</p>
              </div>
            </div>
            <div className="toggle-list">
              <label className="toggle-row">
                <span>
                  <Flags />
                  <b>Drapeaux Algérie & Niger</b>
                </span>
                <input
                  type="checkbox"
                  checked={settings.showFlags}
                  onChange={e => update("showFlags", e.target.checked)}
                />
                <i />
              </label>
              <label className="toggle-row">
                <span>
                  <MinistryMark small />
                  <b>Logo officiel du ministère</b>
                </span>
                <input
                  type="checkbox"
                  checked={settings.showMinistry}
                  onChange={e => update("showMinistry", e.target.checked)}
                />
                <i />
              </label>
              <label className="toggle-row">
                <span>
                  <b>Numéro de passeport</b>
                </span>
                <input
                  type="checkbox"
                  checked={settings.showPassport}
                  onChange={e => update("showPassport", e.target.checked)}
                />
                <i />
              </label>
              <label className="toggle-row">
                <span>
                  <b>Spécialité</b>
                </span>
                <input
                  type="checkbox"
                  checked={settings.showSpecialty}
                  onChange={e => update("showSpecialty", e.target.checked)}
                />
                <i />
              </label>
              <label className="toggle-row">
                <span>
                  <b>Photo personnelle</b>
                </span>
                <input
                  type="checkbox"
                  checked={settings.showPhoto}
                  onChange={e => update("showPhoto", e.target.checked)}
                />
                <i />
              </label>
              <label className="toggle-row">
                <span>
                  <b>Année scolaire</b>
                </span>
                <input
                  type="checkbox"
                  checked={settings.showStudyYear}
                  onChange={e => update("showStudyYear", e.target.checked)}
                />
                <i />
              </label>
            </div>
            <div className="scale-controls">
              <ScaleSlider label="السطر الأول / Organisation" value={settings.orgScale} onChange={v => update("orgScale", v)} />
              <ScaleSlider label="السطر الثاني / Institut" value={settings.line2Scale} onChange={v => update("line2Scale", v)} />
              <ScaleSlider label="السطر الثالث / Adresse" value={settings.line3Scale} onChange={v => update("line3Scale", v)} />
              <ScaleSlider label="الاسم واللقب" value={settings.nameScale} onChange={v => update("nameScale", v)} />
              <ScaleSlider label="Matricule" value={settings.registrationScale} onChange={v => update("registrationScale", v)} />
              <ScaleSlider label="Passeport" value={settings.passportScale} onChange={v => update("passportScale", v)} />
              <ScaleSlider label="Spécialité" value={settings.specialtyScale} onChange={v => update("specialtyScale", v)} />
              <ScaleSlider label="النص السفلي" value={settings.footerScale} onChange={v => update("footerScale", v)} />
              <ScaleSlider label="Année scolaire" value={settings.studyYearScale} onChange={v => update("studyYearScale", v)} />
              <ScaleSlider label="العلمان" value={settings.flagsScale} onChange={v => update("flagsScale", v)} min={70} />
              <ScaleSlider label="شعار الوزارة" value={settings.ministryLogoScale} onChange={v => update("ministryLogoScale", v)} min={70} />
            </div>
            <div className="color-picker">
              <span>Couleur principale</span>
              <div>
                <button
                  className={settings.accent === "green" ? "selected" : ""}
                  onClick={() => update("accent", "green")}
                >
                  <i className="swatch green" />
                  Vert algérien
                </button>
                <button
                  className={settings.accent === "blue" ? "selected" : ""}
                  onClick={() => update("accent", "blue")}
                >
                  <i className="swatch blue" />
                  Bleu institutionnel
                </button>
                <button
                  className={settings.accent === "amber" ? "selected" : ""}
                  onClick={() => update("accent", "amber")}
                >
                  <i className="swatch amber" />
                  Ocre chaleureux
                </button>
              </div>
            </div>
          </div>
        </section>
        <section className="preview-panel">
          <div className="preview-label">
            <span>
              <span className="live-dot" /> معاينة مباشرة
            </span>
            <small>85.6 × 53.98 mm · CR80</small>
          </div>
          <div className="preview-frame">
            <CardPreview trainee={trainee} settings={settings} />
          </div>
          <div className="preview-caption">
            <span>النموذج: {traineeName(trainee)}</span>
            <button
              className="text-button"
              onClick={() => toast.success("تم Design enregistré localement")}
            >
              <Save size={15} /> Enregistrer le design
            </button>
          </div>
        </section>
      </div>
    </div>
  );
}

function PrintView({
  trainees,
  settings,
}: {
  trainees: Trainee[];
  settings: Settings;
}) {
  const [selected, setSelected] = useState(trainees.slice(0, 9).map(t => t.id));
  useEffect(() => {
    setSelected(trainees.slice(0, 9).map(t => t.id));
  }, [trainees.length]);
  const chosen = trainees.filter(t => selected.includes(t.id));
  const pages = chosen.length
    ? Array.from({ length: Math.ceil(chosen.length / 9) }, (_, pageIndex) =>
        chosen.slice(pageIndex * 9, pageIndex * 9 + 9)
      )
    : [[] as Trainee[]];
  const toggle = (id: string) =>
    setSelected(
      selected.includes(id) ? selected.filter(x => x !== id) : [...selected, id]
    );
  return (
    <div className="page-content print-workspace">
      <div className="view-toolbar">
        <div>
          <span className="section-kicker">Prêt à imprimer</span>
          <h2 className="page-title">Cartes sur une feuille A4</h2>
          <p className="page-lead">
            Sélectionnez les cartes, puis imprimez ou choisissez « Enregistrer
            au format PDF ».
          </p>
        </div>
        <div className="print-actions">
          <button
            className="button button-ghost"
            onClick={() => window.print()}
          >
            <Printer size={17} /> Imprimer
          </button>
          <button
            className="button button-primary"
            onClick={() => {
              const previousTitle = document.title;
              document.title = "cartes-stagiaires";
              window.onafterprint = () => {
                document.title = previousTitle;
                window.onafterprint = null;
              };
              toast.info(
                "Dans la fenêtre d’impression, choisissez « Enregistrer au format PDF »."
              );
              window.print();
            }}
          >
            <Download size={17} /> Exporter PDF
          </button>
        </div>
      </div>
      <div className="print-layout">
        <section className="surface print-list">
          <div className="print-list-head">
            <div>
              <strong>Sélection des cartes</strong>
              <span>
                {selected.length} sur {trainees.length} sélectionnées
              </span>
            </div>
            <button
              className="link-button"
              onClick={() =>
                setSelected(
                  selected.length === trainees.length
                    ? []
                    : trainees.map(t => t.id)
                )
              }
            >
              {selected.length === trainees.length
                ? "Tout désélectionner"
                : "Tout sélectionner"}
            </button>
          </div>
          {trainees.map(t => (
            <label
              className={`print-select-row ${selected.includes(t.id) ? "selected" : ""}`}
              key={t.id}
            >
              <input
                type="checkbox"
                checked={selected.includes(t.id)}
                onChange={() => toggle(t.id)}
              />
              <div className="mini-print-avatar">
                {t.photo ? (
                  <img src={t.photo} alt="" />
                ) : (
                  traineeName(t).charAt(0)
                )}
              </div>
              <div>
                <strong>{traineeName(t)}</strong>
                <span>{t.registration}</span>
              </div>
              <BadgeCheck
                size={15}
                className={t.photo ? "text-green" : "text-muted"}
              />
            </label>
          ))}
        </section>
        <section className="print-preview-wrap">
          <div className="print-preview-toolbar">
            <span>Feuille A4 · grille 3 × 3</span>
            <span className="scale-note">Proportions CR80 conservées</span>
          </div>
          {pages.map((page, pageIndex) => (
            <div className="a4-sheet" key={`page-${pageIndex}`}>
              <div className="print-grid">
                {page.map(t => (
                  <CardPreview
                    key={t.id}
                    trainee={t}
                    settings={settings}
                    compact
                  />
                ))}
                {Array.from({ length: Math.max(0, 9 - page.length) }).map(
                  (_, i) => (
                    <div
                      key={`empty-${pageIndex}-${i}`}
                      className="empty-print-cell"
                    >
                      <Plus size={16} />
                    </div>
                  )
                )}
              </div>
            </div>
          ))}
        </section>
      </div>
    </div>
  );
}

function SettingsView({
  trainees,
  settings,
  setTrainees,
  setSettings,
}: {
  trainees: Trainee[];
  settings: Settings;
  setTrainees: (t: Trainee[]) => void;
  setSettings: (s: Settings) => void;
}) {
  const reset = () => {
    if (
      confirm("Réinitialiser les données de démonstration وحذف السجل الحالي؟")
    ) {
      setTrainees(demoTrainees);
      setSettings(defaultSettings);
      toast.success("تمت Réinitialisation");
    }
  };
  return (
    <div className="page-content">
      <div className="view-toolbar">
        <div>
          <span className="section-kicker">Espace local</span>
          <h2 className="page-title">Paramètres et sauvegarde</h2>
          <p className="page-lead">
            Aucune base de données externe — tout est enregistré dans ce
            navigateur.
          </p>
        </div>
      </div>
      <div className="settings-grid">
        <section className="surface settings-card">
          <div className="settings-icon">
            <Archive size={20} />
          </div>
          <h3>Stockage local</h3>
          <p>
            بيانات Stagiaireين، صورهم، وإعدادات البطاقة محفوظة داخل هذا المتصفح
            باستخدام localStorage.
          </p>
          <div className="setting-status">
            <span className="status-check">
              <Check size={14} />
            </span>
            <div>
              <strong>تخزين نشط</strong>
              <small>{trainees.length} سجلاً محفوظاً</small>
            </div>
          </div>
        </section>
        <section className="surface settings-card">
          <div className="settings-icon">
            <FileText size={20} />
          </div>
          <h3>Fichier du projet</h3>
          <p>
            أنشئ نسخة JSON احتياطية بانتظام، خاصة قبل نقل المشروع إلى حاسوب آخر
            أو تنظيف بيانات المتصفح.
          </p>
          <div className="settings-note">
            <ShieldCheck size={16} />
            <span>النسخ الاحتياطي لا يحتوي على أي بيانات خارج جهازك.</span>
          </div>
        </section>
        <section className="surface settings-card danger-card">
          <div className="settings-icon">
            <RotateCcw size={20} />
          </div>
          <h3>Réinitialisation</h3>
          <p>
            إرجاع التطبيق إلى بيانات العرض التجريبية وإعدادات البطاقة
            الافتراضية.
          </p>
          <button className="button button-danger" onClick={reset}>
            <RotateCcw size={15} /> Réinitialiser les données de démonstration
          </button>
        </section>
      </div>
    </div>
  );
}

function TraineeEditor({
  trainee,
  onSave,
  onClose,
}: {
  trainee?: Trainee;
  onSave: (t: Trainee) => void;
  onClose: () => void;
}) {
  const [form, setForm] = useState<Trainee>(
    trainee || {
      id: uid(),
      firstName: "",
      lastName: "",
      fullName: "",
      registration: "",
      passport: "",
      studyYear: "",
      specialty: "",
      level: "",
    }
  );
  const imageRef = useRef<HTMLInputElement>(null);
  const set = (key: keyof Trainee, value: string) =>
    setForm({ ...form, [key]: value });
  const chooseImage = (file?: File) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => set("photo", String(reader.result));
    reader.readAsDataURL(file);
  };
  const submit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!form.fullName?.trim() || !form.registration) {
      toast.error("أدخل الاسم واللقب وMatricule");
      return;
    }
    onSave(form);
  };
  return (
    <div
      className="modal-backdrop"
      onMouseDown={e => e.target === e.currentTarget && onClose()}
    >
      <form className="editor-modal" onSubmit={submit}>
        <div className="modal-head">
          <div>
            <span className="section-kicker">Carte stagiaire</span>
            <h2>
              {trainee
                ? "Modifier les données"
                : "Ajouter un nouveau stagiaire"}
            </h2>
          </div>
          <button type="button" className="icon-button" onClick={onClose}>
            <X size={18} />
          </button>
        </div>
        <div className="editor-body">
          <div className="photo-upload">
            <div className="editor-photo">
              {form.photo ? (
                <img src={form.photo} alt="" />
              ) : (
                <UserRound size={27} />
              )}
            </div>
            <input
              ref={imageRef}
              type="file"
              hidden
              accept="image/*"
              onChange={e => chooseImage(e.target.files?.[0])}
            />
            <button
              type="button"
              className="text-button"
              onClick={() => imageRef.current?.click()}
            >
              <ImagePlus size={15} />{" "}
              {form.photo ? "Changer la photo" : "Ajouter une photo"}
            </button>
            <small>JPG أو PNG · صورة شخصية</small>
          </div>
          <div className="editor-fields">
            <label>
              <span>
                Nom et prénom <b>*</b>
              </span>
              <input
                autoFocus
                value={form.fullName ?? traineeName(form)}
                onChange={e => setForm({ ...form, fullName: e.target.value })}
                placeholder="Exemple : Ali MOHAMMED"
              />
            </label>
            <label>
              <span>
                Matricule <b>*</b>
              </span>
              <input
                value={form.registration}
                onChange={e => set("registration", e.target.value)}
                placeholder="2025/00147"
              />
            </label>
            <label>
              <span>رقم Passeport</span>
              <input
                dir="ltr"
                value={form.passport}
                onChange={e => set("passport", e.target.value)}
                placeholder="AA384921"
              />
            </label>
            <label>
              <span>Année scolaire</span>
              <input
                dir="ltr"
                value={form.studyYear}
                onChange={e => set("studyYear", e.target.value)}
                placeholder="2025 — 2026"
              />
            </label>
            <label>
              <span>Spécialité</span>
              <input
                value={form.specialty}
                onChange={e => set("specialty", e.target.value)}
                placeholder="Informatique"
              />
            </label>
            <label>
              <span>Niveau</span>
              <input
                value={form.level}
                onChange={e => set("level", e.target.value)}
                placeholder="Technicien supérieur"
              />
            </label>
          </div>
        </div>
        <div className="modal-foot">
          <button
            type="button"
            className="button button-ghost"
            onClick={onClose}
          >
            Annuler
          </button>
          <button type="submit" className="button button-primary">
            <Save size={16} /> Enregistrer le stagiaire
          </button>
        </div>
      </form>
    </div>
  );
}

export default function Home() {
  const initial = useMemo(loadState, []);
  const [trainees, setTraineesState] = useState<Trainee[]>(initial.trainees);
  const [settings, setSettingsState] = useState<Settings>(initial.settings);
  const [activeView, setActiveView] = useState<View>("dashboard");
  const [lang, setLang] = useState<Lang>(initial.settings.cardLanguage || "fr");
  const [editor, setEditor] = useState<{ open: boolean; trainee?: Trainee }>({
    open: false,
  });
  const [mobileOpen, setMobileOpen] = useState(false);
  const setTrainees = (next: Trainee[]) => {
    setTraineesState(next);
    saveState(next, settings);
  };
  const setSettings = (next: Settings) => {
    setSettingsState(next);
    saveState(trainees, next);
  };
  useEffect(() => saveState(trainees, settings), [trainees, settings]);
  const saveTrainee = (next: Trainee) => {
    const exists = trainees.some(t => t.id === next.id);
    const updated = exists
      ? trainees.map(t => (t.id === next.id ? next : t))
      : [next, ...trainees];
    setTrainees(updated);
    setEditor({ open: false });
    toast.success(exists ? "تم تحديث بيانات Stagiaire" : "تمت إضافة Stagiaire");
  };
  const importFile = async (file: File) => {
    try {
      if (file.name.endsWith(".json")) {
        const parsed = JSON.parse(await file.text());
        const nextTrainees = Array.isArray(parsed) ? parsed : parsed.trainees;
        if (!Array.isArray(nextTrainees)) throw new Error("invalid");
        setTrainees(
          nextTrainees.map((t: Trainee) => ({ ...t, id: t.id || uid() }))
        );
        if (parsed.settings) setSettings({ ...settings, ...parsed.settings });
        toast.success(`تمت استعادة ${nextTrainees.length} سجلاً`);
        return;
      }
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
        defval: "",
      });
      const value = (row: Record<string, unknown>, keys: string[]) => {
        const key = Object.keys(row).find(k =>
          keys.some(candidate => k.trim().toLowerCase().includes(candidate))
        );
        return key ? String(row[key] ?? "").trim() : "";
      };
      const imported = rows
        .map(row => {
          const fullName = value(row, [
            "nom et prénom",
            "nom et prenom",
            "full name",
            "fullname",
          ]);
          const legacyFirstName = value(row, [
            "الاسم",
            "prenom",
            "first name",
            "firstname",
          ]);
          const legacyLastName = value(row, [
            "اللقب",
            "nom",
            "last name",
            "lastname",
          ]);
          return {
            id: uid(),
            fullName:
              fullName ||
              [legacyFirstName, legacyLastName].filter(Boolean).join(" "),
            firstName: legacyFirstName,
            lastName: legacyLastName,
            registration: value(row, ["التسجيل", "matricule", "registration"]),
            passport: value(row, ["جواز", "passeport", "passport"]),
            studyYear: value(row, [
              "السنة",
              "année",
              "annee",
              "study year",
              "school year",
            ]),
            specialty: value(row, [
              "Spécialité",
              "specialite",
              "spécialité",
              "specialty",
            ]),
            level: value(row, ["Niveau", "niveau", "level"]),
          };
        })
        .filter(t => t.fullName || t.registration);
      if (!imported.length) throw new Error("empty");
      setTrainees([...imported, ...trainees]);
      toast.success(`تم استيراد ${imported.length} متربصاً بنجاح`);
      setActiveView("trainees");
    } catch {
      toast.error("تعذر قراءة الملف. تحقق sur التنسيق ثم حاول مجدداً.");
    }
  };
  const exportBackup = () =>
    downloadBlob(
      JSON.stringify(
        {
          version: 1,
          exportedAt: new Date().toISOString(),
          trainees,
          settings,
        },
        null,
        2
      ),
      `بطاقات-Stagiaireين-${new Date().toISOString().slice(0, 10)}.json`
    );
  const currentTrainee = trainees[0] || demoTrainees[0];
  return (
    <div className="app-shell" dir={lang === "fr" ? "ltr" : "rtl"}>
      <div
        className={`mobile-overlay ${mobileOpen ? "show" : ""}`}
        onClick={() => setMobileOpen(false)}
      />
      <div className={`sidebar-wrap ${mobileOpen ? "open" : ""}`}>
        <Sidebar
          activeView={activeView}
          setActiveView={setActiveView}
          traineesCount={trainees.length}
          onClose={() => setMobileOpen(false)}
        />
      </div>
      <main className="main-area">
        <Topbar
          activeView={activeView}
          lang={lang}
          setLang={next => {
            setLang(next);
            setSettings({ ...settings, cardLanguage: next });
          }}
          onMenu={() => setMobileOpen(true)}
        />
        {activeView === "dashboard" && (
          <Dashboard
            trainees={trainees}
            settings={settings}
            setView={setActiveView}
            onEdit={t => setEditor({ open: true, trainee: t })}
          />
        )}
        {activeView === "trainees" && (
          <TraineesView
            trainees={trainees}
            setTrainees={setTrainees}
            settings={settings}
            openEditor={t => setEditor({ open: true, trainee: t })}
          />
        )}
        {activeView === "import" && (
          <ImportView
            onImport={importFile}
            onExport={exportBackup}
            trainees={trainees}
          />
        )}
        {activeView === "design" && (
          <DesignView
            settings={settings}
            setSettings={setSettings}
            trainee={currentTrainee}
          />
        )}
        {activeView === "print" && (
          <PrintView trainees={trainees} settings={settings} />
        )}
        {activeView === "settings" && (
          <SettingsView
            trainees={trainees}
            settings={settings}
            setTrainees={setTrainees}
            setSettings={setSettings}
          />
        )}
      </main>
      {editor.open && (
        <TraineeEditor
          trainee={editor.trainee}
          onSave={saveTrainee}
          onClose={() => setEditor({ open: false })}
        />
      )}
    </div>
  );
}
