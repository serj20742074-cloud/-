/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import pptxgen from 'pptxgenjs';
import {
  LayoutDashboard,
  Train,
  Zap,
  ZapOff,
  TrendingUp,
  TrendingDown,
  Tag,
  Filter,
  FileText,
  BookOpen,
  Settings,
  Plus,
  Edit2,
  Trash2,
  Download,
  Upload,
  AlertTriangle,
  Search,
  Moon,
  Sun,
  Check,
  X,
  ChevronRight,
  RefreshCw,
  Activity,
  Calendar,
  ExternalLink,
  Presentation
} from 'lucide-react';

import {
  Station,
  SupplyPoint,
  Reading,
  LossObject,
  LossReading,
  INITIAL_STATIONS,
  INITIAL_SUPPLY_POINTS,
  INITIAL_LOSS_OBJECTS,
  INITIAL_CATEGORIES,
  generateDeterministicDatabase
} from './db/initialData';

import { dbGet, dbSet } from './db/indexedDb';

import {
  getAllAnomalies,
  Anomaly,
  detectStationAnomalies,
  detectSupplyPointAnomalies,
  detectLossObjectAnomalies
} from './db/anomalyDetector';

const RUSSIAN_MONTHS = [
  "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
  "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
];

const RUSSIAN_MONTHS_SHORT = [
  "Янв", "Фев", "Мар", "Апр", "Май", "Июн",
  "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"
];

const PERIOD_OPTIONS = [
  { id: "1", name: "Январь", months: [1] },
  { id: "2", name: "Февраль", months: [2] },
  { id: "3", name: "Март", months: [3] },
  { id: "4", name: "Апрель", months: [4] },
  { id: "5", name: "Май", months: [5] },
  { id: "6", name: "Июнь", months: [6] },
  { id: "7", name: "Июль", months: [7] },
  { id: "8", name: "Август", months: [8] },
  { id: "9", name: "Сентябрь", months: [9] },
  { id: "10", name: "Октябрь", months: [10] },
  { id: "11", name: "Ноябрь", months: [11] },
  { id: "12", name: "Декабрь", months: [12] },
  { id: "2m", name: "2 месяца", months: [1, 2] },
  { id: "q1", name: "1 квартал", months: [1, 2, 3] },
  { id: "4m", name: "4 месяца", months: [1, 2, 3, 4] },
  { id: "5m", name: "5 месяцев", months: [1, 2, 3, 4, 5] },
  { id: "q2", name: "2 квартал", months: [4, 5, 6] },
  { id: "h1", name: "Полугодие", months: [1, 2, 3, 4, 5, 6] },
  { id: "7m", name: "7 месяцев", months: [1, 2, 3, 4, 5, 6, 7] },
  { id: "8m", name: "8 месяцев", months: [1, 2, 3, 4, 5, 6, 7, 8] },
  { id: "q3", name: "3 квартал", months: [7, 8, 9] },
  { id: "9m", name: "9 месяцев", months: [1, 2, 3, 4, 5, 6, 7, 8, 9] },
  { id: "10m", name: "10 месяцев", months: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10] },
  { id: "11m", name: "11 месяцев", months: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11] },
  { id: "q4", name: "4 квартал", months: [10, 11, 12] },
  { id: "12m", name: "12 месяцев", months: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] }
];

export default function App() {
  // --- CORE DATABASE STATE ---
  const [stations, setStations] = useState<Station[]>([]);
  const [supplyPoints, setSupplyPoints] = useState<SupplyPoint[]>([]);
  const [categories, setCategories] = useState<string[]>([]);
  const [readings, setReadings] = useState<Reading[]>([]);
  const [lossObjects, setLossObjects] = useState<LossObject[]>([]);
  const [lossReadings, setLossReadings] = useState<LossReading[]>([]);

  // Themes & UI tabs state
  const [darkTheme, setDarkTheme] = useState<boolean>(true);
  const [activeTab, setActiveTab] = useState<string>('home');
  const [selectedPeriodId, setSelectedPeriodId] = useState<string>('6'); // Defaults to June (month 6)
  const [selectedMonth, setSelectedMonth] = useState<number>(6); // June
  const [selectedYear, setSelectedYear] = useState<number>(2026);

  // Drilldowns
  const [selectedStationId, setSelectedStationId] = useState<string>('st-alexandrov');
  const [selectedLossId, setSelectedLossId] = useState<string>('loss-trans-1');
  const [selectedSupplyPointId, setSelectedSupplyPointId] = useState<string | null>(null);
  const [selectedReportStationId, setSelectedReportStationId] = useState<string | null>(null);
  const [selectedReportSupplyPointId, setSelectedReportSupplyPointId] = useState<string | null>(null);
  const [reportSubTab, setReportSubTab] = useState<'stations' | 'categories' | 'accounting_methods' | 'losses' | 'slides_presentation' | 'detailed_analysis'>('stations');
  const [reportCalculationMethodFilter, setReportCalculationMethodFilter] = useState<'all' | 'meter' | 'estimated'>('all');
  const [selectedReportCategoryId, setSelectedReportCategoryId] = useState<string | null>(null);
  const [categoryCalculationMethodFilter, setCategoryCalculationMethodFilter] = useState<'all' | 'meter' | 'estimated'>('all');
  const [categoryStationSortOrder, setCategoryStationSortOrder] = useState<'worst-first' | 'best-first' | 'volume-desc' | 'volume-asc'>('worst-first');
  const [categoryPointSortOrder, setCategoryPointSortOrder] = useState<'worst-first' | 'best-first' | 'volume-desc' | 'volume-asc'>('worst-first');
  const [reportHoveredPoint, setReportHoveredPoint] = useState<{ month: number; year: number; value: number } | null>(null);
  const [reportLossSortOrder, setReportLossSortOrder] = useState<'worst-first' | 'best-first' | 'volume-desc' | 'volume-asc'>('worst-first');
  const [reportLossSearch, setReportLossSearch] = useState<string>('');
  const [reportSelectedLossObjectId, setReportSelectedLossObjectId] = useState<string | null>(null);
  const [reportHoveredSingleLossPoint, setReportHoveredSingleLossPoint] = useState<{ month: number; year: number; value: number } | null>(null);

  // Category Analytics Specific States (Criticism Zone & Expanding)
  const [categoryCriticismOnly, setCategoryCriticismOnly] = useState<boolean>(false);
  const [expandedCatStations, setExpandedCatStations] = useState<Record<string, boolean>>({});
  const [expandedCatPoints, setExpandedCatPoints] = useState<Record<string, boolean>>({});
  const [expandedDetailedStations, setExpandedDetailedStations] = useState<Record<string, boolean>>({});

  // Interactive Graph variables
  const [visibleYears, setVisibleYears] = useState<{ [key: number]: boolean }>({ 2025: true, 2026: true });
  const [hoveredPoint, setHoveredPoint] = useState<{ month: number; year: number; value: number } | null>(null);

  // Searches
  const [stationSearch, setStationSearch] = useState<string>('');
  const [pointSearch, setPointSearch] = useState<string>('');

  // Modals / Creators state
  const [stationModal, setStationModal] = useState<{ isOpen: boolean; mode: 'add' | 'edit'; id?: string; name: string; section: string; note: string } | null>(null);
  const [pointModal, setPointModal] = useState<{ isOpen: boolean; mode: 'add' | 'edit'; id?: string; name: string; stationId: string; category: string; note: string; isActive: boolean; calculationMethod?: 'meter' | 'estimated'; currKwh?: number; prevKwh?: number } | null>(null);
  const [lossModal, setLossModal] = useState<{ isOpen: boolean; mode: 'add' | 'edit'; id?: string; name: string; stationId: string; section: string; note: string } | null>(null);
  const [categoryInput, setCategoryInput] = useState<string>('');
  const [deleteConfirm, setDeleteConfirm] = useState<{ isOpen: boolean; title: string; message: string; onConfirm: () => void } | null>(null);

  // Excel matching Resolution wizard
  const [importWizardQueue, setImportWizardQueue] = useState<{
    unknownPointName: string;
    proposedStationName: string;
    proposedCategory: string;
    rows: any[];
    meter?: string;
    yearsValues?: { year: number; value: number }[];
  }[]>([]);
  const [importStatus, setImportStatus] = useState<string | null>(null);

  // Mocks matching Huawei MatePad 11.5 indicators
  const [batteryLevel, setBatteryLevel] = useState<number>(92);
  const [currentTimeStr, setCurrentTimeStr] = useState<string>('08:30');
  const [isOnline, setIsOnline] = useState<boolean>(navigator.onLine);

  // States for the anomalies list modal and its filters
  const [anomaliesModalOpen, setAnomaliesModalOpen] = useState<boolean>(false);
  const [navigatedFromAnomalies, setNavigatedFromAnomalies] = useState<boolean>(false);
  const [selectedAnomalyId, setSelectedAnomalyId] = useState<string | null>(null);
  const [anomalySearchQuery, setAnomalySearchQuery] = useState<string>('');
  const [anomalySeverityFilter, setAnomalySeverityFilter] = useState<'all' | 'high' | 'medium' | 'low'>('all');
  const [anomalyTypeFilter, setAnomalyTypeFilter] = useState<'all' | 'station' | 'supply_point' | 'loss'>('all');
  const [selectedAnomalyForDynamics, setSelectedAnomalyForDynamics] = useState<Anomaly | null>(null);

  // Backups state configuration
  interface BackupItem {
    id: string;
    timestamp: string;
    name: string;
    data: string;
  }

  const [autoBackupEnabled, setAutoBackupEnabled] = useState<boolean>(() => {
    const saved = localStorage.getItem('st_auto_backup');
    return saved !== null ? saved === 'true' : true;
  });

  const [backupsList, setBackupsList] = useState<BackupItem[]>([]);
  const [isDataLoading, setIsDataLoading] = useState<boolean>(true);

  // --- PERSISTENCE MOUNT INITIALIZER ---
  useEffect(() => {
    // Android clock tick simulation
    const tickTime = () => {
      const d = new Date();
      setCurrentTimeStr(`${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`);
    };
    tickTime();
    const timer = setInterval(tickTime, 30000);

    // Dynamic battery monitoring (if browser supports it)
    let batteryInstance: any = null;
    const handleLevelChange = () => {
      if (batteryInstance) {
        setBatteryLevel(Math.round(batteryInstance.level * 100));
      }
    };

    if ('getBattery' in navigator) {
      (navigator as any).getBattery().then((battery: any) => {
        batteryInstance = battery;
        setBatteryLevel(Math.round(battery.level * 100));
        battery.addEventListener('levelchange', handleLevelChange);
      }).catch((err: any) => {
        console.warn('Battery status not accessible:', err);
      });
    }

    // Dynamic network connectivity monitoring
    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);

    window.addEventListener('online', handleOnline);
    window.addEventListener('offline', handleOffline);

    return () => {
      clearInterval(timer);
      window.removeEventListener('online', handleOnline);
      window.removeEventListener('offline', handleOffline);
      if (batteryInstance) {
        batteryInstance.removeEventListener('levelchange', handleLevelChange);
      }
    };
  }, []);

  useEffect(() => {
    setSelectedSupplyPointId(null);
  }, [selectedStationId]);

  useEffect(() => {
    async function loadData() {
      try {
        // Try getting from IndexedDB first
        let cachedStations = await dbGet<Station[]>('st_stations');
        let cachedPoints = await dbGet<SupplyPoint[]>('st_supply_points');
        let cachedCategories = await dbGet<string[]>('st_categories');
        let cachedReadings = await dbGet<Reading[]>('st_readings');
        let cachedLossObjects = await dbGet<LossObject[]>('st_loss_objects');
        let cachedLossReadings = await dbGet<LossReading[]>('st_loss_readings');
        let cachedBackups = await dbGet<BackupItem[]>('st_backups_list');

        // Migration from localStorage if IndexedDB is empty but localStorage has data
        const oldStations = localStorage.getItem('st_stations');
        if (!cachedStations && oldStations) {
          try {
            cachedStations = JSON.parse(oldStations);
            cachedPoints = JSON.parse(localStorage.getItem('st_supply_points') || '[]');
            cachedCategories = JSON.parse(localStorage.getItem('st_categories') || '[]');
            cachedReadings = JSON.parse(localStorage.getItem('st_readings') || '[]');
            cachedLossObjects = JSON.parse(localStorage.getItem('st_loss_objects') || '[]');
            cachedLossReadings = JSON.parse(localStorage.getItem('st_loss_readings') || '[]');
            
            const oldBackups = localStorage.getItem('st_backups_list');
            if (oldBackups) {
              cachedBackups = JSON.parse(oldBackups);
            }

            // Save to IndexedDB so it is migrated permanently
            if (cachedStations) {
              await dbSet('st_stations', cachedStations);
              if (cachedPoints) await dbSet('st_supply_points', cachedPoints);
              if (cachedCategories) await dbSet('st_categories', cachedCategories);
              if (cachedReadings) await dbSet('st_readings', cachedReadings);
              if (cachedLossObjects) await dbSet('st_loss_objects', cachedLossObjects);
              if (cachedLossReadings) await dbSet('st_loss_readings', cachedLossReadings);
              if (cachedBackups) await dbSet('st_backups_list', cachedBackups);
            }
          } catch (migrationErr) {
            console.error("Migration from localStorage failed:", migrationErr);
          }
        }

        if (cachedStations && cachedStations.length > 0) {
          setStations(cachedStations);
          setSupplyPoints(cachedPoints || []);
          setCategories(cachedCategories || []);
          setReadings(cachedReadings || []);
          setLossObjects(cachedLossObjects || []);
          setLossReadings(cachedLossReadings || []);
          setBackupsList(cachedBackups || []);
        } else {
          // No cache, generate default database
          const db = generateDeterministicDatabase();
          setStations(INITIAL_STATIONS);
          setSupplyPoints(INITIAL_SUPPLY_POINTS);
          setCategories(INITIAL_CATEGORIES);
          setReadings(db.readings);
          setLossObjects(INITIAL_LOSS_OBJECTS);
          setLossReadings(db.lossReadings);
          setBackupsList([]);
          
          await saveToStorage(INITIAL_STATIONS, INITIAL_SUPPLY_POINTS, INITIAL_CATEGORIES, db.readings, INITIAL_LOSS_OBJECTS, db.lossReadings);
        }
      } catch (err) {
        console.error("Error loading data from IndexedDB", err);
      } finally {
        setIsDataLoading(false);
      }
    }

    loadData();

    const savedTheme = localStorage.getItem('st_theme');
    if (savedTheme === 'light') setDarkTheme(false);
  }, []);

  const saveToStorage = async (
    st: Station[], sps: SupplyPoint[], cats: string[],
    rds: Reading[], los: LossObject[], lrs: LossReading[]
  ) => {
    try {
      await dbSet('st_stations', st);
      await dbSet('st_supply_points', sps);
      await dbSet('st_categories', cats);
      await dbSet('st_readings', rds);
      await dbSet('st_loss_objects', los);
      await dbSet('st_loss_readings', lrs);
    } catch (e) {
      console.warn("IndexedDB write failed", e);
    }

    // Auto backup on change if enabled
    let isAutoOn = true;
    try {
      isAutoOn = localStorage.getItem('st_auto_backup') !== 'false';
    } catch (err) {
      // Ignored
    }
    if (isAutoOn) {
      try {
        const payload = JSON.stringify({ stations: st, supplyPoints: sps, categories: cats, readings: rds, lossObjects: los, lossReadings: lrs });
        const d = new Date();
        const timeStr = `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}`;
        const newBackup: BackupItem = {
          id: `backup-${Date.now()}`,
          timestamp: d.toLocaleString('ru-RU'),
          name: `Авто-копия системы (${timeStr})`,
          data: payload
        };
        const currentList: BackupItem[] = await dbGet<BackupItem[]>('st_backups_list') || [];
        const updated = [newBackup, ...currentList].slice(0, 15);
        await dbSet('st_backups_list', updated);
        
        // Sync React state without blocking standard main thread execution and causing re-render loops
        setTimeout(() => {
          setBackupsList(updated);
        }, 0);
      } catch (e) {
        console.error("Backup failed", e);
      }
    }
  };

  const syncDatabase = (
    st: Station[], sps: SupplyPoint[], cats: string[],
    rds: Reading[], los: LossObject[], lrs: LossReading[]
  ) => {
    setStations(st);
    setSupplyPoints(sps);
    setCategories(cats);
    setReadings(rds);
    setLossObjects(los);
    setLossReadings(lrs);
    saveToStorage(st, sps, cats, rds, los, lrs);
  };

  const toggleTheme = () => {
    const next = !darkTheme;
    setDarkTheme(next);
    localStorage.setItem('st_theme', next ? 'dark' : 'light');
  };

  const selectedMonthsList = useMemo(() => {
    const period = PERIOD_OPTIONS.find(p => p.id === selectedPeriodId);
    return period ? period.months : [selectedMonth];
  }, [selectedPeriodId, selectedMonth]);

  const selectedPeriodName = useMemo(() => {
    const period = PERIOD_OPTIONS.find(p => p.id === selectedPeriodId);
    return period ? period.name : RUSSIAN_MONTHS[selectedMonth - 1];
  }, [selectedPeriodId, selectedMonth]);

  // --- MEMOIZED DERIVED METRICS FOR SELECTED PERIOD ---
  const activeMonthReadings = useMemo(() => {
    return readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month));
  }, [readings, selectedYear, selectedMonthsList]);

  const prevYearMonthReadings = useMemo(() => {
    return readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month));
  }, [readings, selectedYear, selectedMonthsList]);

  const activeLossReadings = useMemo(() => {
    return lossReadings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month));
  }, [lossReadings, selectedYear, selectedMonthsList]);

  const prevYearLossReadings = useMemo(() => {
    return lossReadings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month));
  }, [lossReadings, selectedYear, selectedMonthsList]);

  // Overall calculations (Totals)
  const totalCurrValues = useMemo(() => activeMonthReadings.reduce((sum, r) => sum + r.value, 0), [activeMonthReadings]);
  const totalPrevValues = useMemo(() => prevYearMonthReadings.reduce((sum, r) => sum + r.value, 0), [prevYearMonthReadings]);
  const totalDeltaAbs = totalCurrValues - totalPrevValues;
  const totalDeltaPercent = totalPrevValues > 0 ? (totalDeltaAbs / totalPrevValues) * 100 : 0;

  const totalCurrLosses = useMemo(() => activeLossReadings.reduce((sum, r) => sum + r.value, 0), [activeLossReadings]);
  const totalPrevLosses = useMemo(() => prevYearLossReadings.reduce((sum, r) => sum + r.value, 0), [prevYearLossReadings]);
  const totalLossDeltaAbs = totalCurrLosses - totalPrevLosses;
  const totalLossDeltaPercent = totalPrevLosses > 0 ? (totalLossDeltaAbs / totalPrevLosses) * 100 : 0;

  // Anomalies List
  const currentAnomalies = useMemo(() => {
    return getAllAnomalies(stations, supplyPoints, lossObjects, readings, lossReadings, selectedYear, selectedMonth);
  }, [stations, supplyPoints, lossObjects, readings, lossReadings, selectedYear, selectedMonth]);

  // Overrun and Savings Leaders
  const leadersList = useMemo(() => {
    const list: { station: Station; current: number; prev: number; deltaAbs: number; deltaPct: number }[] = [];

    stations.forEach(st => {
      const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
      const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
      const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
      const deltaAbs = curr - prev;
      const deltaPct = prev > 0 ? (deltaAbs / prev) * 100 : 0;

      list.push({ station: st, current: curr, prev, deltaAbs, deltaPct });
    });

    const overruns = [...list].filter(l => l.deltaAbs > 0).sort((a, b) => b.deltaAbs - a.deltaAbs);
    const savings = [...list].filter(l => l.deltaAbs < 0).sort((a, b) => a.deltaAbs - b.deltaAbs);

    return { overruns, savings };
  }, [stations, supplyPoints, readings, selectedYear, selectedMonthsList]);

  // Station Detail Analytics
  const activeStation = useMemo(() => {
    return stations.find(s => s.id === selectedStationId) || stations[0] || null;
  }, [stations, selectedStationId]);

  const activeSupplyPoint = useMemo(() => {
    if (!activeStation || !selectedSupplyPointId) return null;
    return supplyPoints.find(p => p.id === selectedSupplyPointId && p.stationId === activeStation.id) || null;
  }, [supplyPoints, selectedSupplyPointId, activeStation]);

  const activeLossObject = useMemo(() => {
    return lossObjects.find(l => l.id === selectedLossId) || lossObjects[0] || null;
  }, [lossObjects, selectedLossId]);

  const stationActiveSum = useMemo(() => {
    if (!activeStation) return { current: 0, prev: 0, deltaAbs: 0, deltaPct: 0 };
    const spsIds = supplyPoints.filter(p => p.stationId === activeStation.id && p.isActive).map(p => p.id);
    const current = activeMonthReadings.filter(r => spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
    const prev = prevYearMonthReadings.filter(r => spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
    const deltaAbs = current - prev;
    const deltaPct = prev > 0 ? (deltaAbs / prev) * 100 : 0;
    return { current, prev, deltaAbs, deltaPct };
  }, [activeStation, supplyPoints, activeMonthReadings, prevYearMonthReadings]);

  const stationCategoryBreakdown = useMemo(() => {
    if (!activeStation) return [];
    const sps = supplyPoints.filter(p => p.stationId === activeStation.id && p.isActive);
    const result: { category: string; value: number; share: number }[] = [];

    let total = 0;
    sps.forEach(p => {
      const r = activeMonthReadings.find(rd => rd.supplyPointId === p.id);
      if (r) {
        total += r.value;
        const exists = result.find(v => v.category === p.category);
        if (exists) {
          exists.value += r.value;
        } else {
          result.push({ category: p.category, value: r.value, share: 0 });
        }
      }
    });

    result.forEach(r => { r.share = total > 0 ? (r.value / total) * 100 : 0; });
    return result.sort((a, b) => b.value - a.value);
  }, [activeStation, supplyPoints, activeMonthReadings]);

  // Global Category contribution
  const globalCategoryShares = useMemo(() => {
    const result: { [cat: string]: { value: number; share: number } } = {};
    categories.forEach(c => { result[c] = { value: 0, share: 0 }; });

    supplyPoints.forEach(p => {
      const r = activeMonthReadings.find(rd => rd.supplyPointId === p.id);
      if (r && result[p.category]) {
        result[p.category].value += r.value;
      }
    });

    categories.forEach(c => {
      const v = result[c];
      v.share = totalCurrValues > 0 ? (v.value / totalCurrValues) * 100 : 0;
    });

    return result;
  }, [categories, supplyPoints, activeMonthReadings, totalCurrValues]);

  // --- EXPORT TO WORD (.DOC) ---
  const exportReportToDoc = () => {
    const currentMonthLabel = selectedPeriodName;
    const reportTitle = "АНАЛИТИЧЕСКИЙ ЭНЕРГЕТИЧЕСКИЙ ОТЧЕТ";
    const reportSubtitle = `Учетный период: ${currentMonthLabel} ${selectedYear} года (в сравнении с аналогичным периодом прошлого года)`;

    // Calculate overall totals
    const totalCurrent = stations.reduce((sum, st) => {
      const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
      return sum + readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
    }, 0);
    const totalPrev = stations.reduce((sum, st) => {
      const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
      return sum + readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
    }, 0);
    const totalDiff = totalCurrent - totalPrev;
    const totalPct = totalPrev > 0 ? (totalDiff / totalPrev) * 100 : 0;

    // Calculate details and rankings for stations
    const stationDiffs = stations.map(st => {
      const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
      const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
      const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
      const diff = curr - prev;
      const pct = prev > 0 ? (diff / prev) * 100 : 0;
      return {
        id: st.id,
        name: st.name,
        curr,
        prev,
        diff,
        pct
      };
    });

    // worst stations (overruns): diff > 0, sorted worst to best (largest overrun to smallest overrun)
    const worstStations = stationDiffs
      .filter(s => s.diff > 0)
      .sort((a, b) => b.diff - a.diff)
      .slice(0, 5);

    // best stations (savings): diff < 0, sorted best to worst (largest savings to smallest savings)
    const bestStations = stationDiffs
      .filter(s => s.diff < 0)
      .sort((a, b) => a.diff - b.diff)
      .slice(0, 5);

    // Deep categoric audit focusing on important categories
    const targetCategories = ["Освещение горловин", "Бытовые нагрузки", "Отопление"];
    const categoryInfo = targetCategories.map(cat => {
      const catPoints = supplyPoints.filter(p => p.category === cat && p.isActive);
      const pointsData = catPoints.map(p => {
        const station = stations.find(s => s.id === p.stationId);
        const currVal = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
        const prevVal = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
        const diffVal = currVal - prevVal;
        const pctVal = prevVal > 0 ? (diffVal / prevVal) * 100 : 0;

        return {
          pointName: p.name,
          stationName: station?.name || "Вне станций",
          curr: currVal,
          prev: prevVal,
          diff: diffVal,
          pct: pctVal
        };
      }).sort((a, b) => b.diff - a.diff); // Put highest overrun points first

      const catTotalCurr = pointsData.reduce((sum, d) => sum + d.curr, 0);
      const catTotalPrev = pointsData.reduce((sum, d) => sum + d.prev, 0);
      const catTotalDiff = catTotalCurr - catTotalPrev;
      const catTotalPct = catTotalPrev > 0 ? (catTotalDiff / catTotalPrev) * 100 : 0;

      return {
        category: cat,
        curr: catTotalCurr,
        prev: catTotalPrev,
        diff: catTotalDiff,
        pct: catTotalPct,
        details: pointsData
      };
    });

    // Losses Analysis Calculations
    const totalLossCurrent = lossReadings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
    const totalLossPrev = lossReadings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
    const totalLossDiff = totalLossCurrent - totalLossPrev;
    const totalLossPct = totalLossPrev > 0 ? (totalLossDiff / totalLossPrev) * 100 : 0;

    // Generate comprehensive HTML that Word perfectly understands
    const htmlDocContent = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8">
        <title>Энергетический отчет</title>
        <style>
          body {
            font-family: 'Arial', sans-serif;
            font-size: 11pt;
            line-height: 1.5;
            color: #333333;
          }
          .header-box {
            text-align: center;
            border-bottom: 3px double #1a365d;
            padding-bottom: 15px;
            margin-bottom: 25px;
          }
          .title {
            font-size: 18pt;
            font-weight: bold;
            color: #1a365d;
          }
          .subtitle {
            font-size: 11pt;
            color: #4a5568;
            font-style: italic;
            margin-top: 5px;
          }
          h2 {
            font-size: 13.5pt;
            font-weight: bold;
            color: #2b6cb0;
            border-bottom: 1.5px solid #2b6cb0;
            padding-bottom: 2px;
            margin-top: 25px;
            margin-bottom: 10px;
          }
          h3 {
            font-size: 11.5pt;
            font-weight: bold;
            color: #2d3748;
            margin-top: 15px;
            margin-bottom: 6px;
          }
          p {
            text-align: justify;
            margin-bottom: 10px;
            text-indent: 1cm;
          }
          .lead-summary {
            background-color: #ebf8ff;
            border-left: 5px solid #3182ce;
            padding: 12px;
            margin-bottom: 20px;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 18px;
            font-size: 9.5pt;
          }
          th {
            background-color: #1a365d;
            color: #ffffff;
            font-weight: bold;
            padding: 6px 8px;
            border: 1px solid #cbd5e0;
            text-align: left;
          }
          td {
            padding: 6px 8px;
            border: 1px solid #cbd5e0;
          }
          .text-right {
            text-align: right;
          }
          .saving {
            color: #047857;
            font-weight: bold;
          }
          .overrun {
            color: #b91c1c;
            font-weight: bold;
          }
          .neutral {
            color: #4a5568;
          }
          .badge {
            background-color: #edf2f7;
            padding: 1px 5px;
            border-radius: 3px;
            font-size: 8pt;
            font-weight: bold;
          }
          .badge-saving {
            background-color: #ecfdf5;
            color: #047857;
          }
          .badge-overrun {
            background-color: #fef2f2;
            color: #b91c1c;
          }
          .rec-list {
            padding-left: 20px;
            margin-bottom: 20px;
          }
          .rec-list li {
            margin-bottom: 6px;
            text-align: justify;
          }
          .footer-box {
            margin-top: 35px;
            font-size: 8.5pt;
            color: #718096;
            text-align: center;
            border-top: 1px solid #cbd5e0;
            padding-top: 10px;
          }
        </style>
      </head>
      <body>
        <div class="header-box">
          <div class="title">${reportTitle}</div>
          <div class="subtitle">Информационно-аналитический комплекс учета энергоресурсов ОАО РЖД</div>
          <div class="subtitle" style="font-weight: bold; color: #1a365d;">${reportSubtitle}</div>
        </div>

        <div class="lead-summary">
          <strong>ОБЩИЕ ПОКАЗАТЕЛИ ПОТРЕБЛЕНИЯ ЭЛЕКТРОЭНЕРГИИ ЗА ПЕРИОД:</strong><br/>
          Суммарное фактическое электропотребление по всем хозяйственным объектам дороги за отчетный период составило 
          <strong>${totalCurrent.toLocaleString()} кВт·ч</strong>.<br/>
          За аналогичный учетный период прошлого года потребление составило 
          <strong>${totalPrev.toLocaleString()} кВт·ч</strong>.<br/>
          Изменение показателя (динамика): 
          <span class="${totalDiff > 0 ? "overrun" : "saving"}">
            <strong>${totalDiff > 0 ? `+${totalDiff.toLocaleString()}` : totalDiff.toLocaleString()} кВт·ч</strong> 
            (${totalPrev > 0 ? (totalDiff > 0 ? `+${totalPct.toFixed(2)}` : totalPct.toFixed(2)) : "0.00"}%)
          </span>.<br/>
          Энергетический баланс периода оценивается как: <strong>${totalDiff > 0 ? "ОТРИЦАТЕЛЬНЫЙ (ЗАФИКСИРОВАН СИСТЕМНЫЙ ПЕРЕРАСХОД)" : "ПОЛОЖИТЕЛЬНЫЙ (ДОСТИГНУТА СИСТЕМНАЯ ЭКОНОМИЯ)"}</strong>.
        </div>

        <h2>РАЗДЕЛ 1. РЕЙТИНГ ЭНЕРГОЭФФЕКТИВНОСТИ ЖЕЛЕЗНОДОРОЖНЫХ СТАНЦИЙ</h2>
        
        <h3>1.1. Станции с наибольшим перерасходом электроэнергии (Лидеры роста расхода, не более 5 объектов)</h3>
        <p>Ниже представлены железнодорожные станции, допустившие превышение расходов электроэнергии относительно аналогичного периода прошлого года, отсортированные от худших к лучшим (в порядке уменьшения абсолютного перерасхода):</p>

        <table>
          <thead>
            <tr>
              <th>Железнодорожная станция</th>
              <th class="text-right">Текущее потребление (кВт·ч)</th>
              <th class="text-right">Прошлый год (кВт·ч)</th>
              <th class="text-right">Перерасход (кВт·ч)</th>
              <th class="text-right">Отклонение (%)</th>
            </tr>
          </thead>
          <tbody>
            ${worstStations.length === 0 ? `
              <tr>
                <td colspan="5" style="text-align: center; color: #047857; font-weight: bold; padding: 12px;">Станций с перерасходом электроэнергии не зафиксировано. Все подразделения находятся в зоне экономии.</td>
              </tr>
            ` : worstStations.map(row => `
              <tr>
                <td><strong>${row.name}</strong></td>
                <td class="text-right font-mono">${row.curr.toLocaleString()}</td>
                <td class="text-right font-mono" style="color: #4a5568;">${row.prev.toLocaleString()}</td>
                <td class="text-right font-mono overrun">+${row.diff.toLocaleString()}</td>
                <td class="text-right font-mono overrun">+${row.pct.toFixed(2)}%</td>
              </tr>
            `).join("")}
          </tbody>
        </table>

        <h3>1.2. Станции с наибольшей экономией ресурсов (Лидеры снижения расхода, не более 5 объектов)</h3>
        <p>Ниже представлены железнодорожные станции, достигшие наибольшей экономии электрической энергии относительно сопоставимого прошлого года, отсортированные от лучших (с наибольшим сбережением) к худшим:</p>

        <table>
          <thead>
            <tr>
              <th>Железнодорожная станция</th>
              <th class="text-right">Текущее потребление (кВт·ч)</th>
              <th class="text-right">Прошлый год (кВт·ч)</th>
              <th class="text-right">Экономия (кВт·ч)</th>
              <th class="text-right">Снижение (%)</th>
            </tr>
          </thead>
          <tbody>
            ${bestStations.length === 0 ? `
              <tr>
                <td colspan="5" style="text-align: center; color: #718096; font-style: italic; padding: 12px;">Станций с чистым снижением энергопотребления за период не обнаружено.</td>
              </tr>
            ` : bestStations.map(row => `
              <tr>
                <td><strong>${row.name}</strong></td>
                <td class="text-right font-mono">${row.curr.toLocaleString()}</td>
                <td class="text-right font-mono" style="color: #4a5568;">${row.prev.toLocaleString()}</td>
                <td class="text-right font-mono saving">${row.diff.toLocaleString()}</td>
                <td class="text-right font-mono saving">${row.pct.toFixed(2)}%</td>
              </tr>
            `).join("")}
          </tbody>
        </table>

        <h2>РАЗДЕЛ 2. СЕГМЕНТНЫЙ АНАЛИЗ ЦЕЛЕВЫХ ГРУПП НАГРУЗОК</h2>
        <p>Сопоставительный аудит энергопотребления по ключевым технологическим категориям ("Освещение горловин", "Бытовые нагрузки", "Отопление") позволяет своевременно обнаруживать нештатный расход ресурсов:</p>

        ${categoryInfo.map(cat => {
          const catStatusClass = cat.diff > 0 ? "overrun" : cat.diff < 0 ? "saving" : "neutral";
          const catStatusLabel = cat.diff > 0 ? "ПЕРЕРАСХОД" : cat.diff < 0 ? "ЭКОНОМИЯ" : "В ПРЕДЕЛАХ НОРМЫ";
          const catBadgeClass = cat.diff > 0 ? "badge-overrun" : "badge-saving";

          return `
            <h3>2.1. Группа нагрузок: ${cat.category}</h3>
            <div style="margin-left: 15px; margin-bottom: 8px; font-size: 10pt;">
              Общее потребление группы: <strong>${cat.curr.toLocaleString()} кВт·ч</strong> 
              (предыдущий год: ${cat.prev.toLocaleString()} кВт·ч).<br/>
              Отклонение от нормы: <span class="${catStatusClass}"><strong>${cat.diff > 0 ? `+${cat.diff.toLocaleString()}` : cat.diff.toLocaleString()} кВт·ч</strong> (${cat.prev > 0 ? (cat.diff > 0 ? `+${cat.pct.toFixed(2)}` : cat.pct.toFixed(2)) : "0.00"}%)</span>. 
              Текущий статус: <span class="badge ${catBadgeClass}">${catStatusLabel}</span>
            </div>

            <table style="margin-top: 5px;">
              <thead>
                <tr>
                  <th>Железнодорожная станция</th>
                  <th>Наименование ТП (точка поставки)</th>
                  <th class="text-right">Потребление (кВт·ч)</th>
                  <th class="text-right">Прошлый год (кВт·ч)</th>
                  <th class="text-right">Разница (кВт·ч)</th>
                  <th class="text-right">Динамика</th>
                </tr>
              </thead>
              <tbody>
                ${cat.details.map(pt => {
                  const ptStatusClass = pt.diff > 0 ? "overrun" : pt.diff < 0 ? "saving" : "neutral";
                  const ptDiffText = pt.diff > 0 ? `+${pt.diff.toLocaleString()}` : pt.diff.toLocaleString();
                  const ptPctText = pt.prev > 0 ? (pt.diff > 0 ? `+${pt.pct.toFixed(1)}%` : `${pt.pct.toFixed(1)}%`) : "0.0%";
                  const ptLabel = pt.diff > 0 ? "Превышение" : pt.diff < 0 ? "Снижение" : "Без изменений";

                  return `
                    <tr>
                      <td>${pt.stationName}</td>
                      <td>${pt.pointName}</td>
                      <td class="text-right font-mono">${pt.curr.toLocaleString()}</td>
                      <td class="text-right font-mono" style="color: #718096;">${pt.prev.toLocaleString()}</td>
                      <td class="text-right font-mono ${ptStatusClass}"><strong>${ptDiffText}</strong></td>
                      <td class="text-right ${ptStatusClass}" style="font-size: 8.5pt;">${ptLabel} (${ptPctText})</td>
                    </tr>
                  `;
                }).join("")}
              </tbody>
            </table>
          `;
        }).join("")}

        <h2>РАЗДЕЛ 3. АНАЛИЗ ТЕХНОЛОГИЧЕСКИХ ПОТЕРЬ В ЭЛЕКТРИЧЕСКИХ СЕТЯХ</h2>
        <p>Сопоставительный учет и локализация технологических потерь (линии электропередачи, силовые понижающие трансформаторы, КТП) за отчетный период позволяет оценить КПД сетевой инфраструктуры:</p>

        <div style="margin-left: 15px; margin-bottom: 15px; font-size: 10pt; background-color: #f7fafc; border-left: 4px solid #4a5568; padding: 10px;">
          Суммарные фактические технологические потери в электрических сетях участка составили: <strong>${totalLossCurrent.toLocaleString()} кВт·ч</strong>.<br/>
          Аналогичный показатель за прошлый год: <strong>${totalLossPrev.toLocaleString()} кВт·ч</strong>.<br/>
          Абсолютное изменение уровня потерь: 
          <span class="${totalLossDiff > 0 ? "overrun" : "saving"}">
            <strong>${totalLossDiff > 0 ? `+${totalLossDiff.toLocaleString()}` : totalLossDiff.toLocaleString()} кВт·ч</strong> 
            (${totalLossPrev > 0 ? (totalLossDiff > 0 ? `+${totalLossPct.toFixed(2)}` : totalLossPct.toFixed(2)) : "0.00"}%)
          </span>.<br/>
          Динамика потерь оценивается как: <strong>${totalLossDiff > 0 ? "ОТРИЦАТЕЛЬНАЯ (РОСТ ХОЛОСТОГО ХОДА И ПОТЕРЬ В ПРОВОДНИКАХ)" : "ПОЛОЖИТЕЛЬНАЯ (ПОВЫШЕНИЕ ЭНЕРГЕТИЧЕСКОЙ ЭФФЕКТИВНОСТИ СЕТИ)"}</strong>.
        </div>

        <table>
          <thead>
            <tr>
              <th>Объект электрической сети</th>
              <th>Закрепленная станция</th>
              <th>Тип энергообъекта</th>
              <th class="text-right">Потери за текущий период (кВт·ч)</th>
              <th class="text-right">Потери за прошлый год (кВт·ч)</th>
              <th class="text-right">Изменение потерь (кВт·ч)</th>
              <th class="text-right">Динамика (%)</th>
            </tr>
          </thead>
          <tbody>
            ${lossObjects.map(lo => {
              const currL = lossReadings.filter(r => r.lossObjectId === lo.id && r.year === selectedYear && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
              const prevL = lossReadings.filter(r => r.lossObjectId === lo.id && r.year === selectedYear - 1 && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
              const diffL = currL - prevL;
              const pctL = prevL > 0 ? (diffL / prevL) * 100 : 0;
              const st = stations.find(s => s.id === lo.stationId);
              const loTypeLabel = (lo.name.toLowerCase().includes('ктп') || lo.name.toLowerCase().includes('трансф')) ? 'КТП / Трансформатор' : 'Фидер / ЛЭП';
              const diffText = diffL > 0 ? `+${diffL.toLocaleString()}` : diffL.toLocaleString();
              const pctText = prevL > 0 ? (diffL > 0 ? `+${pctL.toFixed(1)}%` : `${pctL.toFixed(1)}%`) : (currL > 0 ? "+100%" : "0.0%");
              const statusClass = diffL > 0 ? "overrun" : diffL < 0 ? "saving" : "neutral";

              return `
                <tr>
                  <td><strong>${lo.name}</strong></td>
                  <td>${st ? st.name : "Вне станций"}</td>
                  <td>${loTypeLabel}</td>
                  <td class="text-right font-mono">${currL.toLocaleString()}</td>
                  <td class="text-right font-mono" style="color: #718096;">${prevL.toLocaleString()}</td>
                  <td class="text-right font-mono ${statusClass}"><strong>${diffText}</strong></td>
                  <td class="text-right ${statusClass}">${pctText}</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>

        <h2>РАЗДЕЛ 4. ТЕХНИЧЕСКИЕ И ОРГАНИЗАЦИОННЫЕ РЕКОМЕНДАЦИИ</h2>
        <p>На основе полученной сопоставительной картины ведомостей расхода в учетном периоде сформированы адресно-целевые инструкции:</p>
        
        <ul class="rec-list">
          ${totalDiff > 0 ? `
            <li><strong style="color: #b91c1c;">Предупреждение по энергобюджету участка:</strong> В связи со сквозным превышением расхода на +${totalDiff.toLocaleString()} кВт·ч рекомендуется созвать оперативную комиссию службы электрификации для разбора полетов на станциях-лидерах роста.</li>
          ` : `
            <li><strong style="color: #047857;">Контроль энергоэффективности:</strong> Успешно достигнуто снижение расхода на ${Math.abs(totalDiff).toLocaleString()} кВт·ч. Рекомендуется закрепить регламенты рационального энергопотребления, снижающие фоновые внеплановые потери.</li>
          `}
          
          ${categoryInfo.find(c => c.category === "Освещение горловин")?.diff && (categoryInfo.find(c => c.category === "Освещение горловин")?.diff || 0) > 0 ? `
            <li><strong>Рациональное освещение горловин:</strong> Обнаружена негативная динамика по освещению горловин. Рекомендуется произвести проверку работоспособности суточных таймеров дежурного освещения и астрономических реле на фидерах. Организовать контрольный обход для выявления нецелевого горения прожекторов в дневное время суток.</li>
          ` : `
            <li><strong>Параметры осветительного оборудования:</strong> Прожекторные мачты путевого развития станций работают по утвержденному графику. Рекомендуется продолжать модернизацию светотехнического оборудования со ртутных ламп на энергосберегающие полупроводниковые прожекторы (LED).</li>
          `}

          ${categoryInfo.find(c => c.category === "Отопление")?.diff && (categoryInfo.find(c => c.category === "Отопление")?.diff || 0) > 0 ? `
            <li><strong>Аудит систем электрообогрева и отопления:</strong> Зафиксирован перерасход электроэнергии на тепловые цели. Рекомендуется проверить изоляцию оконных и дверных проемов постов ЭЦ, установить локальный контроль температурных режимов (+18°C для обитаемых зон и +14°C для необитаемых технических блок-контейнеров).</li>
          ` : `
            <li><strong>Параметры отопления:</strong> Тепловой режим зданий инфраструктурного комплекса находится в пределах нормы относительно климатических особенностей сезона.</li>
          `}

          ${categoryInfo.find(c => c.category === "Бытовые нагрузки")?.diff && (categoryInfo.find(c => c.category === "Бытовые нагрузки")?.diff || 0) > 0 ? `
            <li><strong>Инструктаж о нецелевом расходе:</strong> Превышение лимитов по бытовым нуждам указывает на потенциальное бесконтрольное использование электрических конвекторов, водонагревателей или холодильников устаревших классов энергопотребления. Требуется провести инструктаж линейного персонала о недопустимости нарушения внутреннего технологического порядка.</li>
          ` : `
            <li><strong>Хозяйственно-бытовой профиль:</strong> Расход на бытовые нужды стабилен и рационален.</li>
          `}
        </ul>

        <div class="footer-box">
          Аналитический отчет подготовлен и верифицирован автоматически в корпоративной ГИС-платформе.<br/>
          Дата генерации документа: ${new Date().toLocaleDateString("ru-RU")} в ${new Date().toLocaleTimeString("ru-RU")} | Электронная подпись оператора: serj20742074@gmail.com
        </div>
      </body>
      </html>
    `;

    // Write text to Blob
    const blob = new Blob(["\ufeff" + htmlDocContent], {
      type: "application/msword;charset=utf-8"
    });

    const url = URL.createObjectURL(blob);
    const downloadLink = document.createElement("a");
    downloadLink.href = url;
    downloadLink.download = `Аналитический_отчет_${selectedMonth.toString().padStart(2, '0')}_${selectedYear}.doc`;
    
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(url);
  };

  // --- EXPORT DETAILED CONSOLIDATED REPORT TO WORD (.DOC) ---
  const exportDetailedReportToDoc = () => {
    const currentPeriodLabel = selectedPeriodName;
    const reportTitle = "ПОДРОБНЫЙ СВОДНЫЙ АНАЛИЗ ЭНЕРГОПОТРЕБЛЕНИЯ РЕГИОНА";
    const reportSubtitle = `Учетный период: ${currentPeriodLabel} ${selectedYear} года (в сопоставлении с прошлым годом и накопительным итогом с начала года)`;

    // 1. Calculate overall totals
    const activePtIds = supplyPoints.filter(p => p.isActive).map(p => p.id);
    const totalCurrent = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && activePtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
    const totalPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && activePtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
    const totalDiff = totalCurrent - totalPrev;
    const totalPct = totalPrev > 0 ? (totalDiff / totalPrev) * 100 : 0;

    // YTD months list from 1 to selectedMonth
    const ytdMonthsList = Array.from({ length: selectedMonth }, (_, i) => i + 1);
    const totalCurrentYtd = readings.filter(r => r.year === selectedYear && ytdMonthsList.includes(r.month) && activePtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
    const totalPrevYtd = readings.filter(r => r.year === selectedYear - 1 && ytdMonthsList.includes(r.month) && activePtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
    const totalYtdDiff = totalCurrentYtd - totalPrevYtd;
    const totalYtdPct = totalPrevYtd > 0 ? (totalYtdDiff / totalPrevYtd) * 100 : 0;

    // 2. Category breakdown
    const categoryAnalysis = categories.map(cat => {
      const catPoints = supplyPoints.filter(p => p.category === cat && p.isActive);
      const catPtIds = catPoints.map(p => p.id);
      
      const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const diff = curr - prev;
      const pct = prev > 0 ? (diff / prev) * 100 : 0;

      const currYtd = readings.filter(r => r.year === selectedYear && ytdMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const prevYtd = readings.filter(r => r.year === selectedYear - 1 && ytdMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const diffYtd = currYtd - prevYtd;
      const pctYtd = prevYtd > 0 ? (diffYtd / prevYtd) * 100 : 0;

      return {
        category: cat,
        pointsCount: catPoints.length,
        curr, prev, diff, pct,
        currYtd, prevYtd, diffYtd, pctYtd
      };
    });

    // 3. Calculation Method breakdown
    const methods = [
      { id: 'meter', name: 'Приборы учета', filter: (p: any) => p.calculationMethod !== 'estimated' },
      { id: 'estimated', name: 'Расчетный способ (уст. мощность)', filter: (p: any) => p.calculationMethod === 'estimated' }
    ];

    const methodAnalysis = methods.map(m => {
      const mPoints = supplyPoints.filter(p => p.isActive && m.filter(p));
      const mPtIds = mPoints.map(p => p.id);

      const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && mPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && mPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const diff = curr - prev;
      const pct = prev > 0 ? (diff / prev) * 100 : 0;

      const currYtd = readings.filter(r => r.year === selectedYear && ytdMonthsList.includes(r.month) && mPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const prevYtd = readings.filter(r => r.year === selectedYear - 1 && ytdMonthsList.includes(r.month) && mPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const diffYtd = currYtd - prevYtd;
      const pctYtd = prevYtd > 0 ? (diffYtd / prevYtd) * 100 : 0;

      return {
        name: m.name,
        pointsCount: mPoints.length,
        curr, prev, diff, pct,
        currYtd, prevYtd, diffYtd, pctYtd
      };
    });

    // 4. Station Profile Breakdown
    const stationAnalysis = stations.map(st => {
      const stPoints = supplyPoints.filter(p => p.stationId === st.id && p.isActive);
      const stPtIds = stPoints.map(p => p.id);

      const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && stPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && stPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const diff = curr - prev;
      const pct = prev > 0 ? (diff / prev) * 100 : 0;

      const currYtd = readings.filter(r => r.year === selectedYear && ytdMonthsList.includes(r.month) && stPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const prevYtd = readings.filter(r => r.year === selectedYear - 1 && ytdMonthsList.includes(r.month) && stPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
      const diffYtd = currYtd - prevYtd;
      const pctYtd = prevYtd > 0 ? (diffYtd / prevYtd) * 100 : 0;

      // Group by category inside this station
      const stCatBreakdown = categories.map(cat => {
        const catPoints = stPoints.filter(p => p.category === cat);
        const catPtIds = catPoints.map(p => p.id);
        const catCurr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
        const catPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
        const catDiff = catCurr - catPrev;
        const catPct = catPrev > 0 ? (catDiff / catPrev) * 100 : 0;
        return { category: cat, curr: catCurr, prev: catPrev, diff: catDiff, pct: catPct };
      }).filter(cb => cb.curr > 0 || cb.prev > 0);

      // Detail list of points inside this station
      const stPointsDetails = stPoints.map(p => {
        const ptCurr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((s, r) => s + r.value, 0);
        const ptPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((s, r) => s + r.value, 0);
        const ptDiff = ptCurr - ptPrev;
        const ptPct = ptPrev > 0 ? (ptDiff / ptPrev) * 100 : 0;

        const ptCurrYtd = readings.filter(r => r.year === selectedYear && ytdMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((s, r) => s + r.value, 0);
        const ptPrevYtd = readings.filter(r => r.year === selectedYear - 1 && ytdMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((s, r) => s + r.value, 0);
        const ptDiffYtd = ptCurrYtd - ptPrevYtd;
        const ptPctYtd = ptPrevYtd > 0 ? (ptDiffYtd / ptPrevYtd) * 100 : 0;

        return {
          point: p,
          curr: ptCurr,
          prev: ptPrev,
          diff: ptDiff,
          pct: ptPct,
          currYtd: ptCurrYtd,
          prevYtd: ptPrevYtd,
          diffYtd: ptDiffYtd,
          pctYtd: ptPctYtd
        };
      });

      return {
        station: st,
        curr, prev, diff, pct,
        currYtd, prevYtd, diffYtd, pctYtd,
        categories: stCatBreakdown,
        points: stPointsDetails
      };
    });

    const htmlDocContent = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8">
        <title>Подробный сводный энергетический анализ</title>
        <style>
          body {
            font-family: 'Arial', sans-serif;
            font-size: 10.5pt;
            line-height: 1.4;
            color: #333333;
          }
          .header-box {
            text-align: center;
            border-bottom: 3px double #1e3a8a;
            padding-bottom: 12px;
            margin-bottom: 25px;
          }
          .title {
            font-size: 16pt;
            font-weight: bold;
            color: #1e3a8a;
            text-transform: uppercase;
          }
          .subtitle {
            font-size: 10pt;
            color: #475569;
            margin-top: 4px;
          }
          .period-subtitle {
            font-size: 10.5pt;
            font-weight: bold;
            color: #1e40af;
            margin-top: 6px;
          }
          h2 {
            font-size: 13pt;
            font-weight: bold;
            color: #1e3a8a;
            border-bottom: 1.5px solid #1e3a8a;
            padding-bottom: 3px;
            margin-top: 25px;
            margin-bottom: 12px;
            text-transform: uppercase;
          }
          h3 {
            font-size: 11.5pt;
            font-weight: bold;
            color: #0f172a;
            margin-top: 18px;
            margin-bottom: 8px;
            border-left: 4px solid #3b82f6;
            padding-left: 8px;
          }
          p {
            text-align: justify;
            margin-bottom: 10px;
            text-indent: 1cm;
          }
          .summary-card {
            background-color: #f8fafc;
            border: 1px solid #e2e8f0;
            border-left: 5px solid #2563eb;
            padding: 12px;
            margin-bottom: 20px;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 18px;
            font-size: 9pt;
          }
          th {
            background-color: #1e3a8a;
            color: #ffffff;
            font-weight: bold;
            padding: 6px 8px;
            border: 1px solid #cbd5e1;
            text-align: left;
          }
          td {
            padding: 5px 8px;
            border: 1px solid #cbd5e1;
          }
          .bg-gray {
            background-color: #f1f5f9;
          }
          .text-right {
            text-align: right;
          }
          .text-center {
            text-align: center;
          }
          .bold {
            font-weight: bold;
          }
          .saving {
            color: #16a34a;
            font-weight: bold;
          }
          .overrun {
            color: #dc2626;
            font-weight: bold;
          }
          .neutral {
            color: #475569;
          }
          .station-header {
            background-color: #f1f5f9;
            padding: 10px;
            border: 1px solid #cbd5e1;
            border-left: 4px solid #1e3a8a;
            margin-top: 20px;
            margin-bottom: 10px;
          }
          .footer-box {
            margin-top: 35px;
            font-size: 8.5pt;
            color: #64748b;
            text-align: center;
            border-top: 1px solid #e2e8f0;
            padding-top: 10px;
          }
        </style>
      </head>
      <body>
        <div class="header-box">
          <div class="title">${reportTitle}</div>
          <div class="subtitle">Информационно-аналитический комплекс ОАО "РЖД" • Система умного учета электроэнергии</div>
          <div class="period-subtitle">${reportSubtitle}</div>
        </div>

        <div class="summary-card">
          <strong style="color: #1e3a8a;">ОБЩИЙ ИТОГ ПО РЕГИОНУ:</strong><br/>
          • Потребление за выбранный период: <strong>${totalCurrent.toLocaleString()} кВт·ч</strong> (прошлый год: ${totalPrev.toLocaleString()} кВт·ч, 
          динамика: <span class="${totalDiff > 0 ? 'overrun' : 'saving'}">${totalDiff > 0 ? `+${totalDiff.toLocaleString()}` : totalDiff.toLocaleString()} кВт·ч (${totalPrev > 0 ? (totalDiff > 0 ? '+' : '') + totalPct.toFixed(2) : '0'}%)</span>).<br/>
          • Накопительный итог с начала года (YTD): <strong>${totalCurrentYtd.toLocaleString()} кВт·ч</strong> (прошлый год YTD: ${totalPrevYtd.toLocaleString()} кВт·ч, 
          динамика YTD: <span class="${totalYtdDiff > 0 ? 'overrun' : 'saving'}">${totalYtdDiff > 0 ? `+${totalYtdDiff.toLocaleString()}` : totalYtdDiff.toLocaleString()} кВт·ч (${totalPrevYtd > 0 ? (totalYtdDiff > 0 ? '+' : '') + totalYtdPct.toFixed(2) : '0'}%)</span>).<br/>
          • Общее состояние энергоэффективности: <strong style="color: ${totalDiff > 0 ? '#dc2626' : '#16a34a'}">${totalDiff > 0 ? 'ЗАФИКСИРОВАН ПЕРЕРАСХОД РЕСУРСОВ' : 'ДОСТИГНУТА СТАБИЛЬНАЯ ЭКОНОМИЯ'}</strong>.
        </div>

        <h2>РАЗДЕЛ 1. АНАЛИЗ ПО ТЕХНОЛОГИЧЕСКИМ КАТЕГОРИЯМ ПОТРЕБИТЕЛЕЙ</h2>
        <p>Ниже представлена сводная ведомость потребления электроэнергии в разрезе основных категорий технологических нужд за отчетный период и с начала года (YTD) в сравнении с прошлым годом:</p>

        <table>
          <thead>
            <tr>
              <th rowspan="2">Технологическая категория</th>
              <th rowspan="2" class="text-center">Кол-во ТП</th>
              <th colspan="3" class="text-center">За отчетный период (кВт·ч)</th>
              <th colspan="3" class="text-center">С начала года (YTD) (кВт·ч)</th>
            </tr>
            <tr>
              <th class="text-right">Текущий</th>
              <th class="text-right">Прошлый г.</th>
              <th class="text-right">Изменение (%)</th>
              <th class="text-right">Текущий YTD</th>
              <th class="text-right">Прошлый YTD</th>
              <th class="text-right">Изменение YTD (%)</th>
            </tr>
          </thead>
          <tbody>
            ${categoryAnalysis.map(c => {
              const diffPctText = c.prev > 0 ? `${c.diff > 0 ? '+' : ''}${c.pct.toFixed(1)}%` : '0%';
              const diffYtdPctText = c.prevYtd > 0 ? `${c.diffYtd > 0 ? '+' : ''}${c.pctYtd.toFixed(1)}%` : '0%';
              return `
                <tr>
                  <td class="bold">${c.category}</td>
                  <td class="text-center font-mono">${c.pointsCount}</td>
                  <td class="text-right font-mono">${c.curr.toLocaleString()}</td>
                  <td class="text-right font-mono text-gray">${c.prev.toLocaleString()}</td>
                  <td class="text-right font-mono ${c.diff > 0 ? 'overrun' : c.diff < 0 ? 'saving' : 'neutral'}">${c.diff > 0 ? `+${c.diff.toLocaleString()}` : c.diff.toLocaleString()} (${diffPctText})</td>
                  <td class="text-right font-mono">${c.currYtd.toLocaleString()}</td>
                  <td class="text-right font-mono text-gray">${c.prevYtd.toLocaleString()}</td>
                  <td class="text-right font-mono ${c.diffYtd > 0 ? 'overrun' : c.diffYtd < 0 ? 'saving' : 'neutral'}">${c.diffYtd > 0 ? `+${c.diffYtd.toLocaleString()}` : c.diffYtd.toLocaleString()} (${diffYtdPctText})</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>

        <h2>РАЗДЕЛ 2. АНАЛИЗ ПО СПОСОБАМ РАСЧЕТА РАСХОДА</h2>
        <p>Сопоставление потребления, рассчитанного по показаниям коммерческих приборов учета (прямое измерение) и по расчетному способу (по установленной мощности оборудования в отсутствие ПУ):</p>

        <table>
          <thead>
            <tr>
              <th rowspan="2">Способ расчета расхода</th>
              <th rowspan="2" class="text-center">Кол-во ТП</th>
              <th colspan="3" class="text-center">За отчетный период (кВт·ч)</th>
              <th colspan="3" class="text-center">С начала года (YTD) (кВт·ч)</th>
            </tr>
            <tr>
              <th class="text-right">Текущий</th>
              <th class="text-right">Прошлый г.</th>
              <th class="text-right">Изменение (%)</th>
              <th class="text-right">Текущий YTD</th>
              <th class="text-right">Прошлый YTD</th>
              <th class="text-right">Изменение YTD (%)</th>
            </tr>
          </thead>
          <tbody>
            ${methodAnalysis.map(m => {
              const diffPctText = m.prev > 0 ? `${m.diff > 0 ? '+' : ''}${m.pct.toFixed(1)}%` : '0%';
              const diffYtdPctText = m.prevYtd > 0 ? `${m.diffYtd > 0 ? '+' : ''}${m.pctYtd.toFixed(1)}%` : '0%';
              return `
                <tr>
                  <td class="bold">${m.name}</td>
                  <td class="text-center font-mono">${m.pointsCount}</td>
                  <td class="text-right font-mono">${m.curr.toLocaleString()}</td>
                  <td class="text-right font-mono text-gray">${m.prev.toLocaleString()}</td>
                  <td class="text-right font-mono ${m.diff > 0 ? 'overrun' : m.diff < 0 ? 'saving' : 'neutral'}">${m.diff > 0 ? `+${m.diff.toLocaleString()}` : m.diff.toLocaleString()} (${diffPctText})</td>
                  <td class="text-right font-mono">${m.currYtd.toLocaleString()}</td>
                  <td class="text-right font-mono text-gray">${m.prevYtd.toLocaleString()}</td>
                  <td class="text-right font-mono ${m.diffYtd > 0 ? 'overrun' : m.diffYtd < 0 ? 'saving' : 'neutral'}">${m.diffYtd > 0 ? `+${m.diffYtd.toLocaleString()}` : m.diffYtd.toLocaleString()} (${diffYtdPctText})</td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>

        <h2>РАЗДЕЛ 3. ДЕТАЛЬНЫЙ ЭНЕРГОАУДИТ В РАЗРЕЗЕ СТАНЦИЙ И ТОЧЕК ПОСТАВКИ</h2>
        <p>Подробный аудит каждой железнодорожной станции с распределением по технологическим категориям и отдельным точкам поставки с динамикой накопительного итога:</p>

        ${stationAnalysis.map(sa => {
          const stDiffPctText = sa.prev > 0 ? `${sa.diff > 0 ? '+' : ''}${sa.pct.toFixed(1)}%` : '0%';
          const stDiffYtdPctText = sa.prevYtd > 0 ? `${sa.diffYtd > 0 ? '+' : ''}${sa.pctYtd.toFixed(1)}%` : '0%';
          
          return `
            <div class="station-header">
              <strong style="font-size: 11pt; color: #1e3a8a;">ЖД СТАНЦИЯ: ${sa.station.name} (${sa.station.section})</strong><br/>
              • Расход за период: <strong>${sa.curr.toLocaleString()} кВт·ч</strong> (предыдущий год: ${sa.prev.toLocaleString()} кВт·ч, динамика: <span class="${sa.diff > 0 ? 'overrun' : 'saving'}">${sa.diff > 0 ? '+' : ''}${sa.diff.toLocaleString()} кВт·ч (${stDiffPctText})</span>)<br/>
              • Накопительно с начала года (YTD): <strong>${sa.currYtd.toLocaleString()} кВт·ч</strong> (прошлый YTD: ${sa.prevYtd.toLocaleString()} кВт·ч, динамика YTD: <span class="${sa.diffYtd > 0 ? 'overrun' : 'saving'}">${sa.diffYtd > 0 ? '+' : ''}${sa.diffYtd.toLocaleString()} кВт·ч (${stDiffYtdPctText})</span>)
            </div>

            <p style="font-size: 9.5pt; margin-bottom: 5px;"><strong>Свод по категориям на станции:</strong></p>
            <table style="margin-bottom: 12px; font-size: 8.5pt;">
              <thead>
                <tr style="background-color: #475569;">
                  <th>Технологическая категория</th>
                  <th class="text-right">Потребление за период (кВт·ч)</th>
                  <th class="text-right">Прошлый год (кВт·ч)</th>
                  <th class="text-right">Изменение (кВт·ч)</th>
                  <th class="text-right">Динамика (%)</th>
                </tr>
              </thead>
              <tbody>
                ${sa.categories.map(c => `
                  <tr>
                    <td>${c.category}</td>
                    <td class="text-right font-mono">${c.curr.toLocaleString()}</td>
                    <td class="text-right font-mono text-gray">${c.prev.toLocaleString()}</td>
                    <td class="text-right font-mono ${c.diff > 0 ? 'overrun' : 'saving'}">${c.diff > 0 ? '+' : ''}${c.diff.toLocaleString()}</td>
                    <td class="text-right font-mono ${c.diff > 0 ? 'overrun' : 'saving'}">${c.prev > 0 ? (c.diff > 0 ? '+' : '') + c.pct.toFixed(1) + '%' : '0%'}</td>
                  </tr>
                `).join("")}
              </tbody>
            </table>

            <p style="font-size: 9.5pt; margin-bottom: 5px;"><strong>Спецификация по точкам поставки (ТП) станции:</strong></p>
            <table>
              <thead>
                <tr style="background-color: #334155;">
                  <th>Точка поставки (ТП)</th>
                  <th>Категория</th>
                  <th>Способ учета</th>
                  <th class="text-right">Период (кВт·ч)</th>
                  <th class="text-right">Прошлый г.</th>
                  <th class="text-right">Начало года YTD</th>
                  <th class="text-right">Прошлый YTD</th>
                  <th class="text-right">Динамика YTD (кВт·ч)</th>
                  <th class="text-right">Откл. YTD</th>
                </tr>
              </thead>
              <tbody>
                ${sa.points.map(p => {
                  const pDiffText = p.diffYtd > 0 ? `+${p.diffYtd.toLocaleString()}` : p.diffYtd.toLocaleString();
                  const pPctText = p.prevYtd > 0 ? `${p.diffYtd > 0 ? '+' : ''}${p.pctYtd.toFixed(1)}%` : '0%';
                  const calcMethodLabel = p.point.calculationMethod === 'estimated' ? 'Расчетный' : 'Прибор учета';
                  const calcColor = p.point.calculationMethod === 'estimated' ? '#d97706' : '#2563eb';
                  
                  return `
                    <tr>
                      <td class="bold">${p.point.name}</td>
                      <td>${p.point.category}</td>
                      <td style="color: ${calcColor}; font-weight: bold; font-size: 8pt;">${calcMethodLabel}</td>
                      <td class="text-right font-mono">${p.curr.toLocaleString()}</td>
                      <td class="text-right font-mono text-gray">${p.prev.toLocaleString()}</td>
                      <td class="text-right font-mono">${p.currYtd.toLocaleString()}</td>
                      <td class="text-right font-mono text-gray">${p.prevYtd.toLocaleString()}</td>
                      <td class="text-right font-mono ${p.diffYtd > 0 ? 'overrun' : p.diffYtd < 0 ? 'saving' : 'neutral'}"><strong>${pDiffText}</strong></td>
                      <td class="text-right font-mono ${p.diffYtd > 0 ? 'overrun' : p.diffYtd < 0 ? 'saving' : 'neutral'}">${pPctText}</td>
                    </tr>
                  `;
                }).join("")}
              </tbody>
            </table>
            <div style="height: 10px;"></div>
          `;
        }).join("")}

        <div class="footer-box">
          Детальный консолидированный аналитический отчет сгенерирован автоматически.<br/>
          Дата формирования: ${new Date().toLocaleDateString("ru-RU")} | ЭЦП оператора: serj20742074@gmail.com | ОАО "РЖД" Информационные Технологии
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(["\ufeff" + htmlDocContent], {
      type: "application/msword;charset=utf-8"
    });

    const url = URL.createObjectURL(blob);
    const downloadLink = document.createElement("a");
    downloadLink.href = url;
    downloadLink.download = `Подробный_анализ_энергоэффективности_${selectedMonth.toString().padStart(2, '0')}_${selectedYear}.doc`;
    
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(url);
  };

  // --- EXPORT TO POWERPOINT (.PPTX) ---
  const exportReportToPptx = () => {
    const currentMonthLabel = selectedPeriodName;
    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    // Calculate totals & stats
    const totalCurrent = stations.reduce((sum, st) => {
      const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
      return sum + readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
    }, 0);
    const totalPrev = stations.reduce((sum, st) => {
      const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
      return sum + readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
    }, 0);
    const totalDiff = totalCurrent - totalPrev;
    const totalPct = totalPrev > 0 ? (totalDiff / totalPrev) * 100 : 0;

    // Slices for leaders
    const stationDiffs = stations.map(st => {
      const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
      const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
      const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
      const diff = curr - prev;
      const pct = prev > 0 ? (diff / prev) * 100 : 0;
      return { id: st.id, name: st.name, curr, prev, diff, pct };
    });

    const worstStations = stationDiffs.filter(s => s.diff > 0).sort((a, b) => b.diff - a.diff).slice(0, 5);
    const bestStations = stationDiffs.filter(s => s.diff < 0).sort((a, b) => a.diff - b.diff).slice(0, 5);

    // Categories
    const targetCategories = ["Освещение горловин", "Бытовые нагрузки", "Отопление"];
    const categoryInfo = targetCategories.map(cat => {
      const catPoints = supplyPoints.filter(p => p.category === cat && p.isActive);
      const catTotalCurr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && catPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
      const catTotalPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && catPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
      const catTotalDiff = catTotalCurr - catTotalPrev;
      const catTotalPct = catTotalPrev > 0 ? (catTotalDiff / catTotalPrev) * 100 : 0;
      return { category: cat, curr: catTotalCurr, prev: catTotalPrev, diff: catTotalDiff, pct: catTotalPct };
    });

    // Losses
    const totalLossCurrent = lossReadings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
    const totalLossPrev = lossReadings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
    const totalLossDiff = totalLossCurrent - totalLossPrev;
    const totalLossPct = totalLossPrev > 0 ? (totalLossDiff / totalLossPrev) * 100 : 0;

    // SLIDE 1: Title Slide
    let slide1 = pptx.addSlide();
    slide1.addText("", { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '0F172A' } }); // Dark background
    slide1.addText("ОАО РЖД • ЦЕНТР СБЕРЕЖЕНИЯ ЭНЕРГОРЕСУРСОВ", { x: 0.8, y: 1.5, w: 10.0, h: 0.4, fontSize: 13, bold: true, color: '3B82F6', fontFace: 'Arial' });
    slide1.addText("ЭНЕРГЕТИЧЕСКИЙ АУДИТ Ж/Д ИНФРАСТРУКТУРЫ", { x: 0.8, y: 2.0, w: 11.5, h: 1.2, fontSize: 26, bold: true, color: 'FFFFFF', fontFace: 'Arial' });
    slide1.addText(`Информационный доклад по итогам комплексного аудита\nОтчетный период: ${currentMonthLabel} ${selectedYear} г.`, { x: 0.8, y: 3.4, w: 10.0, h: 0.8, fontSize: 13, color: '94A3B8', fontFace: 'Arial' });
    slide1.addText("", { x: 0.8, y: 4.5, w: 5.0, h: 0.05, fill: { color: 'E2E8F0' } });
    slide1.addText("Сгенерировано автоматически на основе цифровой платформы", { x: 0.8, y: 4.8, w: 8.0, h: 0.4, fontSize: 11, italic: true, color: '64748B', fontFace: 'Arial' });

    // SLIDE 2: KPI Dashboard
    let slide2 = pptx.addSlide();
    slide2.addText("АНАЛИТИКА ЭНЕРГОПОТРЕБЛЕНИЯ УЧАСТКА", { x: 0.5, y: 0.4, w: 11.0, h: 0.4, fontSize: 18, bold: true, color: '1E3A8A' });
    slide2.addText(`Общие показатели энергопотребления за ${currentMonthLabel} ${selectedYear} г.`, { x: 0.5, y: 0.8, w: 11.0, h: 0.3, fontSize: 11, color: '475569' });

    slide2.addText(`Текущий расход:\n${totalCurrent.toLocaleString()} кВт·ч`, {
      x: 0.5, y: 1.4, w: 3.5, h: 1.5,
      fill: { color: 'EFF6FF' },
      border: { type: 'solid', color: 'BFDBFE', size: 1 },
      align: 'center', valign: 'middle',
      fontSize: 16, bold: true, color: '1E40AF', fontFace: 'Arial'
    } as any);

    slide2.addText(`Аналогичный период прошлого года:\n${totalPrev.toLocaleString()} кВт·ч`, {
      x: 4.3, y: 1.4, w: 3.5, h: 1.5,
      fill: { color: 'F8FAFC' },
      border: { type: 'solid', color: 'E2E8F0', size: 1 },
      align: 'center', valign: 'middle',
      fontSize: 14, color: '334155', fontFace: 'Arial'
    } as any);

    const diffColor = totalDiff > 0 ? 'B91C1C' : '047857';
    const diffBg = totalDiff > 0 ? 'FEF2F2' : 'ECFDF5';
    const diffBorder = totalDiff > 0 ? 'FCA5A5' : 'A7F3D0';
    const statusText = totalDiff > 0 ? 'ОБНАРУЖЕН СИСТЕМНЫЙ ПЕРЕРАСХОД' : 'СИСТЕМНАЯ ЭКОНОМИЯ';
    slide2.addText(`Изменение расхода (Динамика):\n${totalDiff > 0 ? `+${totalDiff.toLocaleString()}` : totalDiff.toLocaleString()} кВт·ч (${totalPct.toFixed(2)}%)\n\n[ ${statusText} ]`, {
      x: 8.1, y: 1.4, w: 4.7, h: 1.5,
      fill: { color: diffBg },
      border: { type: 'solid', color: diffBorder, size: 1 },
      align: 'center', valign: 'middle',
      fontSize: 13, bold: true, color: diffColor, fontFace: 'Arial'
    } as any);

    slide2.addText("Ключевые выводы аудита периода:", { x: 0.5, y: 3.2, w: 10.0, h: 0.3, fontSize: 13, bold: true, color: '1F2937' });
    let conclusions = [
      `Суммарный расход электроэнергии составил ${totalCurrent.toLocaleString()} кВт·ч, разница с прошлым годом: ${totalDiff > 0 ? '+' : ''}${totalDiff.toLocaleString()} кВт·ч.`,
      totalDiff > 0 
        ? "Превышение нормативов обусловлено ростом бытовых и отопительных нагрузок на ключевых станциях. Требуется внеплановый технический аудит регулирующих приборов."
        : "Положительная динамика обусловлена успешным завершением программы модернизации светодиодных осветительных систем и строгим контролем графиков дежурного обогрева.",
      "Доля энергии, учтенной поверенными приборами учета, превышает 85%. Остальной объем рассчитан по расчетным коэффициентам, что требует постепенного дооснащения."
    ];
    slide2.addText(conclusions.map(c => `• ${c}`).join('\n'), { x: 0.5, y: 3.6, w: 12.3, h: 1.6, fontSize: 11, color: '374151', lineSpacing: 20 });

    // SLIDE 3: Worst Performing Stations (Overruns, Max 5)
    let slide3 = pptx.addSlide();
    slide3.addText("ЛИДЕРЫ РОСТА РАСХОДА (ПЕРЕРАСХОД)", { x: 0.5, y: 0.4, w: 11.0, h: 0.4, fontSize: 18, bold: true, color: 'B91C1C' });
    slide3.addText("Станции с наибольшим превышением потребления относительно прошлого года (Worst 5, от худших к лучшим):", { x: 0.5, y: 0.8, w: 11.0, h: 0.3, fontSize: 11, color: '475569' });

    if (worstStations.length === 0) {
      slide3.addText("Превышений нормативов расхода не обнаружено!\nВсе подразделения работают в зоне стабильной экономии.", { x: 1.0, y: 2.0, w: 11.0, h: 1.5, align: 'center', fontSize: 16, bold: true, color: '047857' });
    } else {
      let worstTableRows: any[][] = [
        [
          { text: 'Железнодорожная станция', options: { fill: '991B1B', color: 'FFFFFF', bold: true } as any },
          { text: 'Текущий расход (кВт·ч)', options: { fill: '991B1B', color: 'FFFFFF', bold: true, align: 'right' } as any },
          { text: 'Прошлый год (кВт·ч)', options: { fill: '991B1B', color: 'FFFFFF', bold: true, align: 'right' } as any },
          { text: 'Перерасход (кВт·ч)', options: { fill: '991B1B', color: 'FFFFFF', bold: true, align: 'right' } as any },
          { text: 'Отклонение (%)', options: { fill: '991B1B', color: 'FFFFFF', bold: true, align: 'right' } as any }
        ]
      ];
      worstStations.forEach(s => {
        worstTableRows.push([
          { text: s.name, options: { bold: true } as any },
          { text: s.curr.toLocaleString(), options: { align: 'right' } as any },
          { text: s.prev.toLocaleString(), options: { align: 'right' } as any },
          { text: `+${(s.diff).toLocaleString()}`, options: { align: 'right', color: 'B91C1C', bold: true } as any },
          { text: `+${s.pct.toFixed(2)}%`, options: { align: 'right', color: 'B91C1C', bold: true } as any }
        ]);
      });

      slide3.addTable(worstTableRows, {
        x: 0.5, y: 1.3, w: 12.3, h: 2.0,
        fontSize: 10.5, fontFace: 'Arial',
        colW: [4.0, 2.0, 2.0, 2.2, 2.1],
        border: { type: 'solid', color: 'F3F4F6', size: 1 }
      } as any);

      slide3.addText("Предписываемые корректирующие действия:", { x: 0.5, y: 3.8, w: 10.0, h: 0.3, fontSize: 12, bold: true, color: '111827' });
      let worstRecs = [
        "Проверить исправность систем коммерческого учета на вводных фидерах указанных станций.",
        "Выявить скрытые утечки мощности во внерабочее дежурное время.",
        "Ограничить использование бытовых приборов высокой мощности в дежурных помещениях."
      ];
      slide3.addText(worstRecs.map(r => `• ${r}`).join('\n'), { x: 0.5, y: 4.1, w: 12.3, h: 1.0, fontSize: 10.5, color: '4B5563', lineSpacing: 18 });
    }

    // SLIDE 4: Best Performing Stations (Savings, Max 5)
    let slide4 = pptx.addSlide();
    slide4.addText("ЛИДЕРЫ СНИЖЕНИЯ РАСХОДА (ЭКОНОМИЯ)", { x: 0.5, y: 0.4, w: 11.0, h: 0.4, fontSize: 18, bold: true, color: '047857' });
    slide4.addText("Станции, достигшие наибольшего сокращения электропотребления относительно прошлого года (Best 5, от лучших к худшим):", { x: 0.5, y: 0.8, w: 11.0, h: 0.3, fontSize: 11, color: '475569' });

    if (bestStations.length === 0) {
      slide4.addText("За отчетный период ни одна станция не зафиксировала чистую экономию электроэнергии.", { x: 1.0, y: 2.0, w: 11.0, h: 1.5, align: 'center', fontSize: 14, italic: true, color: '64748B' });
    } else {
      let bestTableRows: any[][] = [
        [
          { text: 'Железнодорожная станция', options: { fill: '065F46', color: 'FFFFFF', bold: true } as any },
          { text: 'Текущий расход (кВт·ч)', options: { fill: '065F46', color: 'FFFFFF', bold: true, align: 'right' } as any },
          { text: 'Прошлый год (кВт·ч)', options: { fill: '065F46', color: 'FFFFFF', bold: true, align: 'right' } as any },
          { text: 'Снижение (кВт·ч)', options: { fill: '065F46', color: 'FFFFFF', bold: true, align: 'right' } as any },
          { text: 'Экономия (%)', options: { fill: '065F46', color: 'FFFFFF', bold: true, align: 'right' } as any }
        ]
      ];
      bestStations.forEach(s => {
        bestTableRows.push([
          { text: s.name, options: { bold: true } as any },
          { text: s.curr.toLocaleString(), options: { align: 'right' } as any },
          { text: s.prev.toLocaleString(), options: { align: 'right' } as any },
          { text: (s.diff).toLocaleString(), options: { align: 'right', color: '047857', bold: true } as any },
          { text: `${s.pct.toFixed(2)}%`, options: { align: 'right', color: '047857', bold: true } as any }
        ]);
      });

      slide4.addTable(bestTableRows, {
        x: 0.5, y: 1.3, w: 12.3, h: 2.0,
        fontSize: 10.5, fontFace: 'Arial',
        colW: [4.0, 2.0, 2.0, 2.2, 2.1],
        border: { type: 'solid', color: 'F3F4F6', size: 1 }
      } as any);

      slide4.addText("Положительный опыт к тиражированию:", { x: 0.5, y: 3.8, w: 10.0, h: 0.3, fontSize: 12, bold: true, color: '111827' });
      let bestRecs = [
        "Поощрить ответственных лиц за строгое соблюдение регламентов энергоэффективности на станциях-лидерах.",
        "Изучить и внедрить схемы автоматического управления наружным освещением на данных объектах.",
        "Масштабировать технологии оптимизации режимов обогрева на другие участки железной дороги."
      ];
      slide4.addText(bestRecs.map(r => `• ${r}`).join('\n'), { x: 0.5, y: 4.1, w: 12.3, h: 1.0, fontSize: 10.5, color: '4B5563', lineSpacing: 18 });
    }

    // SLIDE 5: Segment Analysis
    let slide5 = pptx.addSlide();
    slide5.addText("СЕГМЕНТНЫЙ АНАЛИЗ ГРУПП НАГРУЗОК", { x: 0.5, y: 0.4, w: 11.0, h: 0.4, fontSize: 18, bold: true, color: '1E3A8A' });
    slide5.addText("Результаты сопоставительного анализа ключевых технологических групп энергопотребителей участка:", { x: 0.5, y: 0.8, w: 11.0, h: 0.3, fontSize: 11, color: '475569' });

    let catTableRows: any[][] = [
      [
        { text: 'Категория потребителей', options: { fill: '1E3A8A', color: 'FFFFFF', bold: true } as any },
        { text: 'Текущий расход (кВт·ч)', options: { fill: '1E3A8A', color: 'FFFFFF', bold: true, align: 'right' } as any },
        { text: 'Прошлый год (кВт·ч)', options: { fill: '1E3A8A', color: 'FFFFFF', bold: true, align: 'right' } as any },
        { text: 'Изменение (кВт·ч)', options: { fill: '1E3A8A', color: 'FFFFFF', bold: true, align: 'right' } as any },
        { text: 'Динамика (%)', options: { fill: '1E3A8A', color: 'FFFFFF', bold: true, align: 'right' } as any }
      ]
    ];
    categoryInfo.forEach(c => {
      const isO = c.diff > 0;
      catTableRows.push([
        { text: c.category, options: { bold: true } as any },
        { text: c.curr.toLocaleString(), options: { align: 'right' } as any },
        { text: c.prev.toLocaleString(), options: { align: 'right' } as any },
        { text: `${isO ? '+' : ''}${c.diff.toLocaleString()}`, options: { align: 'right', color: isO ? 'B91C1C' : '047857', bold: true } as any },
        { text: `${isO ? '+' : ''}${c.pct.toFixed(2)}%`, options: { align: 'right', color: isO ? 'B91C1C' : '047857', bold: true } as any }
      ]);
    });

    slide5.addTable(catTableRows, {
      x: 0.5, y: 1.4, w: 12.3, h: 2.0,
      fontSize: 11, fontFace: 'Arial',
      colW: [4.0, 2.0, 2.0, 2.2, 2.1],
      border: { type: 'solid', color: 'F3F4F6', size: 1 }
    } as any);

    slide5.addText("Технологические выводы:", { x: 0.5, y: 3.8, w: 10.0, h: 0.3, fontSize: 12, bold: true, color: '111827' });
    let catInsights = [
      "Категория Освещения горловин стабильна благодаря внедрению автоматических астрономических реле.",
      "Отопление и обогрев демонстрируют высокую температурную чувствительность. Рекомендуется установка локальных термостатов.",
      "Хозяйственно-бытовые нагрузки требуют регулярных плановых проверок на предмет использования неэффективных нагревательных приборов."
    ];
    slide5.addText(catInsights.map(i => `• ${i}`).join('\n'), { x: 0.5, y: 4.1, w: 12.3, h: 1.0, fontSize: 10.5, color: '4B5563', lineSpacing: 18 });

    // SLIDE 6: Network Loss Analysis (Потери)
    let slide6 = pptx.addSlide();
    slide6.addText("АНАЛИЗ ТЕХНОЛОГИЧЕСКИХ ПОТЕРЬ В СЕТИ", { x: 0.5, y: 0.4, w: 11.0, h: 0.4, fontSize: 18, bold: true, color: '334155' });
    slide6.addText("Оценка величины и динамики потерь в ЛЭП и понижающих трансформаторах за отчетный период:", { x: 0.5, y: 0.8, w: 11.0, h: 0.3, fontSize: 11, color: '475569' });

    const lossIsO = totalLossDiff > 0;
    slide6.addText(`Общие сетевые потери:\n${totalLossCurrent.toLocaleString()} кВт·ч`, {
      x: 0.5, y: 1.3, w: 4.0, h: 1.2,
      fill: { color: 'F1F5F9' },
      align: 'center', valign: 'middle',
      fontSize: 14, bold: true, color: '334155', fontFace: 'Arial'
    } as any);
    slide6.addText(`Динамика к прошлому году:\n${lossIsO ? '+' : ''}${totalLossDiff.toLocaleString()} кВт·ч (${lossIsO ? '+' : ''}${totalLossPct.toFixed(2)}%)\n[ ${lossIsO ? 'РОСТ ТЕХНИЧЕСКИХ ПОТЕРЬ' : 'СНИЖЕНИЕ ПОТЕРЬ'} ]`, {
      x: 4.8, y: 1.3, w: 8.0, h: 1.2,
      fill: { color: lossIsO ? 'FEF2F2' : 'ECFDF5' },
      align: 'center', valign: 'middle',
      fontSize: 12, bold: true, color: lossIsO ? 'B91C1C' : '047857', fontFace: 'Arial'
    } as any);

    let lossTableRows: any[][] = [
      [
        { text: 'Объект потерь (фидер / КТП)', options: { fill: '475569', color: 'FFFFFF', bold: true } as any },
        { text: 'Закрепленная станция', options: { fill: '475569', color: 'FFFFFF', bold: true } as any },
        { text: 'Тип', options: { fill: '475569', color: 'FFFFFF', bold: true } as any },
        { text: 'Текущие потери (кВт·ч)', options: { fill: '475569', color: 'FFFFFF', bold: true, align: 'right' } as any },
        { text: 'Изменение (кВт·ч)', options: { fill: '475569', color: 'FFFFFF', bold: true, align: 'right' } as any }
      ]
    ];
    lossObjects.slice(0, 4).forEach(lo => {
      const currL = lossReadings.filter(r => r.lossObjectId === lo.id && r.year === selectedYear && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
      const prevL = lossReadings.filter(r => r.lossObjectId === lo.id && r.year === selectedYear - 1 && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
      const diffL = currL - prevL;
      const st = stations.find(s => s.id === lo.stationId);
      const loO = diffL > 0;
      lossTableRows.push([
        { text: lo.name, options: { bold: true } as any },
        { text: st ? st.name : "Вне станций", options: {} as any },
        { text: (lo.name.toLowerCase().includes('ктп') || lo.name.toLowerCase().includes('трансф')) ? 'Трансф.' : 'ЛЭП', options: {} as any },
        { text: currL.toLocaleString(), options: { align: 'right' } as any },
        { text: `${loO ? '+' : ''}${diffL.toLocaleString()}`, options: { align: 'right', color: loO ? 'B91C1C' : '047857', bold: true } as any }
      ]);
    });

    slide6.addTable(lossTableRows, {
      x: 0.5, y: 2.8, w: 12.3, h: 1.5,
      fontSize: 10, fontFace: 'Arial',
      colW: [4.0, 2.5, 2.0, 2.0, 2.1],
      border: { type: 'solid', color: 'F1F5F9', size: 1 }
    } as any);

    slide6.addText("Рекомендация: Провести локальный тепловизионный аудит и инструментальные замеры на КТП с растущим уровнем потерь.", { x: 0.5, y: 4.6, w: 12.3, h: 0.3, fontSize: 10.5, italic: true, color: '475569' });

    // SLIDE 7: Closing / Technical Recommendations
    let slide7 = pptx.addSlide();
    slide7.addText("", { x: 0, y: 0, w: '100%', h: '100%', fill: { color: '1E3A8A' } }); // Solid Navy Closing
    slide7.addText("ЭНЕРГОСБЕРЕЖЕНИЕ И КОРПОРАТИВНЫЕ СТАНДАРТЫ", { x: 1.0, y: 1.2, w: 10.0, h: 0.4, fontSize: 14, bold: true, color: '93C5FD', fontFace: 'Arial' });
    slide7.addText("СВОДНЫЙ КОРПОРАТИВНЫЙ РЕГЛАМЕНТ ДЕЙСТВИЙ", { x: 1.0, y: 1.7, w: 11.0, h: 0.8, fontSize: 24, bold: true, color: 'FFFFFF', fontFace: 'Arial' });
    
    let endRecs = [
      "1. Обеспечить 100% охват автоматизированных приборов учета во всех контрольных точках.",
      "2. Провести регулировку астрономических реле и суточных таймеров прожекторного освещения.",
      "3. Ограничить отопительные лимиты в необитаемых технических блок-контейнерах до +14°C.",
      "4. Организовать регулярную тепловизионную диагностику контактных соединений КТП.",
      "5. Своевременно производить замену и государственную поверку приборов коммерческого учета."
    ];
    slide7.addText(endRecs.join('\n\n'), { x: 1.0, y: 2.7, w: 11.0, h: 2.5, fontSize: 11.5, color: 'E0F2FE', fontFace: 'Arial', lineSpacing: 10 });
    
    slide7.addText(`Доклад подготовлен: ГИС-аналитика ОАО РЖД | ${new Date().toLocaleDateString("ru-RU")} в ${new Date().toLocaleTimeString("ru-RU")}`, { x: 1.0, y: 5.6, w: 10.0, h: 0.3, fontSize: 10, italic: true, color: '93C5FD', fontFace: 'Arial' });

    const fileName = `Презентация_аудита_${selectedMonth.toString().padStart(2, '0')}_${selectedYear}.pptx`;
    pptx.writeFile({ fileName });
  };

  // --- ACTIONS (CRUD) ---
  const handleSaveStation = (e: React.FormEvent) => {
    e.preventDefault();
    if (!stationModal) return;

    try {
      const nameVal = (stationModal.name || '').trim();
      const sectionVal = (stationModal.section || '').trim();
      const noteVal = (stationModal.note || '').trim();

      if (stationModal.mode === 'add') {
        const newId = `st-${Date.now()}`;
        const updated = [...stations, { id: newId, name: nameVal, section: sectionVal, note: noteVal }];
        syncDatabase(updated, supplyPoints, categories, readings, lossObjects, lossReadings);
        setSelectedStationId(newId);
      } else {
        const updated = stations.map(s => s.id === stationModal.id ? { ...s, name: nameVal, section: sectionVal, note: noteVal } : s);
        syncDatabase(updated, supplyPoints, categories, readings, lossObjects, lossReadings);
      }
    } catch (err) {
      console.error("Error saving station", err);
    } finally {
      setStationModal(null);
    }
  };

  const handleDeleteStation = (id: string) => {
    const station = stations.find(s => s.id === id);
    setDeleteConfirm({
      isOpen: true,
      title: "Удалить станцию?",
      message: `Вы действительно хотите удалить железнодорожную станцию "${station?.name || ''}"? Это также приведет к удалению связанных точек поставки (ТП) и данных о расходах.`,
      onConfirm: () => {
        const updatedSt = stations.filter(s => s.id !== id);
        const updatedSps = supplyPoints.filter(p => p.stationId !== id);
        syncDatabase(updatedSt, updatedSps, categories, readings, lossObjects, lossReadings);
        if (selectedStationId === id) {
          setSelectedStationId(updatedSt.length > 0 ? updatedSt[0].id : '');
        }
        setDeleteConfirm(null);
      }
    });
  };

  const handleSavePoint = (e: React.FormEvent) => {
    e.preventDefault();
    if (!pointModal) return;

    try {
      const nameVal = (pointModal.name || '').trim();
      const noteVal = (pointModal.note || '').trim();
      const currK = pointModal.currKwh !== undefined ? Number(pointModal.currKwh) : 0;
      const prevK = pointModal.prevKwh !== undefined ? Number(pointModal.prevKwh) : 0;

      if (pointModal.mode === 'add') {
        const newId = `tp-${Date.now()}`;
        const newPt: SupplyPoint = {
          id: newId,
          name: nameVal,
          stationId: pointModal.stationId,
          category: pointModal.category,
          note: noteVal,
          isActive: pointModal.isActive,
          calculationMethod: pointModal.calculationMethod || 'meter'
        };
        const updatedSps = [...supplyPoints, newPt];
        const updatedRds = [
          ...readings, 
          { supplyPointId: newId, year: selectedYear, month: selectedMonth, value: currK },
          { supplyPointId: newId, year: selectedYear - 1, month: selectedMonth, value: prevK }
        ];
        syncDatabase(stations, updatedSps, categories, updatedRds, lossObjects, lossReadings);
      } else {
        const updatedSps = supplyPoints.map(p => p.id === pointModal.id ? {
          ...p,
          name: nameVal,
          stationId: pointModal.stationId,
          category: pointModal.category,
          note: noteVal,
          isActive: pointModal.isActive,
          calculationMethod: pointModal.calculationMethod || 'meter'
        } : p);

        let updatedRds = [...readings];
        
        // update or insert for selectedYear
        const currIdx = updatedRds.findIndex(r => r.supplyPointId === pointModal.id && r.year === selectedYear && r.month === selectedMonth);
        if (currIdx > -1) {
          updatedRds[currIdx] = { ...updatedRds[currIdx], value: currK };
        } else {
          updatedRds.push({ supplyPointId: pointModal.id!, year: selectedYear, month: selectedMonth, value: currK });
        }

        // update or insert for selectedYear - 1
        const prevIdx = updatedRds.findIndex(r => r.supplyPointId === pointModal.id && r.year === selectedYear - 1 && r.month === selectedMonth);
        if (prevIdx > -1) {
          updatedRds[prevIdx] = { ...updatedRds[prevIdx], value: prevK };
        } else {
          updatedRds.push({ supplyPointId: pointModal.id!, year: selectedYear - 1, month: selectedMonth, value: prevK });
        }

        syncDatabase(stations, updatedSps, categories, updatedRds, lossObjects, lossReadings);
      }
    } catch (err) {
      console.error("Error saving supply point", err);
    } finally {
      setPointModal(null);
    }
  };

  const handleDeletePoint = (id: string) => {
    const point = supplyPoints.find(p => p.id === id);
    setDeleteConfirm({
      isOpen: true,
      title: "Удалить точку поставки?",
      message: `Вы действительно хотите удалить точку поставки (ТП) "${point?.name || ''}"? Это также удалит все зафиксированные показания расходов для неё.`,
      onConfirm: () => {
        const updatedSps = supplyPoints.filter(p => p.id !== id);
        const updatedRds = readings.filter(r => r.supplyPointId !== id);
        syncDatabase(stations, updatedSps, categories, updatedRds, lossObjects, lossReadings);
        setDeleteConfirm(null);
      }
    });
  };

  const handleEditPointFromDetailed = (point: SupplyPoint) => {
    const currVal = readings.find(r => r.supplyPointId === point.id && r.year === selectedYear && r.month === selectedMonth)?.value || 0;
    const prevVal = readings.find(r => r.supplyPointId === point.id && r.year === (selectedYear - 1) && r.month === selectedMonth)?.value || 0;
    setPointModal({
      isOpen: true,
      mode: 'edit',
      id: point.id,
      name: point.name,
      stationId: point.stationId,
      category: point.category,
      note: point.note || '',
      isActive: point.isActive,
      calculationMethod: point.calculationMethod || 'meter',
      currKwh: currVal,
      prevKwh: prevVal
    });
  };

  const handleToggleArchivePoint = (id: string) => {
    const updatedSps = supplyPoints.map(p => p.id === id ? { ...p, isActive: !p.isActive } : p);
    syncDatabase(stations, updatedSps, categories, readings, lossObjects, lossReadings);
  };

  const updatePointValueInline = (spId: string, valStr: string, year: number = selectedYear) => {
    const val = Number(valStr.replace(/\D/g, '')) || 0;
    const exists = readings.find(r => r.supplyPointId === spId && r.year === year && r.month === selectedMonth);
    let updatedRds: Reading[];
    if (exists) {
      updatedRds = readings.map(r => (r.supplyPointId === spId && r.year === year && r.month === selectedMonth) ? { ...r, value: val } : r);
    } else {
      updatedRds = [...readings, { supplyPointId: spId, year: year, month: selectedMonth, value: val }];
    }
    syncDatabase(stations, supplyPoints, categories, updatedRds, lossObjects, lossReadings);
  };

  const updatePointNoteInline = (spId: string, noteStr: string) => {
    const updatedSps = supplyPoints.map(p => p.id === spId ? { ...p, note: noteStr } : p);
    syncDatabase(stations, updatedSps, categories, readings, lossObjects, lossReadings);
  };

  const updatePointCalculationMethodInline = (spId: string, method: 'meter' | 'estimated') => {
    const updatedSps = supplyPoints.map(p => p.id === spId ? { ...p, calculationMethod: method } : p);
    syncDatabase(stations, updatedSps, categories, readings, lossObjects, lossReadings);
  };

  const updateLossValueInline = (loId: string, valStr: string) => {
    const val = Number(valStr.replace(/\D/g, '')) || 0;
    const exists = lossReadings.find(r => r.lossObjectId === loId && r.year === selectedYear && r.month === selectedMonth);
    let updatedLrs: LossReading[];
    if (exists) {
      updatedLrs = lossReadings.map(r => (r.lossObjectId === loId && r.year === selectedYear && r.month === selectedMonth) ? { ...r, value: val } : r);
    } else {
      updatedLrs = [...lossReadings, { lossObjectId: loId, year: selectedYear, month: selectedMonth, value: val }];
    }
    syncDatabase(stations, supplyPoints, categories, readings, lossObjects, updatedLrs);
  };

  const handleSaveLossObject = (e: React.FormEvent) => {
    e.preventDefault();
    if (!lossModal) return;

    try {
      const nameVal = (lossModal.name || '').trim();
      const sectionVal = (lossModal.section || '').trim();
      const noteVal = (lossModal.note || '').trim();

      if (lossModal.mode === 'add') {
        const newId = `loss-${Date.now()}`;
        const newLo: LossObject = {
          id: newId,
          name: nameVal,
          stationId: lossModal.stationId || undefined,
          section: sectionVal || undefined,
          note: noteVal
        };
        const updatedLo = [...lossObjects, newLo];
        const updatedLrs = [...lossReadings, { lossObjectId: newId, year: selectedYear, month: selectedMonth, value: 0 }];
        syncDatabase(stations, supplyPoints, categories, readings, updatedLo, updatedLrs);
        setSelectedLossId(newId);
      } else {
        const updatedLo = lossObjects.map(lo => lo.id === lossModal.id ? {
          ...lo,
          name: nameVal,
          stationId: lossModal.stationId || undefined,
          section: sectionVal || undefined,
          note: noteVal
        } : lo);
        syncDatabase(stations, supplyPoints, categories, readings, updatedLo, lossReadings);
      }
    } catch (err) {
      console.error("Error saving loss object", err);
    } finally {
      setLossModal(null);
    }
  };

  const handleDeleteLossObject = (id: string) => {
    const lo = lossObjects.find(l => l.id === id);
    setDeleteConfirm({
      isOpen: true,
      title: "Удалить объект потерь?",
      message: `Вы действительно хотите удалить объект технологических потерь "${lo?.name || ''}"? Это также сотрет связанные показания этих потерь.`,
      onConfirm: () => {
        const updatedLo = lossObjects.filter(lo => lo.id !== id);
        const updatedLrs = lossReadings.filter(r => r.lossObjectId !== id);
        syncDatabase(stations, supplyPoints, categories, readings, updatedLo, updatedLrs);
        if (selectedLossId === id && updatedLo.length > 0) {
          setSelectedLossId(updatedLo[0].id);
        }
        setDeleteConfirm(null);
      }
    });
  };

  const handleAddCategory = () => {
    const trimmed = categoryInput.trim();
    if (!trimmed || categories.includes(trimmed)) return;
    const updated = [...categories, trimmed];
    syncDatabase(stations, supplyPoints, updated, readings, lossObjects, lossReadings);
    setCategoryInput('');
  };

  const handleDeleteCategory = (catToDelete: string) => {
    setDeleteConfirm({
      isOpen: true,
      title: "Удалить категорию?",
      message: `Вы действительно хотите удалить категорию нагрузки "${catToDelete}"? Точки поставки из этой категории будут перенесены в другую категорию.`,
      onConfirm: () => {
        const updated = categories.filter(c => c !== catToDelete);
        const safeCats = updated.length > 0 ? updated : ['Прочие'];
        const updatedSps = supplyPoints.map(sp => sp.category === catToDelete ? { ...sp, category: safeCats[0] } : sp);
        syncDatabase(stations, updatedSps, safeCats, readings, lossObjects, lossReadings);
      }
    });
  };

  // --- SVG MULTI-YEAR CHART ---
  const CustomSvgTrendChart = ({ targetId, type }: { targetId: string; type: 'station' | 'supply_point' | 'loss' }) => {
    const activeYearsList = [2025, 2026];
    
    const chartData = useMemo(() => {
      const store: { [year: number]: number[] } = { 2025: Array(12).fill(0), 2026: Array(12).fill(0) };

      activeYearsList.forEach(yr => {
        for (let m = 1; m <= 12; m++) {
          if (type === 'station') {
            const spsIds = supplyPoints.filter(p => p.stationId === targetId && p.isActive).map(p => p.id);
            store[yr][m - 1] = readings
              .filter(rd => rd.supplyPointId && rd.year === yr && rd.month === m && spsIds.includes(rd.supplyPointId))
              .reduce((sum, r) => sum + r.value, 0);
          } else if (type === 'supply_point') {
            const r = readings.find(rd => rd.supplyPointId === targetId && rd.year === yr && rd.month === m);
            store[yr][m - 1] = r ? r.value : 0;
          } else {
            const r = lossReadings.find(ld => ld.lossObjectId === targetId && ld.year === yr && ld.month === m);
            store[yr][m - 1] = r ? r.value : 0;
          }
        }
      });

      return store;
    }, [targetId, type, readings, lossReadings, supplyPoints]);

    const maxLimit = useMemo(() => {
      let high = 800;
      activeYearsList.forEach(yr => {
        if (!visibleYears[yr]) return;
        const m = Math.max(...chartData[yr]);
        if (m > high) high = m;
      });
      return Math.ceil(high * 1.15);
    }, [chartData, visibleYears]);

    const pad = 45;
    const h = 200;
    const w = 500;
    const innerH = h - pad * 2;
    const innerW = w - pad * 2;

    const xOf = (mIdx: number) => pad + (mIdx / 11) * innerW;
    const yOf = (v: number) => h - pad - (v / maxLimit) * innerH;

    return (
      <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-[#1e293b]/70 border-slate-700' : 'bg-white border-slate-250'} shadow-sm`}>
        <div className="flex items-center justify-between gap-4 mb-3">
          <div className="flex items-center gap-3">
            <span className="text-xs text-slate-400 font-semibold uppercase">Сравнить года:</span>
            {activeYearsList.map(yr => (
              <label key={yr} className="inline-flex items-center gap-1 cursor-pointer text-xs">
                <input
                  type="checkbox"
                  checked={visibleYears[yr] || false}
                  onChange={() => setVisibleYears(prev => ({ ...prev, [yr]: !prev[yr] }))}
                  className="rounded accent-blue-600 focus:ring-0"
                />
                <span className={yr === 2026 ? "text-amber-500 font-bold" : "text-sky-400 font-bold"}>{yr} г.</span>
              </label>
            ))}
          </div>
          <span className="text-[10px] text-slate-400">кВт·ч</span>
        </div>

        <div className="relative">
          <svg viewBox={`0 0 ${w} ${h}`} className="w-full h-auto select-none">
            {/* Grid */}
            {[0, 0.25, 0.5, 0.75, 1].map((r, i) => {
              const val = r * maxLimit;
              const yPos = yOf(val);
              return (
                <g key={i}>
                  <line x1={pad} y1={yPos} x2={w - pad} y2={yPos} stroke={darkTheme ? "#334155" : "#e2e8f0"} strokeDasharray="2 3" />
                  <text x={pad - 6} y={yPos + 3} textAnchor="end" fill={darkTheme ? "#64748b" : "#94a3b8"} className="text-[9px] font-mono">
                    {Math.round(val).toLocaleString()}
                  </text>
                </g>
              );
            })}

            {/* Months indicators */}
            {RUSSIAN_MONTHS_SHORT.map((item, idx) => {
              const xPos = xOf(idx);
              return (
                <text key={idx} x={xPos} y={h - pad + 14} textAnchor="middle" fill="#64748b" className="text-[8px] font-semibold">
                  {item}
                </text>
              );
            })}

            {/* Path 2025 */}
            {visibleYears[2025] && (
              <g>
                <path
                  d={chartData[2025].map((val, idx) => `${idx === 0 ? 'M' : 'L'} ${xOf(idx)} ${yOf(val)}`).join(' ')}
                  fill="none"
                  stroke="#38bdf8"
                  strokeWidth="2"
                  strokeLinecap="round"
                />
                {chartData[2025].map((val, idx) => (
                  <circle
                    key={idx}
                    cx={xOf(idx)}
                    cy={yOf(val)}
                    r="3.5"
                    fill={darkTheme ? "#0f172a" : "#ffffff"}
                    stroke="#38bdf8"
                    strokeWidth="1.5"
                    className="cursor-pointer"
                    onMouseEnter={() => setHoveredPoint({ month: idx + 1, year: 2025, value: val })}
                    onMouseLeave={() => setHoveredPoint(null)}
                  />
                ))}
              </g>
            )}

            {/* Path 2026 */}
            {visibleYears[2026] && (
              <g>
                <path
                  d={chartData[2026].filter((_, idx) => !(2026 === selectedYear && idx + 1 > 6)).map((val, idx) => `${idx === 0 ? 'M' : 'L'} ${xOf(idx)} ${yOf(val)}`).join(' ')}
                  fill="none"
                  stroke="#f59e0b"
                  strokeWidth="2.5"
                  strokeLinecap="round"
                />
                {chartData[2026].filter((_, idx) => !(2026 === selectedYear && idx + 1 > 6)).map((val, idx) => (
                  <circle
                    key={idx}
                    cx={xOf(idx)}
                    cy={yOf(val)}
                    r="4"
                    fill={darkTheme ? "#0f172a" : "#ffffff"}
                    stroke="#f59e0b"
                    strokeWidth="2"
                    className="cursor-pointer"
                    onMouseEnter={() => setHoveredPoint({ month: idx + 1, year: 2026, value: val })}
                    onMouseLeave={() => setHoveredPoint(null)}
                  />
                ))}
              </g>
            )}
          </svg>

          {/* Interactive Hover Tooltip */}
          {hoveredPoint && (
            <div className={`absolute top-1 left-1/2 -translate-x-1/2 px-2.5 py-1 text-[10px] rounded border text-center z-25 flex items-center gap-1 ${
              darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-slate-50 border-slate-200 text-slate-800'
            }`}>
              <Calendar className="w-3 h-3 text-blue-500" />
              <span>{RUSSIAN_MONTHS[hoveredPoint.month - 1]} {hoveredPoint.year}: <strong>{hoveredPoint.value.toLocaleString()}</strong> кВт·ч</span>
            </div>
          )}
        </div>
      </div>
    );
  };

  // --- EXCEL IMPORTER HANDLERS ---
  const triggerExcelSelect = () => {
    document.getElementById('hidden-excel-picker')?.click();
  };

  const handleExcelImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const u8 = new Uint8Array(evt.target?.result as ArrayBuffer);
        const wb = XLSX.read(u8, { type: 'array' });

        if (wb.SheetNames.length === 0) {
          alert("Данные в файле не обнаружены!");
          return;
        }

        let pointsState = [...supplyPoints];
        let readingsState = [...readings];
        let stationsState = [...stations];
        let lossObjectsState = [...lossObjects];
        let lossReadingsState = [...lossReadings];

        let totalStandardImportCount = 0;
        let totalRjdImportCount = 0;
        let detectedYearsGlobal: number[] = [];
        let rjdSheetsProcessedCount = 0;
        let originalSheetsProcessedCount = 0;

        // --- Custom Robust Matcher for existing points ---
        const normalizeForComp = (s: string) => {
          if (!s) return "";
          let res = s.toLowerCase()
            .replace(/ё/g, 'е')
            .replace(/[\u00A0\u1680\u180e\u2000-\u200a\u2028\u2029\u202f\u205f\u3000]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
          
          // Map Latin homoglyphs to Cyrillic
          const map: { [key: string]: string } = {
            'a': 'а', 'c': 'с', 'e': 'е', 'o': 'о', 'p': 'р', 'x': 'х', 'y': 'у',
            't': 'т', 'h': 'н', 'm': 'м', 'b': 'в', 'k': 'к'
          };
          return res.split('').map(char => map[char] || char).join('');
        };

        const cleanCompareStr = (sStr: string) => {
          return normalizeForComp(sStr)
            .replace(/ст\.|станция|постэц|пост|эц|фидер|тп|№|э\/отопление|эотопление|отопление|освещение|осв/gi, '')
            .replace(/[^a-zа-я0-9]/gi, '')
            .trim();
        };

        const findMatchingStation = (nameInFile: string, stationsList: Station[]) => {
          if (!nameInFile) return null;
          const nameNorm = normalizeForComp(nameInFile);
          
          // 1. Exact or contains check
          const directMatch = stationsList.find(s => {
            const sNameNorm = normalizeForComp(s.name);
            return sNameNorm.includes(nameNorm) || nameNorm.includes(sNameNorm);
          });
          if (directMatch) return directMatch;

          // 2. Fuzzy match by significant words (longer than 3 chars, ignoring common prefixes)
          const cleanWords = nameInFile
            .replace(/станция|ст\./gi, '')
            .toLowerCase()
            .split(/[^a-zа-я0-9]+/i)
            .filter(w => w.length > 3);

          if (cleanWords.length > 0) {
            const wordMatch = stationsList.find(s => {
              const sWords = s.name.replace(/станция|ст\./gi, '').toLowerCase();
              return cleanWords.some(w => sWords.includes(w));
            });
            if (wordMatch) return wordMatch;
          }

          return null;
        };

        const extractStationFromConsumer = (consumer: string, stationsList: Station[]) => {
          if (!consumer) return null;
          // 1. Try to extract using "ст." or "станция" prefix
          const match = consumer.match(/(?:ст\s*\.\s*|станция\s+)([^\s,;\(\)]+(?:\s+[^\s,;\(\)]+)?)/i);
          if (match && match[1]) {
            let extractedName = match[1].trim();
            // Clean trailing non-station words
            extractedName = extractedName.replace(/\s+(фидер|тп|пост|эц|освещение|отопление|расход|потери|трансформатор|шкаф|сзади|активный|реактивный|ввод|пассажирский|грузовой|парк|север|юг|восток|запад).*$/i, '').trim();
            if (extractedName.length > 2) {
              return extractedName;
            }
          }

          // 2. See if the consumer name contains any of the known stations
          for (const st of stationsList) {
            const cleanStName = st.name.replace(/станция/i, '').trim();
            if (cleanStName && cleanStName.length > 3) {
              if (consumer.toLowerCase().includes(cleanStName.toLowerCase())) {
                return cleanStName;
              }
            }
          }

          return null;
        };

        const findMatchingLossObject = (
          nameInFile: string,
          stationMatch: any,
          lossList: LossObject[]
        ) => {
          if (!nameInFile) return null;
          
          const fileNorm = normalizeForComp(nameInFile);
          const cleanFile = fileNorm.replace(/[^a-zа-я0-9]/gi, '');

          // --- Step 1: Global Exact & Cleaned-Exact Match first! ---
          let globalExact = lossList.find(lo => {
            const configNorm = normalizeForComp(lo.name);
            return configNorm === fileNorm || configNorm.replace(/[^a-zа-я0-9]/gi, '') === cleanFile;
          });
          if (globalExact) return globalExact;

          // Helper to find match within a specific list
          const findInList = (list: LossObject[]) => {
            if (list.length === 0) return null;

            // 1. Direct exact/simple normalized match
            let match = list.find(lo => {
              const configNorm = normalizeForComp(lo.name);
              return configNorm === fileNorm;
            });
            if (match) return match;

            // 2. Direct substring match (case insensitive, normalized)
            const sortedList = [...list].sort((a, b) => b.name.length - a.name.length);
            match = sortedList.find(lo => {
              const configNorm = normalizeForComp(lo.name);
              return fileNorm.includes(configNorm) || configNorm.includes(fileNorm);
            });
            if (match) return match;

            // 3. Overlap/Fuzzy match based on words
            const getWords = (s: string) => s.toLowerCase().split(/[^a-zа-я0-9]+/gi).filter(w => w.length > 0);
            const fileWords = getWords(nameInFile);

            let bestMatch: LossObject | null = null;
            let maxOverlap = 0;

            for (const lo of list) {
              const configWords = getWords(lo.name);
              const overlap = fileWords.filter(w => configWords.includes(w)).length;
              if (overlap > maxOverlap) {
                maxOverlap = overlap;
                bestMatch = lo;
              }
            }

            if (bestMatch && maxOverlap > 0) {
              const configWords = getWords(bestMatch.name);
              const common = fileWords.filter(w => configWords.includes(w));
              const hasSignificantCommon = common.some(w => w !== 'потери' && w !== 'сети' && w !== 'тп' && w !== 'объект' && w !== 'станция' && w !== 'ст');
              if (hasSignificantCommon || maxOverlap >= 2) {
                return bestMatch;
              }
            }

            // 4. Fallback: match by digits only (e.g. both contain "1" or "2")
            const fileDigits = nameInFile.match(/\d+/);
            if (fileDigits) {
              const numStr = fileDigits[0];
              match = sortedList.find(lo => {
                const loDigits = lo.name.match(/\d+/);
                return loDigits && loDigits[0] === numStr;
              });
              if (match) return match;
            }

            return null;
          };

          // --- Phase 1: Local station-scoped search ---
          if (stationMatch) {
            const scopedList = lossList.filter(lo => lo.stationId === stationMatch.id);
            const localMatch = findInList(scopedList);
            if (localMatch) return localMatch;
          }

          // --- Phase 2: Global cross-station search ---
          return findInList(lossList);
        };

        const findMatchingSupplyPoint = (
          cand: { consumer: string; meter: string },
          stationMatch: any,
          pointsList: any[]
        ) => {
          const consumerClean = normalizeForComp(cand.consumer);
          const meterClean = cand.meter ? cand.meter.trim() : "";
          const cleanConsumerOnly = consumerClean.replace(/[^a-zа-я0-9]/gi, '');

          // --- Step 1: Global Exact & Cleaned-Exact Match first! ---
          // 1a. Global exact meter match
          if (meterClean) {
            const byMeterGlobal = pointsList.find(p => p.meterNumber && p.meterNumber.trim() === meterClean);
            if (byMeterGlobal) return byMeterGlobal;
          }

          // 1b. Global exact name match
          const exactMatchGlobal = pointsList.find(p => normalizeForComp(p.name) === consumerClean);
          if (exactMatchGlobal) return exactMatchGlobal;

          // 1c. Global clean-exact name match (ignores punctuation and spaces completely)
          const cleanExactMatchGlobal = pointsList.find(p => normalizeForComp(p.name).replace(/[^a-zа-я0-9]/gi, '') === cleanConsumerOnly);
          if (cleanExactMatchGlobal) return cleanExactMatchGlobal;

          // Compatibility check helper to prevent wrong category matches
          const isCompatible = (nameA: string, nameB: string): boolean => {
            const normA = normalizeForComp(nameA);
            const normB = normalizeForComp(nameB);

            // 1. Check "освещение" / "осв"
            const hasOsvA = normA.includes("освещение") || normA.includes("осв");
            const hasOsvB = normB.includes("освещение") || normB.includes("осв");
            if (hasOsvA !== hasOsvB) return false;

            // 2. Check "отопление" / "э/отопление" / "эотопление" / "тепло" / "печ" / "обогрев"
            const hasOtopA = normA.includes("отопление") || normA.includes("э/отопление") || normA.includes("эотопление") || normA.includes("тепло") || normA.includes("печ") || normA.includes("обогрев");
            const hasOtopB = normB.includes("отопление") || normB.includes("э/отопление") || normB.includes("эотопление") || normB.includes("тепло") || normB.includes("печ") || normB.includes("обогрев");
            if (hasOtopA !== hasOtopB) return false;

            // 3. Check "потери"
            const hasLossA = normA.includes("потери");
            const hasLossB = normB.includes("потери");
            if (hasLossA !== hasLossB) return false;

            // 4. Check "пост эц" / "эц" / "пост"
            const hasEcA = normA.includes("эц") || normA.includes("пост эц") || normA.includes("постэц") || normA.includes("пост");
            const hasEcB = normB.includes("эц") || normB.includes("пост эц") || normB.includes("постэц") || normB.includes("пост");
            if (hasEcA !== hasEcB) return false;

            // 5. Check digits/numbers (e.g. "ТП-4" vs "ТП-5", "фидер 10" vs "фидер 11")
            const digitsA = normA.match(/\d+/g);
            const digitsB = normB.match(/\d+/g);
            if (digitsA && digitsB) {
              if (digitsA.join(',') !== digitsB.join(',')) {
                return false;
              }
            }

            return true;
          };

          // Helper to split a string into significant lowercase words
          const getWords = (s: string) => {
            return s.toLowerCase()
              .replace(/ст\.|станция|№|точка|потребитель|поставка/gi, '')
              .split(/[^a-zа-я0-9]+/gi)
              .filter(w => w.length > 1);
          };

          const fileWords = getWords(cand.consumer);

          // --- 2. LOCAL SEARCH (inside specified stationMatch if provided) ---
          if (stationMatch) {
            // 2a. Try exact meter match inside the station
            if (meterClean) {
              const byMeterInStation = pointsList.find(p => p.stationId === stationMatch.id && p.meterNumber && p.meterNumber.trim() === meterClean);
              if (byMeterInStation) return byMeterInStation;
            }

            // 2b. Try exact name match inside the station
            const exactMatchInStation = pointsList.find(p => p.stationId === stationMatch.id && normalizeForComp(p.name) === consumerClean);
            if (exactMatchInStation) return exactMatchInStation;

            // 2c. Try smart word-based overlap match inside the station
            let bestLocalSp = null;
            let maxLocalOverlap = 0;
            const localPoints = pointsList.filter(p => p.stationId === stationMatch.id);
            for (const p of localPoints) {
              if (!isCompatible(cand.consumer, p.name)) continue;

              const pWords = getWords(p.name);
              const overlap = fileWords.filter(w => pWords.includes(w)).length;
              if (overlap > maxLocalOverlap) {
                maxLocalOverlap = overlap;
                bestLocalSp = p;
              }
            }
            if (bestLocalSp && maxLocalOverlap >= 1) {
              return bestLocalSp;
            }

            // 2d. Try fuzzy meter match inside notes or name inside the station
            if (meterClean) {
              const byMeterInNoteInStation = pointsList.find(p => 
                p.stationId === stationMatch.id && 
                ((p.note && p.note.includes(meterClean)) || p.name.includes(meterClean))
              );
              if (byMeterInNoteInStation) return byMeterInNoteInStation;
            }
          }

          // --- 3. GLOBAL FUZZY SEARCH (across all stations as a fallback) ---
          // 3a. Global smart word-based overlap match
          let bestGlobalSp = null;
          let maxGlobalOverlap = 0;
          for (const p of pointsList) {
            if (!isCompatible(cand.consumer, p.name)) continue;

            const pWords = getWords(p.name);
            const overlap = fileWords.filter(w => pWords.includes(w)).length;
            if (overlap > maxGlobalOverlap) {
              maxGlobalOverlap = overlap;
              bestGlobalSp = p;
            }
          }
          if (bestGlobalSp && maxGlobalOverlap >= 1) {
            return bestGlobalSp;
          }

          // 3b. Global fuzzy meter/note match
          if (meterClean) {
            const byMeterInNoteGlobal = pointsList.find(p => 
              (p.note && p.note.includes(meterClean)) || 
              p.name.includes(meterClean)
            );
            if (byMeterInNoteGlobal) return byMeterInNoteGlobal;
          }

          return null;
        };

        for (const sheetName of wb.SheetNames) {
          const sheetNameLower = sheetName.toLowerCase();
          if (
            sheetNameLower.includes("сумма") ||
            sheetNameLower.includes("итого") ||
            sheetNameLower.includes("всего") ||
            sheetNameLower.includes("свод") ||
            sheetNameLower.includes("накоп") ||
            sheetNameLower.includes("квартал") ||
            sheetNameLower.includes("период") ||
            sheetNameLower.includes("средн") ||
            sheetNameLower.includes("разниц")
          ) {
            continue; // Skip aggregate/summary worksheets entirely
          }

          const sheet = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
          if (rows.length === 0) continue;

          // 1. Detect if it is the RJD format or the standard template format
          let isRjdFormat = false;
          for (let r = 0; r < Math.min(25, rows.length); r++) {
            const row = rows[r];
            if (!row) continue;
            for (let c = 0; c < row.length; c++) {
              const text = String(row[c] || '').toLowerCase();
              if (text.includes('наименование потребителя') || text.includes('№ счетчика') || text.includes('счётчик')) {
                isRjdFormat = true;
                break;
              }
            }
            if (isRjdFormat) break;
          }

          if (!isRjdFormat) {
            // --- ORIGINAL TEMPLATE PARSING ---
            const listRows = XLSX.utils.sheet_to_json(sheet) as any[];
            if (listRows.length === 0) continue;

            originalSheetsProcessedCount++;
            
            listRows.forEach(r => {
              let rawObj = "";
              let rawTp = "";
              let rawVal = 0;
              let rawNote = "";
              let m = Number(selectedMonth);
              let y = Number(selectedYear);

              Object.keys(r).forEach(key => {
                const kLower = key.toLowerCase();
                const valStr = String(r[key] || "").trim();
                
                if (kLower.includes("объект") || kLower.includes("станц")) {
                  rawObj = valStr;
                } else if (kLower.includes("тп") || kLower.includes("точка") || kLower.includes("потребител") || kLower.includes("постав")) {
                  rawTp = valStr;
                } else if (kLower.includes("расход") || kLower.includes("значен") || kLower.includes("квт") || kLower.includes("показан")) {
                  rawVal = Number(r[key] || 0);
                } else if (kLower.includes("месяц")) {
                  const parsedM = Number(r[key]);
                  if (!isNaN(parsedM) && parsedM >= 1 && parsedM <= 12) {
                    m = parsedM;
                  }
                } else if (kLower.includes("год")) {
                  const parsedY = Number(r[key]);
                  if (!isNaN(parsedY) && parsedY >= 2020 && parsedY <= 2040) {
                    y = parsedY;
                  }
                } else if (
                  kLower.includes("причин") || 
                  kLower.includes("примеч") || 
                  kLower.includes("заметк") || 
                  kLower.includes("коммент") || 
                  kLower.includes("описан") ||
                  kLower.includes("note") || 
                  kLower.includes("reason") || 
                  kLower.includes("remark")
                ) {
                  rawNote = valStr;
                }
              });

              if (!rawObj && !rawTp) return;

              const isLoss = rawObj.toLowerCase().includes("потери") || rawTp.toLowerCase().includes("потери");

              if (isLoss) {
                const searchStr = rawTp.toLowerCase().includes("потери") ? rawTp : rawObj;
                const stationMatch = findMatchingStation(rawObj, stationsState);

                let matchLo = findMatchingLossObject(searchStr, stationMatch, lossObjectsState);

                if (!matchLo) {
                  // Do not modify database or add new loss objects from Excel file
                }

                if (matchLo) {
                  totalStandardImportCount++;
                  const idx = lossReadingsState.findIndex(l => l.lossObjectId === matchLo.id && l.year === y && l.month === m);
                  if (idx > -1) {
                    lossReadingsState[idx].value = rawVal;
                  } else {
                    lossReadingsState.push({ lossObjectId: matchLo.id, year: y, month: m, value: rawVal });
                  }
                }
              } else if (rawTp) {
                const stationMatch = findMatchingStation(rawObj, stationsState);
                
                const matchSp = findMatchingSupplyPoint({ consumer: rawTp, meter: "" }, stationMatch, pointsState);
                if (matchSp) {
                  totalStandardImportCount++;
                  if (rawNote) {
                    matchSp.note = rawNote;
                  }
                  const idx = readingsState.findIndex(rd => rd.supplyPointId === matchSp.id && rd.year === y && rd.month === m);
                  if (idx > -1) {
                    readingsState[idx].value = rawVal;
                  } else {
                    readingsState.push({ supplyPointId: matchSp.id, year: y, month: m, value: rawVal });
                  }
                }
              }
            });
          } else {
            // --- RJD FORMAT SPECIAL INTELLIGENT PARSER ---
            let colMeterIdx = 0;
            let colConsumerIdx = 1;
            let col2025Idx = 2;
            let col2026Idx = 3;
            let colNoteIdx = -1;
            let headerRowIdx = -1;

            // Find column headers
            for (let r = 0; r < Math.min(25, rows.length); r++) {
              const row = rows[r];
              if (!row) continue;
              let meterFound = -1;
              let consumerFound = -1;
              let noteFound = -1;
              for (let c = 0; c < row.length; c++) {
                const txt = String(row[c] || '').toLowerCase();
                if (txt.includes('счетчик') || txt.includes('счётчик') || txt.includes('№ cч') || txt.includes('№сч')) {
                  meterFound = c;
                }
                if (txt.includes('наименование') || txt.includes('потребител') || txt.includes('название')) {
                  consumerFound = c;
                }
                if (
                  txt.includes('причин') || 
                  txt.includes('примеч') || 
                  txt.includes('заметк') || 
                  txt.includes('коммент') || 
                  txt.includes('описан') || 
                  txt.includes('note') || 
                  txt.includes('reason') || 
                  txt.includes('remark')
                ) {
                  noteFound = c;
                }
              }
              if (meterFound !== -1 && consumerFound !== -1) {
                colMeterIdx = meterFound;
                colConsumerIdx = consumerFound;
                if (noteFound !== -1) {
                  colNoteIdx = noteFound;
                }
                headerRowIdx = r;
                break;
              }
            }

            // Detect dynamic years from column headers
            let detectedYears = [2025, 2026];
            if (headerRowIdx !== -1) {
              let kwhCols: number[] = [];
              const headerRow = rows[headerRowIdx];
              const prevRow = headerRowIdx > 0 ? rows[headerRowIdx - 1] : null;
              for (let c = colConsumerIdx + 1; c < headerRow.length; c++) {
                if (String(headerRow[c] || '').toLowerCase().includes('квт')) {
                  kwhCols.push(c);
                }
              }
              if (kwhCols.length >= 2) {
                col2025Idx = kwhCols[0];
                col2026Idx = kwhCols[1];
                
                const findYearValue = (col: number) => {
                  if (prevRow) {
                    const val = Number(prevRow[col]);
                    if (!isNaN(val) && val >= 2020 && val <= 2040) return val;
                    const valLeft = Number(prevRow[col - 1]);
                    if (!isNaN(valLeft) && valLeft >= 2020 && valLeft <= 2040) return valLeft;
                  }
                  return null;
                };
                const y1 = findYearValue(kwhCols[0]);
                const y2 = findYearValue(kwhCols[1]);
                if (y1 && y2) {
                  detectedYears = [y1, y2];
                }
              }
            }
            detectedYears.forEach(y => {
              if (!detectedYearsGlobal.includes(y)) {
                detectedYearsGlobal.push(y);
              }
            });

            // Detect Russian Month name from first rows or Fallback to Sheet Name
            let detectedMonth = selectedMonth;
            const russianMonths = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"];
            const russianMonthsShort = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"];
            
            let foundMonth = false;
            // 1. Try to find in cells (skipping headers/footers containing math or period labels)
            for (let r = 0; r < Math.min(12, rows.length); r++) {
              const row = rows[r];
              if (!row) continue;
              for (let c = 0; c < row.length; c++) {
                const cellVal = String(row[c] || "").toLowerCase();
                if (
                  cellVal.includes("сумма") || 
                  cellVal.includes("итого") || 
                  cellVal.includes("всего") || 
                  cellVal.includes("период") || 
                  cellVal.includes("факт") || 
                  cellVal.includes("план") || 
                  cellVal.includes("-") || 
                  cellVal.includes("—")
                ) {
                  continue; // Skip scanning helper or calculation cells for the primary single-month detection
                }
                for (let i = 0; i < russianMonths.length; i++) {
                  if (cellVal.includes(russianMonths[i])) {
                    detectedMonth = i + 1;
                    foundMonth = true;
                    break;
                  }
                }
                if (foundMonth) break;
              }
              if (foundMonth) break;
            }

            // 2. If not found, try sheet title
            if (!foundMonth) {
              const lowerSheetName = sheetName.toLowerCase();
              for (let i = 0; i < russianMonths.length; i++) {
                if (lowerSheetName.includes(russianMonths[i]) || lowerSheetName.includes(russianMonthsShort[i])) {
                  detectedMonth = i + 1;
                  foundMonth = true;
                  break;
                }
              }
            }

            if (!foundMonth) {
              // Skip sheet if no specific month was detected in its title or cells (this prevents sum/aggregate sheets from polluting the June monthly chart)
              continue;
            }

            // Parse block grouping from sub-totals rows e.g. "Итого Спас-Деменск"
            interface CandidatePoint {
              meter: string;
              consumer: string;
              val2025: number;
              val2026: number;
              isLossesRow: boolean;
              stationName?: string;
              note?: string;
            }

            let allCandidates: CandidatePoint[] = [];
            let accumulatedBlock: CandidatePoint[] = [];
            
            const startIdx = headerRowIdx !== -1 ? headerRowIdx + 1 : 0;
            for (let r = startIdx; r < rows.length; r++) {
              const row = rows[r];
              if (!row || row.length <= colConsumerIdx) continue;

              const cellConsumer = String(row[colConsumerIdx] || '').trim();
              if (!cellConsumer) continue;

              const isTotalRow = cellConsumer.toLowerCase().startsWith('итого') || cellConsumer.toLowerCase().includes('итого ');
              if (isTotalRow) {
                const stationNameRaw = cellConsumer.replace(/итого/i, '').trim();
                if (stationNameRaw) {
                  accumulatedBlock.forEach(cand => {
                    cand.stationName = stationNameRaw;
                    allCandidates.push(cand);
                  });
                  accumulatedBlock = []; // Reset block division
                }
                continue;
              }

              const isLossesRow = cellConsumer.toLowerCase().includes('потери');
              const normConsumer = cellConsumer.toLowerCase();
              if (
                !isLossesRow && (
                  normConsumer.includes('справочно') || 
                  normConsumer.includes('факт') || 
                  normConsumer.includes('план') || 
                  normConsumer === 'итого' ||
                  normConsumer.includes('итого ') ||
                  normConsumer.includes('сумма') || 
                  normConsumer.includes('всего') ||
                  normConsumer.includes('среднее') ||
                  normConsumer.includes('разница') ||
                  normConsumer.includes('январ') ||
                  normConsumer.includes('февр') ||
                  normConsumer.includes('март') ||
                  normConsumer.includes('апрел') ||
                  normConsumer.includes('май') ||
                  normConsumer.includes('июн') ||
                  normConsumer.includes('июл') ||
                  normConsumer.includes('август') ||
                  normConsumer.includes('сент') ||
                  normConsumer.includes('окт') ||
                  normConsumer.includes('нояб') ||
                  normConsumer.includes('декаб') ||
                  normConsumer.includes('квт') ||
                  normConsumer.includes('расход')
                )
              ) {
                continue;
              }

              const val2025 = Number(row[col2025Idx] || 0);
              const val2026 = Number(row[col2026Idx] || 0);
              const cellMeter = String(row[colMeterIdx] || '').trim();
              
              // Prefer Column G (index 6, user preference) and fallback to colNoteIdx
              let cellNote = '';
              if (row[6] !== undefined && row[6] !== null && String(row[6]).trim() !== '') {
                cellNote = String(row[6]).trim();
              } else if (colNoteIdx !== -1 && row[colNoteIdx] !== undefined && row[colNoteIdx] !== null) {
                cellNote = String(row[colNoteIdx]).trim();
              }

              accumulatedBlock.push({
                meter: cellMeter,
                consumer: cellConsumer,
                val2025,
                val2026,
                isLossesRow,
                note: cellNote
              });
            }

            // Flush any trailing accumulated rows
            if (accumulatedBlock.length > 0) {
              accumulatedBlock.forEach(cand => {
                allCandidates.push(cand);
              });
            }

            rjdSheetsProcessedCount++;

            // Match and write to local candidates list
            allCandidates.forEach(cand => {
              let stationName = cand.stationName || "";

              // Determine stationName if missing or generic
              const isGeneric = !stationName || 
                                stationName.toLowerCase() === "общий участок" || 
                                stationName.toLowerCase().includes("дистанции") || 
                                stationName.toLowerCase().includes("всего") || 
                                stationName.toLowerCase() === "итого";
              
              if (isGeneric) {
                const extracted = extractStationFromConsumer(cand.consumer, stationsState);
                if (extracted) {
                  stationName = extracted;
                } else {
                  // Try from sheet name
                  const extractedFromSheet = extractStationFromConsumer(sheetName, stationsState);
                  if (extractedFromSheet) {
                    stationName = extractedFromSheet;
                  } else {
                    const hasMonth = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек", "год", "свод"].some(m => sheetName.toLowerCase().includes(m));
                    if (!hasMonth && sheetName.trim().length > 2) {
                      stationName = sheetName.trim();
                    } else {
                      // Look at previous candidate in this sheet with a valid non-generic station name
                      const prevWithStation = [...allCandidates].reverse().find(c => {
                        if (!c.stationName) return false;
                        const lower = c.stationName.toLowerCase();
                        return lower !== "общий участок" && !lower.includes("дистанц") && !lower.includes("всего") && lower !== "итого";
                      });
                      if (prevWithStation && prevWithStation.stationName) {
                        stationName = prevWithStation.stationName;
                      } else {
                        stationName = stationsState[0]?.name || "Вне станций";
                      }
                    }
                  }
                }
              }

              // Ensure it's clean and doesn't contain redundant "Станция" prefix
              stationName = stationName.replace(/^станция\s+/i, '').trim();
              
              // Match the Station in the DB
              const stationMatch = findMatchingStation(stationName, stationsState);

              if (cand.isLossesRow) {
                // PROCESS NETWORK TECHNOLOGY LOSSES
                let lossObj = findMatchingLossObject(cand.consumer, stationMatch, lossObjectsState);

                if (!lossObj) {
                  // Do not modify database or add new loss objects from Excel file
                }

                if (lossObj) {
                  // Apply values for detected years
                  const lossVals = [
                    { y: detectedYears[0], v: cand.val2025 },
                    { y: detectedYears[1], v: cand.val2026 }
                  ];
                  lossVals.forEach(item => {
                    const matchIdx = lossReadingsState.findIndex(lr => lr.lossObjectId === lossObj.id && lr.year === item.y && lr.month === detectedMonth);
                    if (matchIdx > -1) {
                      lossReadingsState[matchIdx].value = item.v;
                    } else {
                      lossReadingsState.push({ lossObjectId: lossObj.id, year: item.y, month: detectedMonth, value: item.v });
                    }
                  });
                  totalRjdImportCount++;
                }

              } else {
                // PROCESS STANDARD METER SUPPLY POINT WITH ROBUST MATCHER
                const matchSp = findMatchingSupplyPoint({ consumer: cand.consumer, meter: cand.meter }, stationMatch, pointsState);

                if (matchSp) {
                  if (cand.meter && !matchSp.meterNumber) {
                    matchSp.meterNumber = cand.meter;
                  }
                  if (cand.note) {
                    matchSp.note = cand.note;
                  }

                  // Directly write readings for both years
                  const vals = [
                    { y: detectedYears[0], v: cand.val2025 },
                    { y: detectedYears[1], v: cand.val2026 }
                  ];
                  vals.forEach(item => {
                    const matchIdx = readingsState.findIndex(rd => rd.supplyPointId === matchSp.id && rd.year === item.y && rd.month === detectedMonth);
                    if (matchIdx > -1) {
                      readingsState[matchIdx].value = item.v;
                    } else {
                      readingsState.push({ supplyPointId: matchSp.id, year: item.y, month: detectedMonth, value: item.v });
                    }
                  });
                  totalRjdImportCount++;
                }
              }
            });
          }
        } // Loop end

        // Save State
        setStations(stationsState);
        setSupplyPoints(pointsState);
        setLossObjects(lossObjectsState);
        setReadings(readingsState);
        setLossReadings(lossReadingsState);
        saveToStorage(stationsState, pointsState, categories, readingsState, lossObjectsState, lossReadingsState);

        setImportWizardQueue([]);
        
        if (originalSheetsProcessedCount > 0) {
          setImportStatus(`Импорт завершен! Обработано ${originalSheetsProcessedCount} листов. Записано ${totalStandardImportCount} показаний.`);
        } else if (rjdSheetsProcessedCount > 0) {
          setImportStatus(`Импорт завершен! Обработано ${rjdSheetsProcessedCount} листов РЖД. Записано ${totalRjdImportCount} показаний за годы: ${detectedYearsGlobal.length > 0 ? detectedYearsGlobal.join(', ') : '2025, 2026'} гг.`);
        } else {
          setImportStatus("Файл импортирован, но значимые данные не найдены.");
        }
        setTimeout(() => setImportStatus(null), 5000);

      } catch (err) {
        alert("Ошибка разбора ведомости. Пожалуйста, используйте официальный шаблон или стандартный лист РЖД.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const resolveSingleWizardPoint = (decision: 'create' | 'link' | 'skip', payload?: any) => {
    const curObj = importWizardQueue[0] as any;
    if (!curObj) return;

    let pointsState = [...supplyPoints];
    let readingsState = [...readings];

    const currentYears = curObj.years || [2025, 2026];
    const spMeterValue = curObj.meter || "";

    if (decision === 'create') {
      const spId = `tp-imp-gen-${Date.now()}`;
      const stationMatch = stations.find(s => s.name.toLowerCase() === curObj.proposedStationName.toLowerCase()) || stations[0];
      
      pointsState.push({
        id: spId,
        name: curObj.unknownPointName,
        stationId: stationMatch.id,
        category: payload.category || "Прочие",
        note: spMeterValue ? `Счетчик №${spMeterValue}` : "Создана из файла ведомости",
        isActive: true,
        meterNumber: spMeterValue
      });

      curObj.rows.forEach((r: any) => {
        const v2025 = Number(r.val2025 || 0);
        const v2026 = Number(r.val2026 || 0);
        
        // Push both years readings dynamically
        readingsState.push({ supplyPointId: spId, year: currentYears[0], month: selectedMonth, value: v2025 });
        readingsState.push({ supplyPointId: spId, year: currentYears[1], month: selectedMonth, value: v2026 });
      });

    } else if (decision === 'link') {
      const spId = payload.existingId;
      
      // Update point with meter mapping for next imports
      pointsState = pointsState.map(p => {
        if (p.id === spId && spMeterValue) {
          return { ...p, meterNumber: spMeterValue };
        }
        return p;
      });

      curObj.rows.forEach((r: any) => {
        const v2025 = Number(r.val2025 || 0);
        const v2026 = Number(r.val2026 || 0);

        const vals = [
          { y: currentYears[0], v: v2025 },
          { y: currentYears[1], v: v2026 }
        ];

        vals.forEach(item => {
          const matchIdx = readingsState.findIndex(rd => rd.supplyPointId === spId && rd.year === item.y && rd.month === selectedMonth);
          if (matchIdx > -1) {
            readingsState[matchIdx].value = item.v;
          } else {
            readingsState.push({ supplyPointId: spId, year: item.y, month: selectedMonth, value: item.v });
          }
        });
      });
    }

    const rest = importWizardQueue.slice(1);
    setSupplyPoints(pointsState);
    setReadings(readingsState);
    saveToStorage(stations, pointsState, categories, readingsState, lossObjects, lossReadings);

    setImportWizardQueue(rest);
    if (rest.length === 0) {
      setImportStatus("Ведомость успешно сопоставлена и загружена!");
      setTimeout(() => setImportStatus(null), 4000);
    }
  };

  const commitImporterRows = (rowsList: any[]) => {
    let readingsState = [...readings];
    let lossesState = [...lossReadings];
    let count = 0;

    rowsList.forEach(r => {
      const rawObj = (r["Название объекта"] || r["Станция"] || r["Объект"] || "").toString().trim();
      const rawTp = (r["Название ТП"] || r["Точка поставки"] || r["ТП"] || "").toString().trim();
      const rawVal = Number(r["Расход (кВтч)"] || r["Расход"] || r["Значение"] || 0);
      const m = Number(r["Месяц"] || selectedMonth);
      const y = Number(r["Год"] || selectedYear);

      if (!rawObj) return;

      if (rawObj.toLowerCase().includes("потери")) {
        const matchLo = lossObjects.find(lo => lo.name.toLowerCase() === rawObj.toLowerCase());
        if (matchLo) {
          count++;
          const idx = lossesState.findIndex(l => l.lossObjectId === matchLo.id && l.year === y && l.month === m);
          if (idx > -1) lossesState[idx].value = rawVal;
          else lossesState.push({ lossObjectId: matchLo.id, year: y, month: m, value: rawVal });
        }
      } else if (rawTp) {
        const matchSp = supplyPoints.find(p => p.name.toLowerCase() === rawTp.toLowerCase());
        if (matchSp) {
          count++;
          const idx = readingsState.findIndex(rd => rd.supplyPointId === matchSp.id && rd.year === y && rd.month === m);
          if (idx > -1) readingsState[idx].value = rawVal;
          else readingsState.push({ supplyPointId: matchSp.id, year: y, month: m, value: rawVal });
        }
      }
    });

    setReadings(readingsState);
    setLossReadings(lossesState);
    saveToStorage(stations, supplyPoints, categories, readingsState, lossObjects, lossesState);

    setImportStatus(`Импортировано записей энергопотребления: ${count}`);
    setTimeout(() => setImportStatus(null), 4005);
  };

  // Template Excel Builder
  const downloadReportTemplate = () => {
    const data = [
      ["Месяц", "Год", "Название объекта", "Название ТП", "Категория нагрузки", "Расход (кВтч)", "Примечание"],
      [6, 2026, "Станция Александров-1", "ТП-1 Фидер №1 (Пост ЭЦ)", "Пост ЭЦ", 4850, "Нормал"],
      [6, 2026, "Станция Александров-1", "ТП-1 Фидер №5 (Освещение)", "Освещение парков", 1920, "Вышки"],
      [6, 2026, "Потери 1 (Трансформатор Т-1 Александров)", "", "", 1430, "По сети"],
      [6, 2026, "Станция Александров-1", "НОВАЯ ТП-99 (Удаленный фидер)", "Прочие", 955, "Демонстрация сопоставления"]
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Шаблон");
    XLSX.writeFile(wb, "Shablon_Energomonitoring.xlsx");
  };

  const downloadDbDump = () => {
    const content = { stations, supplyPoints, categories, readings, lossObjects, lossReadings };
    const blob = new Blob([JSON.stringify(content, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `dump_energomonitoring_${Date.now()}.json`;
    link.click();
  };

  const handleDbReset = () => {
    if (confirm("Вы желаете сбросить локальную базу данных к исходным демонстрационным показателям ОАО РЖД?")) {
      const db = generateDeterministicDatabase();
      syncDatabase(INITIAL_STATIONS, INITIAL_SUPPLY_POINTS, INITIAL_CATEGORIES, db.readings, INITIAL_LOSS_OBJECTS, db.lossReadings);
      alert("Локальное хранилище успешно восстановлено!");
    }
  };

  const handleClearAllReadings = () => {
    if (confirm("Вы желаете полностью стереть все текущие показания и потери по всем месяцам и годам? Все значения будут сброшены в 0. Вы сможете заново импортировать файлы за Январь и Февраль, и на графиках за остальные месяцы будут отображаться только нули (без демонстрационных данных).")) {
      syncDatabase(stations, supplyPoints, categories, [], lossObjects, []);
      alert("Все показания успешно очищены до 0! Теперь вы можете повторно импортировать ваши Excel-файлы.");
    }
  };

  const handleCreateManualBackup = async () => {
    const payload = JSON.stringify({ stations, supplyPoints, categories, readings, lossObjects, lossReadings });
    const d = new Date();
    const timeStr = d.toLocaleString('ru-RU');
    const customName = prompt("Введите имя или комментарий к резервной копии:", `Ручное сохранение (${timeStr})`);
    if (customName === null) return; // cancelled
    
    const newBackup: BackupItem = {
      id: `backup-${Date.now()}`,
      timestamp: timeStr,
      name: customName.trim() || `Ручное сохранение (${timeStr})`,
      data: payload
    };
    
    const updated = [newBackup, ...backupsList].slice(0, 15);
    setBackupsList(updated);
    await dbSet('st_backups_list', updated);
    alert("Резервная копия успешно создана в памяти устройства!");
  };

  const handleRestoreFromBackup = (backup: BackupItem) => {
    if (confirm(`Вы уверены, что хотите восстановить базу данных из копии "${backup.name}"?\nВсе текущие несохраненные изменения будут перезаписаны.`)) {
      try {
        const parsed = JSON.parse(backup.data);
        if (parsed.stations && parsed.supplyPoints && parsed.categories && parsed.readings && parsed.lossObjects && parsed.lossReadings) {
          // Temporarily disable auto backup for restore to avoid creating a backup of the restoration itself:
          const wasAutoEnabled = autoBackupEnabled;
          localStorage.setItem('st_auto_backup', 'false');
          
          syncDatabase(
            parsed.stations,
            parsed.supplyPoints,
            parsed.categories,
            parsed.readings,
            parsed.lossObjects,
            parsed.lossReadings
          );
          
          localStorage.setItem('st_auto_backup', wasAutoEnabled ? 'true' : 'false');
          alert("База данных успешно восстановлена из резервной копии!");
        } else {
          alert("Неверный формат данных резервной копии.");
        }
      } catch (err) {
        alert("Ошибка декодирования резервной копии.");
      }
    }
  };

  const handleDeleteBackup = (id: string) => {
    setDeleteConfirm({
      isOpen: true,
      title: "Удалить точку восстановления?",
      message: "Вы уверены, что хотите окончательно удалить эту резервную копию?",
      onConfirm: async () => {
        const updated = backupsList.filter(b => b.id !== id);
        setBackupsList(updated);
        await dbSet('st_backups_list', updated);
      }
    });
  };

  const handleImportBackupJsonFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const parsed = JSON.parse(event.target?.result as string);
        if (parsed.stations && parsed.supplyPoints && parsed.categories && parsed.readings && parsed.lossObjects && parsed.lossReadings) {
          syncDatabase(
            parsed.stations,
            parsed.supplyPoints,
            parsed.categories,
            parsed.readings,
            parsed.lossObjects,
            parsed.lossReadings
          );
          alert("База данных успешно загружена и импортирована из файла резервной копии!");
        } else {
          alert("Файл резервной копии не содержит корректных таблиц РЖД.");
        }
      } catch (err) {
        alert("Ошибка чтения JSON файла.");
      }
    };
    reader.readAsText(file);
    e.target.value = ''; // Reset input to allow selecting same file again
  };

  // --- FILTERED ARRAYS ---
  const stationsListFiltered = useMemo(() => {
    return stations.filter(s => s.name.toLowerCase().includes(stationSearch.toLowerCase()) || s.note.toLowerCase().includes(stationSearch.toLowerCase()));
  }, [stations, stationSearch]);

  const pointsListFiltered = useMemo(() => {
    return supplyPoints.filter(p => p.name.toLowerCase().includes(pointSearch.toLowerCase()) || p.note.toLowerCase().includes(pointSearch.toLowerCase()));
  }, [supplyPoints, pointSearch]);

  if (isDataLoading) {
    return (
      <div className={`min-h-screen flex flex-col items-center justify-center ${darkTheme ? 'bg-[#0f172a] text-slate-100' : 'bg-[#f8fafc] text-slate-800'}`}>
        <div className="flex flex-col items-center gap-4 max-w-sm px-6 text-center">
          <div className="w-14 h-14 rounded-2xl bg-blue-600 flex items-center justify-center shadow-lg shadow-blue-500/20">
            <RefreshCw className="w-7 h-7 text-white animate-spin animate-infinite" />
          </div>
          <div>
            <h2 className="text-base font-bold tracking-tight">Энергомониторинг РЖД</h2>
            <p className="text-xs text-slate-400 mt-1">Загрузка автономной базы данных...</p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className={`min-h-screen text-sans flex flex-col transition-colors duration-150 select-none ${darkTheme ? 'bg-[#0f172a] text-slate-100' : 'bg-[#f8fafc] text-slate-800'}`}>
      
      {/* Huawei Status Indicator & Header Block */}
      <header className={`min-h-[64px] flex items-center justify-between border-b px-6 ${darkTheme ? 'bg-[#1e293b]/90 border-slate-700/80' : 'bg-white border-slate-205'} sticky top-0 z-40 shadow-sm backdrop-blur`}>
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-blue-600 flex items-center justify-center shadow-lg">
            <Train className="w-5 h-5 text-white" />
          </div>
          <div>
            <h1 className="text-sm font-bold tracking-tight text-blue-500">Энергомониторинг станций</h1>
            <span className="text-[9px] text-slate-400 block font-semibold uppercase tracking-wider flex items-center gap-1.5">
              <span>Huawei MatePad 11.5</span>
              <span>•</span>
              {isOnline ? (
                <span className="text-emerald-500 font-bold flex items-center gap-1">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                  В сети
                </span>
              ) : (
                <span className="text-amber-500 font-bold flex items-center gap-1">
                  <span className="w-1.5 h-1.5 rounded-full bg-amber-500 animate-pulse"></span>
                  Автономно
                </span>
              )}
            </span>
          </div>
        </div>

        {/* Filters */}
        <div className="flex items-center gap-2">
          <select
            value={selectedPeriodId}
            onChange={(e) => {
              const val = e.target.value;
              setSelectedPeriodId(val);
              const opt = PERIOD_OPTIONS.find(p => p.id === val);
              if (opt && opt.months.length > 0) {
                setSelectedMonth(opt.months[opt.months.length - 1]);
              }
            }}
            className={`px-3 py-1.5 rounded-lg text-xs border font-semibold outline-none cursor-pointer ${darkTheme ? 'bg-slate-800 border-slate-705 text-white' : 'bg-slate-50 border-slate-200 text-slate-800'}`}
          >
            {PERIOD_OPTIONS.map((opt) => (
              <option key={opt.id} value={opt.id}>
                {opt.name}
              </option>
            ))}
          </select>
          <select
            value={selectedYear}
            onChange={(e) => setSelectedYear(Number(e.target.value))}
            className={`px-3 py-1.5 rounded-lg text-xs border font-semibold outline-none cursor-pointer ${darkTheme ? 'bg-slate-800 border-slate-705 text-white' : 'bg-slate-50 border-slate-200 text-slate-800'}`}
          >
            <option value={2025}>2025 г.</option>
            <option value={2026}>2026 г.</option>
          </select>
        </div>

        {/* Right tools */}
        <div className="flex items-center gap-3">
          <button onClick={toggleTheme} className={`p-2 rounded-xl border transition-all ${darkTheme ? 'bg-slate-800 border-slate-700 text-yellow-400' : 'bg-slate-50 border-slate-200 text-slate-650'}`}>
            {darkTheme ? <Sun className="w-4.5 h-4.5" /> : <Moon className="w-4.5 h-4.5" />}
          </button>
          
          <div className="hidden md:flex items-center gap-3 text-[11px] font-mono border-l pl-3 border-slate-700">
            <span>АКБ: <strong className="text-emerald-500">{batteryLevel}%</strong></span>
            <span>{currentTimeStr}</span>
          </div>

          <button onClick={triggerExcelSelect} className="flex items-center gap-1.5 bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs py-2 px-3 rounded-lg active:scale-95">
            <Upload className="w-4 h-4" />
            <span>Загрузить ведомость</span>
          </button>
          <input id="hidden-excel-picker" type="file" accept=".xlsx,.xls" onChange={handleExcelImport} className="hidden" />
        </div>
      </header>

      {/* RESOLUTION DIALOG MODAL */}
      {importWizardQueue.length > 0 && (
        <div className="fixed inset-0 bg-black/75 flex items-center justify-center p-4 z-50 backdrop-blur-sm">
          <div className={`w-full max-w-md rounded-2xl border p-5 shadow-2xl space-y-4 ${darkTheme ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-220'}`}>
            <div className="flex items-center gap-2 border-b pb-3 border-slate-755">
              <AlertTriangle className="w-5.5 h-5.5 text-amber-500" />
              <div>
                <h3 className="font-bold text-sm">Не известная в БД точка</h3>
                <span className="text-[10px] text-slate-400">Осталось обработать: {importWizardQueue.length}</span>
              </div>
            </div>

            <div className={`p-3.5 rounded-xl border text-xs space-y-2 ${darkTheme ? 'bg-slate-800/60 border-slate-700' : 'bg-slate-50 border-slate-200'}`}>
              <div>
                <span className="text-slate-400 font-medium">Точка поставки фидера:</span>
                <strong className="block text-sm text-blue-400 mt-0.5">{importWizardQueue[0].unknownPointName}</strong>
              </div>
              <div>
                <span className="text-slate-400 font-medium">Объект станции из Excel:</span>
                <strong className="block text-slate-200">{importWizardQueue[0].proposedStationName}</strong>
              </div>
            </div>

            <div className="space-y-2 pt-1">
              <button
                onClick={() => resolveSingleWizardPoint('create', { category: importWizardQueue[0].proposedCategory })}
                className="w-full text-left p-2.5 rounded-xl border border-blue-500/30 text-xs hover:bg-blue-500/10 active:scale-[99%]"
              >
                <span className="font-bold text-blue-400 block text-[10px]">Вариант 1</span>
                <strong>Создать новую точку для текущей станции</strong>
              </button>

              <div className={`p-2.5 rounded-xl border text-xs space-y-2 ${darkTheme ? 'bg-slate-850 border-slate-800' : 'bg-slate-50'}`}>
                <span className="font-bold text-amber-500 block text-[10px]">Вариант 2</span>
                <span>Связать с работающей точкой ТП:</span>
                <div className="flex gap-1.5 mt-1">
                  <select
                    id="wizard-link-selector"
                    className={`text-xs px-2 py-1 rounded border grow outline-none ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
                  >
                    {supplyPoints.map(p => (
                      <option key={p.id} value={p.id}>{p.name}</option>
                    ))}
                  </select>
                  <button
                    onClick={() => {
                      const el = document.getElementById('wizard-link-selector') as HTMLSelectElement;
                      if (el) resolveSingleWizardPoint('link', { existingId: el.value });
                    }}
                    className="bg-amber-600 hover:bg-amber-500 text-white px-3.5 py-1 text-xs rounded font-bold"
                  >
                    Привязать
                  </button>
                </div>
              </div>

              <button onClick={() => resolveSingleWizardPoint('skip')} className="w-full text-center text-xs text-slate-400 hover:text-white pt-2 block">
                Пропустить импорт этой строки
              </button>
            </div>
          </div>
        </div>
      )}

      {/* CORE FRAME FOR LEFT DRAWER & ACTIVE BOX */}
      <div className="flex grow overflow-hidden">
        
        {/* NAV BAR SIDEBAR */}
        <aside className={`w-[245px] shrink-0 border-r flex flex-col justify-between ${darkTheme ? 'bg-[#0f172a] border-slate-750' : 'bg-[#f1f5f9] border-slate-200'}`}>
          <nav className="p-3 space-y-1.5">
            <span className="text-xs text-slate-450 block font-extrabold uppercase tracking-wider px-3 py-1.5">Меню управления</span>
            {[
              { id: 'home', label: 'Главная', icon: LayoutDashboard },
              { id: 'stations', label: 'Станции', icon: Train },
              { id: 'losses', label: 'Потери', icon: ZapOff },
              { id: 'analytics', label: 'Аналитика', icon: TrendingUp },
              { id: 'reports', label: 'Отчеты', icon: FileText },
              { id: 'directories', label: 'Справочники', icon: BookOpen },
              { id: 'settings', label: 'Настройки', icon: Settings }
            ].map(tab => {
              const Icon = tab.icon;
              const active = activeTab === tab.id;
              return (
                <button
                  key={tab.id}
                  onClick={() => { setActiveTab(tab.id); setPointSearch(''); setStationSearch(''); }}
                  className={`w-full flex items-center gap-3 py-3 px-3.5 rounded-xl text-sm font-bold text-left transition-all ${
                    active ? 'bg-blue-600/15 text-blue-400 border border-blue-500/20 shadow-sm' : 'text-slate-400 hover:bg-slate-800/10 hover:text-slate-300'
                  }`}
                >
                  <Icon className="w-4.5 h-4.5 shrink-0 text-blue-500/80" />
                  <span>{tab.label}</span>
                </button>
              );
            })}
          </nav>

          <div className="p-4 border-t border-slate-800 text-[10px] font-mono text-slate-450 space-y-1">
            <div>Базовых станций: {stations.length}</div>
            <div>Контрольных ТП: {supplyPoints.length}</div>
          </div>
        </aside>

        {/* WORKSPACE PREVIEW PAGE CONTAINER */}
        <main className="grow overflow-y-auto p-6 md:p-8">
          
          {importStatus && (
            <div className="mb-4 p-3 rounded-xl border bg-emerald-500/10 border-emerald-500/30 text-emerald-400 text-xs font-bold flex items-center justify-between">
              <span>{importStatus}</span>
              <button onClick={() => setImportStatus(null)}><X className="w-4 h-4 cursor-pointer" /></button>
            </div>
          )}

          {navigatedFromAnomalies && (
            <div className="mb-6 p-4 rounded-xl border bg-amber-500/10 border-amber-500/30 text-amber-400 text-xs font-bold flex flex-col sm:flex-row sm:items-center justify-between gap-3 shadow-lg shadow-amber-500/5 animate-fade-in">
              <div className="flex items-center gap-2.5">
                <AlertTriangle className="w-5 h-5 text-amber-500 animate-pulse shrink-0" />
                <div>
                  <span className="block font-black">Режим анализа аномалий активен</span>
                  <span className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider block mt-0.5">Исследуются показатели и динамика выбранного объекта</span>
                </div>
              </div>
              <div className="flex items-center gap-2 shrink-0">
                <button 
                  onClick={() => setAnomaliesModalOpen(true)}
                  className="px-3.5 py-2 bg-amber-500 text-slate-950 hover:bg-amber-400 transition-all font-black rounded-lg flex items-center gap-1 text-xs cursor-pointer"
                >
                  ◀ Вернуться к списку аномалий
                </button>
                <button 
                  onClick={() => setNavigatedFromAnomalies(false)}
                  className="text-slate-400 hover:text-white transition-colors p-1.5 hover:bg-slate-800/50 rounded-lg"
                  title="Закрыть режим анализа"
                >
                  <X className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}

          {/* TAB 1: ГЛАВНАЯ (HOME/SWOT) */}
          {activeTab === 'home' && (
            <div className="space-y-6">
              
              {/* KPI indicators */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                
                <div role="contentinfo" className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200 shadow-sm'}`}>
                  <span className="text-xs text-slate-400 block uppercase font-extrabold tracking-wider">Общий расход ТП</span>
                  <strong className="text-2xl font-black block mt-1.5">{totalCurrValues.toLocaleString()} кВт·ч</strong>
                  <span className={`text-xs block mt-1.5 font-bold ${totalDeltaAbs > 0 ? "text-rose-500" : "text-emerald-500"}`}>
                    {totalDeltaAbs > 0 ? `+${totalDeltaAbs.toLocaleString()}` : totalDeltaAbs.toLocaleString()} ({totalDeltaPercent > 0 ? `+${totalDeltaPercent.toFixed(1)}%` : `${totalDeltaPercent.toFixed(1)}%`})
                  </span>
                </div>

                <div role="contentinfo" className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200 shadow-sm'}`}>
                  <span className="text-xs text-slate-400 block uppercase font-extrabold tracking-wider">Расход прошлого года</span>
                  <strong className="text-2xl font-black block mt-1.5">{totalPrevValues.toLocaleString()} кВт·ч</strong>
                  <span className="text-xs text-slate-500 block mt-1.5 font-semibold">Базовый месяц {selectedYear - 1} г.</span>
                </div>

                <div role="contentinfo" className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200 shadow-sm'}`}>
                  <span className="text-xs text-slate-400 block uppercase font-extrabold tracking-wider">Технические потери</span>
                  <strong className="text-2xl font-black block mt-1.5">{totalCurrLosses.toLocaleString()} кВт·ч</strong>
                  <span className={`text-xs block mt-1.5 font-bold ${totalLossDeltaAbs > 0 ? "text-rose-400" : "text-emerald-400"}`}>
                    Лимит {totalPrevLosses.toLocaleString()} ({totalLossDeltaPercent > 0 ? 'Превышение' : 'Экономия'})
                  </span>
                </div>

                <div 
                  role="button"
                  tabIndex={0}
                  onClick={() => setAnomaliesModalOpen(true)}
                  className={`p-4 rounded-xl border cursor-pointer transition-all duration-250 active:scale-[0.98] ${
                    darkTheme 
                      ? 'bg-slate-800/40 border-slate-700 hover:border-amber-500 hover:bg-slate-800/60 shadow-lg shadow-amber-500/5' 
                      : 'bg-white border-slate-200 shadow-sm hover:border-amber-400 hover:shadow-md'
                  }`}
                >
                  <div className="flex justify-between items-start">
                    <span className="text-xs text-slate-400 block uppercase font-extrabold tracking-wider">Критические Аномалии</span>
                    {currentAnomalies.length > 0 && (
                      <span className="flex h-2 w-2 relative">
                        <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
                        <span className="relative inline-flex rounded-full h-2 w-2 bg-amber-500"></span>
                      </span>
                    )}
                  </div>
                  <strong className={`text-2xl font-black block mt-1.5 ${currentAnomalies.length > 0 ? 'text-amber-500' : 'text-emerald-500'}`}>
                    {currentAnomalies.length} инцидентов
                  </strong>
                  <span className="text-xs text-slate-500 block mt-1.5 font-semibold flex items-center gap-1.5">
                    <span>⚡ Нажмите для просмотра</span>
                    <ChevronRight className="w-3.5 h-3.5 text-slate-400" />
                  </span>
                </div>

              </div>

              {/* Leaderboards */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
                
                <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/30 border-slate-700' : 'text-slate-800 border-slate-200'}`}>
                  <span className="text-xs font-bold uppercase text-slate-400 tracking-wider block mb-3">Наибольший перерасход (Отклонения в +)</span>
                  <div className="space-y-2 text-xs">
                    {leadersList.overruns.slice(0, 3).map((l, i) => (
                      <div key={i} className="flex justify-between items-center py-1.5 border-b border-slate-700 pb-1.5 last:border-0 cursor-pointer" onClick={() => { setSelectedStationId(l.station.id); setActiveTab('stations'); }}>
                        <span className="font-bold">{l.station.name}</span>
                        <span className="text-rose-500 font-semibold font-mono">+{l.deltaAbs.toLocaleString()} кВт·ч</span>
                      </div>
                    ))}
                  </div>
                </div>

                <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/30 border-slate-700' : 'text-slate-800 border-slate-200'}`}>
                  <span className="text-xs font-bold uppercase text-slate-400 tracking-wider block mb-3">Наибольшая экономия (Отклонения в -)</span>
                  <div className="space-y-2 text-xs">
                    {leadersList.savings.slice(0, 3).map((l, i) => (
                      <div key={i} className="flex justify-between items-center py-1.5 border-b border-slate-700 pb-1.5 last:border-0 cursor-pointer" onClick={() => { setSelectedStationId(l.station.id); setActiveTab('stations'); }}>
                        <span className="font-bold">{l.station.name}</span>
                        <span className="text-emerald-550 font-semibold font-mono">{l.deltaAbs.toLocaleString()} кВт·ч</span>
                      </div>
                    ))}
                  </div>
                </div>

              </div>

              {/* Anomaly list panel */}
              <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/20 border-slate-700' : 'text-slate-800 border-slate-200'}`}>
                <div 
                  className="flex justify-between items-center mb-4 cursor-pointer group" 
                  onClick={() => {
                    setAnomalySearchQuery('');
                    setAnomaliesModalOpen(true);
                  }}
                >
                  <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400 group-hover:text-blue-500 transition-colors flex items-center gap-1.5">
                    <span>Журнал событий, пиков и отклонений</span>
                    <ChevronRight className="w-3.5 h-3.5 text-slate-500 group-hover:text-blue-500 transition-colors" />
                  </h3>
                  <button 
                    onClick={(e) => { 
                      e.stopPropagation(); 
                      setAnomalySearchQuery('');
                      setAnomaliesModalOpen(true); 
                    }}
                    className="text-[10px] font-bold text-blue-500 hover:text-blue-400 transition-colors uppercase tracking-wider flex items-center gap-1"
                  >
                    <span>Открыть список ({currentAnomalies.length})</span>
                    <ChevronRight className="w-3 h-3" />
                  </button>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  {currentAnomalies.length === 0 ? (
                    <p className="text-xs text-slate-450 text-center md:col-span-2 py-4">Все показатели укладываются в рамки лимитов.</p>
                  ) : (
                    currentAnomalies.slice(0, 6).map(anom => {
                      const getLabel = (type: string) => {
                        if (type === 'station') return 'Станция';
                        if (type === 'supply_point') return 'Точка ТП';
                        return 'Потери';
                      };
                      return (
                        <div 
                          key={anom.id} 
                          onClick={() => {
                            setAnomalySearchQuery(anom.targetName);
                            setAnomaliesModalOpen(true);
                          }}
                          className={`p-3 rounded-lg text-xs border cursor-pointer transition-all duration-150 active:scale-[0.98] group/card ${
                            darkTheme 
                              ? 'bg-slate-900/60 border-slate-800 hover:border-amber-500/50 hover:bg-slate-900' 
                              : 'bg-slate-50 border-slate-200 hover:border-amber-450 hover:bg-slate-100 shadow-sm'
                          } space-y-1.5`}
                        >
                          <div className="flex justify-between items-center">
                            <span className="text-[10px] uppercase font-bold text-slate-400">{getLabel(anom.targetType)}</span>
                            <span className="text-[10px] bg-red-650 text-white font-mono px-2 py-0.5 rounded font-bold">{anom.metric}</span>
                          </div>
                          <strong className="block font-bold group-hover/card:text-blue-500 transition-colors">{anom.targetName}</strong>
                          <p className="text-[11px] text-slate-400 leading-relaxed">{anom.description}</p>
                          
                          <div className="flex justify-between items-center pt-1 border-t border-slate-800/20 mt-1">
                            <span className="text-[9.5px] text-amber-500 font-bold flex items-center gap-0.5">
                              <span>Детали</span>
                              <ChevronRight className="w-2.5 h-2.5" />
                            </span>
                            <button 
                              onClick={(e) => {
                                e.stopPropagation();
                                if (anom.targetType === 'station') {
                                  setSelectedStationId(anom.targetId);
                                  setActiveTab('stations');
                                } else if (anom.targetType === 'supply_point') {
                                  const point = supplyPoints.find(p => p.id === anom.targetId);
                                  if (point) {
                                    setSelectedStationId(point.stationId);
                                    setSelectedSupplyPointId(point.id);
                                    setActiveTab('stations');
                                  }
                                } else if (anom.targetType === 'loss') {
                                  setSelectedLossId(anom.targetId);
                                  setActiveTab('losses');
                                }
                              }}
                              className="text-[10px] bg-blue-600/10 hover:bg-blue-600 hover:text-white text-blue-400 font-bold px-2 py-1 rounded transition-all duration-150 flex items-center gap-0.5"
                            >
                              <span>К объекту</span>
                              <ChevronRight className="w-2.5 h-2.5" />
                            </button>
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>
                {currentAnomalies.length > 6 && (
                  <div className="text-center mt-3 pt-3 border-t border-slate-800/30">
                    <button 
                      onClick={() => setAnomaliesModalOpen(true)}
                      className="text-[10px] font-bold text-slate-400 hover:text-white transition-colors uppercase tracking-wider"
                    >
                      Показать еще {currentAnomalies.length - 6} аномалий...
                    </button>
                  </div>
                )}
              </div>

            </div>
          )}
          {/* TAB 2: СТАНЦИИ (STATION CARDS & MANAGEMENT) */}
          {activeTab === 'stations' && (
            <div className="grid grid-cols-1 xl:grid-cols-3 gap-6 items-start">
              
              {/* Sidebar list directory */}
              <div className="xl:col-span-1 space-y-4">
                <div className="flex items-center justify-between gap-2 border-b pb-2 border-slate-700">
                  <h3 className="text-base font-extrabold uppercase tracking-wider">Железнодорожные Станции</h3>
                  <button onClick={() => setStationModal({ isOpen: true, mode: 'add', name: '', section: '', note: '' })} className="bg-blue-600 hover:bg-blue-500 text-white text-xs px-2.5 py-1.5 rounded font-bold">
                    + Новая
                  </button>
                </div>
                
                <input
                  type="text"
                  value={stationSearch}
                  onChange={(e) => setStationSearch(e.target.value)}
                  placeholder="Фильтр станций..."
                  className={`w-full text-sm px-3.5 py-2.5 rounded-lg border outline-none ${darkTheme ? 'bg-slate-800 border-slate-700' : 'bg-white'}`}
                />

                <div className="space-y-2 max-h-[360px] overflow-y-auto">
                  {stationsListFiltered.map(st => (
                    <div
                      key={st.id}
                      onClick={() => setSelectedStationId(st.id)}
                      className={`p-3 rounded-lg border text-sm cursor-pointer transition ${
                        st.id === selectedStationId ? 'border-blue-500 bg-blue-500/[4%]' : 'border-slate-800 bg-slate-900/40 hover:bg-slate-800/40'
                      }`}
                    >
                      <h4 className="font-bold">{st.name}</h4>
                      <p className="text-xs text-slate-400 line-clamp-1 mt-1">{st.note || 'Примечания отсутствуют.'}</p>
                      <div className="mt-2.5 flex justify-end gap-3 text-xs">
                        <button onClick={(e) => { e.stopPropagation(); setStationModal({ isOpen: true, mode: 'edit', id: st.id, name: st.name, section: st.section || '', note: st.note }); }} className="text-slate-450 hover:text-sky-400">Правка</button>
                        <button onClick={(e) => { e.stopPropagation(); handleDeleteStation(st.id); }} className="text-slate-450 hover:text-rose-500 font-bold">Удалить</button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Detail Profile Panel */}
              <div className="xl:col-span-2 space-y-6">
                {activeStation ? (
                  <div className="space-y-5">
                    
                    {/* Header values block */}
                    <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-[#1e293b]/55 border-slate-700' : 'bg-white shadow-sm'}`}>
                      <div className="flex justify-between items-start gap-4 flex-wrap sm:flex-nowrap">
                        <div>
                          <h3 className="text-xl font-black">{activeStation.name}</h3>
                          <div className="flex flex-wrap items-center gap-2 mt-1.5">
                            {activeStation.section && (
                              <span className="text-xs font-bold bg-blue-500/10 text-blue-400 border border-blue-500/20 px-2.5 py-1 rounded">
                                Участок контроля: {activeStation.section}
                              </span>
                            )}
                            <p className="text-sm text-slate-400 font-medium">{activeStation.note || 'Заметка отсутствует.'}</p>
                          </div>
                        </div>
                        <div className="flex gap-1.5 shrink-0">
                          <button
                            onClick={() => setStationModal({ isOpen: true, mode: 'edit', id: activeStation.id, name: activeStation.name, section: activeStation.section || '', note: activeStation.note })}
                            className="bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-bold px-2.5 py-1.5 rounded border border-slate-700 transition"
                          >
                            Правка
                          </button>
                          <button
                            onClick={() => handleDeleteStation(activeStation.id)}
                            className="bg-rose-500/10 hover:bg-rose-500/20 text-rose-450 text-xs font-bold px-2.5 py-1.5 rounded border border-rose-500/20 transition"
                          >
                            Удалить
                          </button>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mt-4 border-t border-slate-700/60 pt-4 text-sm">
                        <div>
                          <span className="text-xs block text-slate-400 font-bold mb-1">Контролируемые ТП</span>
                          <span className="font-mono font-bold text-base">{supplyPoints.filter(p => p.stationId === activeStation.id).length} ТП</span>
                        </div>
                        <div>
                          <span className="text-xs block text-slate-400 font-bold mb-1">Расход текущий</span>
                          <span className="font-mono font-bold text-base text-blue-400">{stationActiveSum.current.toLocaleString()} кВт·ч</span>
                        </div>
                        <div>
                          <span className="text-xs block text-slate-400 font-bold mb-1">Прошлый год</span>
                          <span className="font-mono font-semibold text-base text-slate-450">{stationActiveSum.prev.toLocaleString()} кВт·ч</span>
                        </div>
                        <div>
                          <span className="text-xs block text-slate-400 font-bold mb-1">Отклонение</span>
                          <span className={`font-bold text-base ${stationActiveSum.deltaAbs > 0 ? "text-rose-450" : "text-emerald-400"}`}>
                            {stationActiveSum.deltaAbs > 0 ? `+${stationActiveSum.deltaAbs.toLocaleString()}` : stationActiveSum.deltaAbs.toLocaleString()}{' '}
                            ({stationActiveSum.deltaPct > 0 ? `+${stationActiveSum.deltaPct.toFixed(1)}%` : `${stationActiveSum.deltaPct.toFixed(1)}%`})
                          </span>
                        </div>
                      </div>
                    </div>

                    {/* Chart change view */}
                    <div className="space-y-4">
                      <div>
                        <div className="text-[10px] font-bold text-slate-400 mb-1.5 uppercase">Общий тренд расхода по станции (все ТП)</div>
                        <CustomSvgTrendChart targetId={activeStation.id} type="station" />
                      </div>

                      {activeSupplyPoint && (
                        <div className="space-y-2 border border-blue-500/20 p-4 rounded-xl bg-blue-500/[2%] transition-all">
                          <div className="flex justify-between items-center">
                            <div className="flex items-center gap-1.5">
                              <span className="w-2 h-2 rounded-full bg-blue-500 animate-pulse"></span>
                              <h4 className="text-xs font-bold uppercase text-blue-400">Тренд потребления ТП фидера: {activeSupplyPoint.name}</h4>
                            </div>
                            <button 
                              onClick={() => setSelectedSupplyPointId(null)} 
                              className="text-slate-450 hover:text-rose-400 text-xs px-1.5 py-0.5 rounded hover:bg-slate-800 transition"
                            >
                              ✕ Скрыть
                            </button>
                          </div>
                          <CustomSvgTrendChart targetId={activeSupplyPoint.id} type="supply_point" />
                        </div>
                      )}
                    </div>

                    {/* Category breakdowns */}
                    <div className="space-y-3">
                      <h4 className="text-xs font-bold uppercase text-slate-400">Расход по категориям внутри станции</h4>
                      <div className="space-y-2">
                        {stationCategoryBreakdown.map((item, idx) => (
                          <div key={idx} className="text-xs space-y-1">
                            <div className="flex justify-between">
                              <span className="font-semibold">{item.category}</span>
                              <span className="text-slate-400">{item.value.toLocaleString()} кВт·ч ({item.share.toFixed(1)}%)</span>
                            </div>
                            <div className="h-1.5 w-full bg-slate-800 rounded-full overflow-hidden">
                              <div className="h-full bg-blue-500 rounded-full" style={{ width: `${item.share}%` }}></div>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Table of points for editing values */}
                    <div className="space-y-3 pt-2">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold uppercase text-slate-450">Точки ТП станции</span>
                        <button
                          onClick={() => setPointModal({ isOpen: true, mode: 'add', name: '', stationId: activeStation.id, category: categories[0] || 'Прочие', note: '', isActive: true, calculationMethod: 'meter' })}
                          className="bg-slate-800 text-slate-200 text-[10px] px-2 py-1 rounded"
                        >
                          + Добавить ТП
                        </button>
                      </div>

                      <div className="overflow-x-auto rounded-lg border border-slate-800 shadow-sm">
                        <table className="w-full text-left text-sm border-collapse">
                          <thead className="bg-slate-800 text-slate-400">
                            <tr>
                              <th className="p-2.5">ТП фидер</th>
                              <th className="p-2.5">Группа нагрузок</th>
                              <th className="p-2.5">Способ расчета</th>
                              <th className="p-2.5 text-right">Потребление {selectedYear} (кВт·ч)</th>
                              <th className="p-2.5 text-right">Потребление {selectedYear - 1} (кВт·ч)</th>
                              <th className="p-2.5 min-w-[200px] w-64">Причины / Заметки отклонений</th>
                              <th className="p-2.5 text-center">Действия</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-800">
                            {supplyPoints.filter(p => p.stationId === activeStation.id).map(p => {
                              const reading = activeMonthReadings.find(r => r.supplyPointId === p.id);
                              const prevReading = prevYearMonthReadings.find(r => r.supplyPointId === p.id);
                              const isSelected = p.id === selectedSupplyPointId;
                              return (
                                <tr 
                                  key={p.id} 
                                  onClick={() => {
                                    if (isSelected) {
                                      setSelectedSupplyPointId(null);
                                    } else {
                                      setSelectedSupplyPointId(p.id);
                                    }
                                  }}
                                  className={`transition-all duration-150 cursor-pointer ${
                                    isSelected 
                                      ? (darkTheme ? 'bg-blue-500/10 hover:bg-blue-500/15 border-l-2 border-blue-500' : 'bg-blue-50/70 border-l-2 border-blue-500')
                                      : p.isActive 
                                        ? 'hover:bg-slate-900/60' 
                                        : 'opacity-45 h-9'
                                  }`}
                                >
                                  <td className="p-2.5 font-bold">
                                    <div className="flex items-center gap-2">
                                      <span className={`text-sm ${isSelected ? 'text-blue-400' : ''}`}>{p.name}</span>
                                      {isSelected && (
                                        <span className="text-[10px] bg-blue-500/20 text-blue-400 px-1.5 py-0.5 rounded font-normal shrink-0">Выбран</span>
                                      )}
                                    </div>
                                    <span className="text-xs text-slate-400 block mt-0.5">{p.note || 'ТП в норме'}</span>
                                  </td>
                                  <td className="p-2.5">
                                    <span className="text-xs bg-slate-900 font-bold px-2 py-0.5 rounded text-blue-400">{p.category}</span>
                                  </td>
                                  <td className="p-2.5" onClick={(e) => e.stopPropagation()}>
                                    <select
                                      value={p.calculationMethod || 'meter'}
                                      onChange={(e) => updatePointCalculationMethodInline(p.id, e.target.value as 'meter' | 'estimated')}
                                      className={`text-xs px-2 py-1 border outline-none rounded font-semibold ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-slate-50 text-slate-800 border-slate-300'}`}
                                    >
                                      <option value="meter">📈 Прибор учета</option>
                                      <option value="estimated">🧮 Расчетный способ</option>
                                    </select>
                                  </td>
                                  <td className="p-2.5 text-right" onClick={(e) => e.stopPropagation()}>
                                    <input
                                      type="text"
                                      key={`curr-${p.id}-${selectedYear}-${selectedMonth}-${reading ? reading.value : 0}`}
                                      defaultValue={reading ? reading.value : 0}
                                      onBlur={(e) => updatePointValueInline(p.id, e.target.value, selectedYear)}
                                      className={`w-24 text-right px-2 py-1.5 text-sm border outline-none font-mono font-bold rounded ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white text-slate-800 border-slate-300'}`}
                                    />
                                  </td>
                                  <td className="p-2.5 text-right" onClick={(e) => e.stopPropagation()}>
                                    <input
                                      type="text"
                                      key={`prev-${p.id}-${selectedYear - 1}-${selectedMonth}-${prevReading ? prevReading.value : 0}`}
                                      defaultValue={prevReading ? prevReading.value : 0}
                                      onBlur={(e) => updatePointValueInline(p.id, e.target.value, selectedYear - 1)}
                                      className={`w-24 text-right px-2 py-1.5 text-sm border outline-none font-mono font-bold rounded ${darkTheme ? 'bg-slate-800 border-slate-700 text-amber-500' : 'bg-white text-amber-700 border-slate-300'}`}
                                    />
                                  </td>
                                  <td className="p-2.5" onClick={(e) => e.stopPropagation()}>
                                    <textarea
                                      rows={2}
                                      placeholder="Укажите причины отклонения / замера..."
                                      defaultValue={p.note || ''}
                                      onBlur={(e) => updatePointNoteInline(p.id, e.target.value)}
                                      className={`w-full text-xs px-2 py-1.5 border outline-none rounded resize-y min-h-[42px] font-medium ${darkTheme ? 'bg-slate-850 border-slate-700 text-white placeholder-slate-500' : 'bg-slate-50 text-slate-800 placeholder-slate-400 border-slate-300'}`}
                                    />
                                  </td>
                                  <td className="p-2.5 text-center" onClick={(e) => e.stopPropagation()}>
                                    <div className="flex items-center justify-center gap-2">
                                      <button 
                                        onClick={(e) => { 
                                          e.stopPropagation(); 
                                          const currK = readings.find(r => r.supplyPointId === p.id && r.year === selectedYear && r.month === selectedMonth)?.value || 0;
                                          const prevK = readings.find(r => r.supplyPointId === p.id && r.year === selectedYear - 1 && r.month === selectedMonth)?.value || 0;
                                          setPointModal({ 
                                            isOpen: true, 
                                            mode: 'edit', 
                                            id: p.id, 
                                            name: p.name, 
                                            stationId: p.stationId, 
                                            category: p.category, 
                                            note: p.note, 
                                            isActive: p.isActive, 
                                            calculationMethod: p.calculationMethod || 'meter',
                                            currKwh: currK,
                                            prevKwh: prevK
                                          }); 
                                        }} 
                                        className="text-slate-400 hover:text-sky-400 text-xs font-bold"
                                      >
                                        Правка
                                      </button>
                                      <button onClick={(e) => { e.stopPropagation(); handleDeletePoint(p.id) }} className="text-slate-400 hover:text-rose-500 text-xs">Удалить</button>
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>

                    {/* Table of associated losses */}
                    <div className="space-y-3 pt-3 border-t border-slate-800/80">
                      <div className="flex items-center justify-between">
                        <span className="text-xs font-bold uppercase text-slate-450">Выявленные потери на станции</span>
                        <button
                          onClick={() => setLossModal({
                            isOpen: true,
                            mode: 'add',
                            name: '',
                            stationId: activeStation.id,
                            section: activeStation.section || '',
                            note: ''
                          })}
                          className="bg-blue-600 hover:bg-blue-500 text-white text-[10px] px-2.5 py-1 rounded font-bold transition-all"
                        >
                          + Добавить потери
                        </button>
                      </div>

                      {lossObjects.filter(lo => lo.stationId === activeStation.id).length === 0 ? (
                        <div className="p-4 text-center rounded-lg border border-dashed border-slate-800 text-xs text-slate-450">
                          На этой станции еще нет зарегистрированных объектов потерь. Нажмите "+ Добавить потери", чтобы добавить.
                        </div>
                      ) : (
                        <div className="overflow-x-auto rounded-lg border border-slate-800 shadow-sm">
                          <table className="w-full text-left text-xs border-collapse">
                            <thead className="bg-[#1e293b] text-slate-400">
                              <tr>
                                <th className="p-2.5">Объект потерь сети</th>
                                <th className="p-2.5">Участок контроля</th>
                                <th className="p-2.5 text-right">Потери за месяц (кВт·ч)</th>
                                <th className="p-2.5 text-center">Действия</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-800">
                              {lossObjects.filter(lo => lo.stationId === activeStation.id).map(lo => {
                                const reading = activeLossReadings.find(l => l.lossObjectId === lo.id);
                                return (
                                  <tr key={lo.id} className="hover:bg-slate-900/60">
                                    <td className="p-2.5 font-bold">
                                      <span>{lo.name}</span>
                                      <span className="text-[10px] text-slate-450 block mt-0.5">{lo.note || 'Потери сети в норме'}</span>
                                    </td>
                                    <td className="p-2.5">
                                      <span className="text-[10px] bg-amber-500/10 text-amber-405 border border-amber-500/20 px-2 py-0.5 rounded font-medium text-amber-400">
                                        {lo.section || activeStation.section || 'Не назначен'}
                                      </span>
                                    </td>
                                    <td className="p-2.5 text-right">
                                      <input
                                        type="text"
                                        defaultValue={reading ? reading.value : 0}
                                        onBlur={(e) => updateLossValueInline(lo.id, e.target.value)}
                                        className={`w-24 text-right px-2 py-1 text-xs border outline-none font-mono font-bold ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
                                      />
                                    </td>
                                    <td className="p-2.5 text-center">
                                      <div className="flex items-center justify-center gap-2">
                                        <button
                                          onClick={() => setLossModal({
                                            isOpen: true,
                                            mode: 'edit',
                                            id: lo.id,
                                            name: lo.name,
                                            stationId: lo.stationId || '',
                                            section: lo.section || '',
                                            note: lo.note
                                          })}
                                          className="text-slate-400 hover:text-sky-400 font-semibold"
                                        >
                                          Правка
                                        </button>
                                        <button
                                          onClick={() => handleDeleteLossObject(lo.id)}
                                          className="text-slate-400 hover:text-rose-500 font-semibold"
                                        >
                                          Удалить
                                        </button>
                                      </div>
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>

                  </div>
                ) : (
                  <p className="text-xs text-slate-400">Пожалуйста, выберите станцию из справочника.</p>
                )}
              </div>

            </div>
          )}

          {/* TAB 3: ПОТЕРИ (LOSS MANAGER - Section 9) */}
          {activeTab === 'losses' && (
            <div className="grid grid-cols-1 xl:grid-cols-3 gap-6 items-start">
              
              {/* Left sidebar: list of loss cards */}
              <div className="xl:col-span-1 space-y-4">
                <div className="flex items-center justify-between gap-2 border-b pb-2 border-slate-700">
                  <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400">Объекты потерь сети</h3>
                  <button 
                    onClick={() => setLossModal({ isOpen: true, mode: 'add', name: '', stationId: activeStation ? activeStation.id : (stations[0]?.id || ''), section: '', note: '' })} 
                    className="bg-blue-600 hover:bg-blue-500 text-white text-[10px] px-2 py-1 rounded font-bold"
                  >
                    + Новая
                  </button>
                </div>

                <div className="space-y-2 max-h-[500px] overflow-y-auto pr-1">
                  {lossObjects.length === 0 ? (
                    <div className="text-center py-6 text-xs text-slate-450 border border-dashed border-slate-800 rounded-lg">
                      Нет зарегистрированных объектов потерь. Нажмите "+ Новая" чтобы создать.
                    </div>
                  ) : (
                    lossObjects.map(lo => {
                      const associatedStation = stations.find(s => s.id === lo.stationId);
                      const currentSection = lo.section || associatedStation?.section || 'Не назначен';
                      
                      return (
                        <div
                          key={lo.id}
                          onClick={() => setSelectedLossId(lo.id)}
                          className={`p-3.5 rounded-lg border text-xs cursor-pointer transition-all ${
                            lo.id === selectedLossId 
                              ? 'border-blue-500 bg-blue-500/[5%]' 
                              : 'border-slate-800 bg-slate-900/40 hover:bg-slate-800/20'
                          }`}
                        >
                          <div className="flex justify-between items-start gap-1">
                            <h4 className="font-bold text-slate-200">{lo.name}</h4>
                            <span className="text-[9px] bg-slate-800 text-slate-400 px-1.5 py-0.5 rounded shrink-0 font-mono">
                              ID: {lo.id.split('-')[1] || lo.id}
                            </span>
                          </div>
                          <p className="text-[10px] text-slate-400 line-clamp-2 mt-1 leading-relaxed">{lo.note || 'Примечания отсутствуют.'}</p>
                          
                          <div className="mt-2.5 pt-2 border-t border-slate-800/80 flex flex-wrap gap-1.5 items-center">
                            <span className="text-[9px] bg-amber-500/10 text-amber-400 border border-amber-500/20 px-1.5 py-0.5 rounded font-medium">
                              Участок: {currentSection}
                            </span>
                            {associatedStation && (
                              <span className="text-[9px] bg-sky-500/10 text-sky-400 border border-sky-500/20 px-1.5 py-0.5 rounded">
                                {associatedStation.name}
                              </span>
                            )}
                          </div>
                        </div>
                      );
                    })
                  )}
                </div>
              </div>

              {/* Right content panel: Active loss profile, chart, values and inline control form */}
              <div className="xl:col-span-2 space-y-6">
                {activeLossObject ? (
                  <div className="space-y-6">
                    
                    {/* Active loss card profile details */}
                    <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-[#1e293b]/55 border-slate-700' : 'bg-white text-slate-850'}`}>
                      <div className="flex justify-between items-start gap-4">
                        <div>
                          <h3 className="text-base font-black text-white">{activeLossObject.name}</h3>
                          <div className="flex flex-wrap items-center gap-2 mt-1.5">
                            <span className="text-[10px] font-bold bg-amber-550/15 text-amber-500 border border-amber-500/30 px-2 py-0.5 text-amber-400 rounded">
                              Участок контроля: {activeLossObject.section || (stations.find(s => s.id === activeLossObject.stationId)?.section) || 'Не заполнен'}
                            </span>
                            {activeLossObject.stationId && (
                              <span className="text-[10px] font-bold bg-blue-500/15 text-blue-400 border border-blue-500/30 px-2 py-0.5 text-blue-400 rounded font-bold">
                                Станция: {stations.find(s => s.id === activeLossObject.stationId)?.name || 'Неизвестная'}
                              </span>
                            )}
                          </div>
                          <p className="text-xs text-slate-400 mt-2 italic">{activeLossObject.note || 'Инструкции для узла учета не заданы.'}</p>
                        </div>
                        
                        <div className="flex gap-2 shrink-0">
                          <button 
                            onClick={() => setLossModal({ 
                              isOpen: true, 
                              mode: 'edit', 
                              id: activeLossObject.id, 
                              name: activeLossObject.name, 
                              stationId: activeLossObject.stationId || '', 
                              section: activeLossObject.section || '', 
                              note: activeLossObject.note 
                            })} 
                            className="bg-slate-800 hover:bg-slate-700 text-slate-300 text-[10px] font-bold px-2.5 py-1.5 rounded border border-slate-700 font-mono transition-all"
                          >
                            Правка
                          </button>
                          <button 
                            onClick={() => handleDeleteLossObject(activeLossObject.id)}
                            className="bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 text-[10px] font-bold px-2.5 py-1.5 rounded border border-rose-500/20 transition-all font-mono"
                          >
                            Удалить
                          </button>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 md:grid-cols-3 gap-4 mt-5 border-t border-slate-700/60 pt-4">
                        <div>
                          <span className="text-[10px] text-slate-400 block font-semibold uppercase tracking-wider">Потери в выбранном месяце</span>
                          <strong className="font-mono font-extrabold text-sm block mt-0.5 text-blue-400">
                            {(activeLossReadings.find(l => l.lossObjectId === activeLossObject.id)?.value || 0).toLocaleString()} кВт·ч
                          </strong>
                        </div>
                        <div>
                          <span className="text-[10px] text-slate-400 block font-semibold uppercase tracking-wider">Аналогичный месяц прошлого года</span>
                          <strong className="font-mono font-extrabold text-sm block mt-0.5 text-slate-350">
                            {(prevYearLossReadings.find(l => l.lossObjectId === activeLossObject.id)?.value || 0).toLocaleString()} кВт·ч
                          </strong>
                        </div>
                        <div className="col-span-2 md:col-span-1">
                          <span className="text-[10px] text-slate-400 block font-semibold uppercase tracking-wider">Отклонение</span>
                          {(() => {
                            const cur = activeLossReadings.find(l => l.lossObjectId === activeLossObject.id)?.value || 0;
                            const prev = prevYearLossReadings.find(l => l.lossObjectId === activeLossObject.id)?.value || 0;
                            const diff = cur - prev;
                            const pct = prev > 0 ? (diff / prev) * 100 : 0;
                            return (
                              <strong className={`font-mono text-sm block mt-0.5 font-bold ${diff > 0 ? 'text-rose-450' : 'text-emerald-400'}`}>
                                {diff > 0 ? `+${diff.toLocaleString()}` : diff.toLocaleString()} ({pct > 0 ? `+${pct.toFixed(1)}%` : `${pct.toFixed(1)}%`})
                              </strong>
                            );
                          })()}
                        </div>
                      </div>
                    </div>

                    {/* Trend line SVG visualization */}
                    <div className="bg-slate-900/30 p-4 rounded-xl border border-slate-800">
                      <CustomSvgTrendChart targetId={activeLossObject.id} type="loss" />
                    </div>

                    {/* Inline direct value editor & "Участок контроля" direct form */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                      
                      {/* Left half: value input */}
                      <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700/80' : 'bg-white border-slate-200'} space-y-2`}>
                        <label className="text-xs block font-bold text-slate-300">Ввод показаний текущих потерь (кВт·ч):</label>
                        <div className="flex items-center gap-2 pt-1">
                          <input
                            type="text"
                            value={activeLossReadings.find(l => l.lossObjectId === activeLossObject.id)?.value || ''}
                            onChange={(e) => updateLossValueInline(activeLossObject.id, e.target.value)}
                            placeholder="0"
                            className={`w-32 text-right px-3 py-2 rounded text-xs border outline-none font-mono font-bold ${darkTheme ? 'bg-slate-900 border-slate-700 text-white' : 'bg-white'}`}
                          />
                          <span className="text-xs text-slate-400 font-semibold">кВт·ч за {selectedPeriodName} {selectedYear} года</span>
                        </div>
                        <p className="text-[10px] text-slate-450">Значение сохранится автоматически при изменении.</p>
                      </div>

                      {/* Right half: "Участок контроля" Direct Card Form */}
                      <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700/80' : 'bg-white border-slate-200'} space-y-2`}>
                        <span className="text-xs block font-bold text-slate-300">Форма: Участок контроля</span>
                        
                        <div className="space-y-2 pt-1 text-xs">
                          <div>
                            <input
                              type="text"
                              value={activeLossObject.section || ''}
                              onChange={(e) => {
                                const val = e.target.value;
                                const updated = lossObjects.map(lo => lo.id === activeLossObject.id ? { ...lo, section: val } : lo);
                                setLossObjects(updated);
                                saveToStorage(stations, supplyPoints, categories, readings, updated, lossReadings);
                              }}
                              placeholder="Имя диспетчерского участка контроля..."
                              className={`w-full px-2.5 py-1.5 rounded border outline-none text-xs ${darkTheme ? 'bg-slate-900 border-slate-700 text-white' : 'bg-white'}`}
                            />
                          </div>

                          <div className="flex items-center gap-1.5 pt-0.5">
                            <span className="text-[10px] text-slate-450">Текущий участок:</span>
                            <strong className="text-[10px] text-amber-400 uppercase tracking-wide font-mono">
                              {activeLossObject.section || (stations.find(s => s.id === activeLossObject.stationId)?.section) || 'Не заполнен'}
                            </strong>
                          </div>
                        </div>
                      </div>

                    </div>

                  </div>
                ) : (
                  <div className="text-center py-10 text-xs text-slate-400 bg-slate-900/20 rounded-xl border border-slate-800">
                    Пожалуйста, создайте или выберите объект технологических потерь слева.
                  </div>
                )}
              </div>

            </div>
          )}

          {/* TAB 4: АНАЛИТИКА (DRILLDOWN MATRIX - Section 11) */}
          {activeTab === 'analytics' && (() => {
            const activeSps = supplyPoints.filter(p => p.isActive);
            const spsMap = new Map(activeSps.map(p => [p.id, p]));
            
            let meterPointsCount = 0;
            let estimatedPointsCount = 0;
            let meterTotalKwh = 0;
            let estimatedTotalKwh = 0;
            
            activeSps.forEach(p => {
              if (p.calculationMethod === 'estimated') {
                estimatedPointsCount++;
              } else {
                meterPointsCount++;
              }
            });
            
            const curMonthReadings = readings.filter(rd => rd.year === selectedYear && selectedMonthsList.includes(rd.month));
            curMonthReadings.forEach(rd => {
              const p = spsMap.get(rd.supplyPointId);
              if (p) {
                if (p.calculationMethod === 'estimated') {
                  estimatedTotalKwh += rd.value;
                } else {
                  meterTotalKwh += rd.value;
                }
              }
            });
            
            const grandTotalKwh = meterTotalKwh + estimatedTotalKwh;
            const grandTotalPoints = meterPointsCount + estimatedPointsCount;
            
            const meterKwhShare = grandTotalKwh > 0 ? (meterTotalKwh / grandTotalKwh) * 100 : 0;
            const estimatedKwhShare = grandTotalKwh > 0 ? (estimatedTotalKwh / grandTotalKwh) * 100 : 0;
            
            const meterPointsShare = grandTotalPoints > 0 ? (meterPointsCount / grandTotalPoints) * 100 : 0;
            const estimatedPointsShare = grandTotalPoints > 0 ? (estimatedPointsCount / grandTotalPoints) * 100 : 0;

            const stationCalculations = stations.map(s => {
              const stationSps = activeSps.filter(p => p.stationId === s.id);
              const stationSpsIds = stationSps.map(p => p.id);
              const stationSpsMap = new Map(stationSps.map(p => [p.id, p]));
              
              let sMeterVal = 0;
              let sEstimatedVal = 0;
              
              const sReadings = readings.filter(rd => rd.year === selectedYear && selectedMonthsList.includes(rd.month) && stationSpsIds.includes(rd.supplyPointId));
              sReadings.forEach(rd => {
                const p = stationSpsMap.get(rd.supplyPointId);
                if (p) {
                  if (p.calculationMethod === 'estimated') {
                    sEstimatedVal += rd.value;
                  } else {
                    sMeterVal += rd.value;
                  }
                }
              });
              
              const sTotal = sMeterVal + sEstimatedVal;
              return {
                stationId: s.id,
                stationName: s.name,
                meterValue: sMeterVal,
                estimatedValue: sEstimatedVal,
                totalValue: sTotal,
                estimatedShare: sTotal > 0 ? (sEstimatedVal / sTotal) * 100 : 0,
                meterShare: sTotal > 0 ? (sMeterVal / sTotal) * 100 : 0,
              };
            });

            return (
              <div className="space-y-6">
                
                {/* Метрики способов расчета */}
                <div className={`p-5 rounded-2xl border ${darkTheme ? 'bg-slate-900 border-slate-800 animate-fadeIn' : 'bg-white border-slate-200'} shadow-md space-y-4`}>
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-2 border-b pb-3 border-slate-800/80">
                    <div>
                      <h3 className="font-extrabold text-sm flex items-center gap-1.5">
                        <Activity className="w-4 h-4 text-sky-400" />
                        Анализ структуры способов расчета энергии
                      </h3>
                      <p className="text-[10px] text-slate-400">Соотношение коммерческого учета по приборам и оценочного расчетного способа</p>
                    </div>
                    <span className="text-xs bg-slate-850 px-2.5 py-1 rounded font-semibold text-slate-350">
                      Период: {selectedPeriodName} {selectedYear}
                    </span>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    {/* Приборы учета */}
                    <div className="p-4 rounded-xl bg-gradient-to-br from-blue-500/5 to-sky-500/10 border border-blue-500/10 space-y-2">
                      <div className="flex justify-between items-center text-xs">
                        <span className="text-slate-400 uppercase font-bold text-[10px]">По приборам учета</span>
                        <span className="bg-blue-500/20 text-blue-400 px-2 py-0.5 rounded-full text-[10px] font-bold">🎯 Точный учет</span>
                      </div>
                      <div className="space-y-1">
                        <span className="text-2xl font-black font-mono block text-blue-405">
                          {Math.round(meterTotalKwh).toLocaleString()} <span className="text-xs font-normal">кВт·ч</span>
                        </span>
                        <span className="text-xs text-slate-400 block">
                          Доля нагрузки: <strong className="text-slate-200 font-mono">{meterKwhShare.toFixed(1)}%</strong>
                        </span>
                        <span className="text-xs text-slate-400 block">
                          Активных ТП: <strong className="text-slate-200 font-mono">{meterPointsCount}</strong> ТП ({meterPointsShare.toFixed(0)}%)
                        </span>
                      </div>
                    </div>

                    {/* Расчетный способ */}
                    <div className="p-4 rounded-xl bg-gradient-to-br from-amber-500/5 to-orange-500/10 border border-amber-500/10 space-y-2">
                      <div className="flex justify-between items-center text-xs">
                        <span className="text-slate-400 uppercase font-bold text-[10px]">Расчетным способом</span>
                        <span className="bg-amber-500/20 text-amber-400 px-2 py-0.5 rounded-full text-[10px] font-bold">⚠️ Оценка</span>
                      </div>
                      <div className="space-y-1">
                        <span className="text-2xl font-black font-mono block text-amber-400">
                          {Math.round(estimatedTotalKwh).toLocaleString()} <span className="text-xs font-normal">кВт·ч</span>
                        </span>
                        <span className="text-xs text-slate-400 block">
                          Доля нагрузки: <strong className="text-slate-200 font-mono">{estimatedKwhShare.toFixed(1)}%</strong>
                        </span>
                        <span className="text-xs text-slate-400 block">
                          Активных ТП: <strong className="text-slate-200 font-mono">{estimatedPointsCount}</strong> ТП ({estimatedPointsShare.toFixed(0)}%)
                        </span>
                      </div>
                    </div>

                    {/* Общее по сети */}
                    <div className="p-4 rounded-xl bg-slate-850 border border-slate-700/50 space-y-2 md:col-span-2 lg:col-span-1">
                      <div className="flex justify-between items-center text-xs">
                        <span className="text-slate-400 uppercase font-bold text-[10px]">Итого по станциям</span>
                        <span className="text-slate-400 text-[10px] font-mono font-bold">Общая сеть</span>
                      </div>
                      <div className="space-y-1">
                        <span className="text-2xl font-black font-mono block text-slate-200">
                          {Math.round(grandTotalKwh).toLocaleString()} <span className="text-xs font-normal">кВт·ч</span>
                        </span>
                        <span className="text-xs text-slate-400 block">
                          Всего на балансе: <strong className="text-slate-200 font-mono">{grandTotalPoints}</strong> точек ТП
                        </span>
                        <span className="text-xs text-slate-400 block">
                          Уровень оцифровки приборами: <strong className="text-emerald-400 font-mono">{meterPointsShare.toFixed(1)}%</strong>
                        </span>
                      </div>
                    </div>
                  </div>

                  {/* Визуальная шкала пропорций */}
                  <div className="space-y-1.5 pt-1">
                    <div className="flex justify-between text-xs font-bold">
                      <span className="text-blue-400 flex items-center gap-1">📈 По приборам ({meterKwhShare.toFixed(1)}%)</span>
                      <span className="text-amber-400 flex items-center gap-1">🧮 По расчету ({estimatedKwhShare.toFixed(1)}%)</span>
                    </div>
                    <div className="w-full flex h-3.5 bg-slate-800 rounded-lg overflow-hidden border border-slate-700/60 p-0.5">
                      <div className="h-full bg-blue-500 rounded-full transition-all duration-500" style={{ width: `${meterKwhShare}%` }} />
                      <div className="h-full bg-amber-550 rounded-full transition-all duration-500 ml-0.5" style={{ width: `${estimatedKwhShare}%` }} />
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
                  
                  {/* Техническое потребление по станциям */}
                  <div className={`p-5 rounded-2xl border ${darkTheme ? 'bg-slate-900 border-slate-800' : 'bg-white'} space-y-4`}>
                    <span className="text-xs font-black uppercase text-slate-450 block">Техническое потребление по станциям</span>
                    <div className="space-y-2 text-xs">
                      {stations.map(s => {
                        const spsIds = supplyPoints.filter(p => p.stationId === s.id && p.isActive).map(p => p.id);
                        const yrRds = readings.filter(rd => rd.year === selectedYear && spsIds.includes(rd.supplyPointId));
                        const tot = yrRds.reduce((sum, r) => sum + r.value, 0);
                        const avg = tot / (new Set(yrRds.map(r => r.month)).size || 1);

                        return (
                          <div key={s.id} className="flex justify-between items-center py-2 border-b border-slate-800/55 last:border-0">
                            <span className="font-bold">{s.name}</span>
                            <span className="font-mono text-slate-300">Средн. расход: <strong className={darkTheme ? "text-white" : "text-slate-800"}>{Math.round(avg).toLocaleString()}</strong> кВт·ч</span>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Доля Оценок (Расчетных способов) по Станциям */}
                  <div className={`p-5 rounded-2xl border ${darkTheme ? 'bg-slate-900 border-slate-800' : 'bg-white'} space-y-4`}>
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-black uppercase text-slate-450 block">Индикатор незамеряемых линий (Расчет)</span>
                      <span className="text-[10px] text-amber-500 font-bold bg-amber-500/10 px-2 py-0.5 rounded-full">Требует дооснащения</span>
                    </div>
                    <div className="space-y-3.5 text-xs">
                      {stationCalculations.map(sc => {
                        return (
                          <div key={sc.stationId} className="space-y-1.5">
                            <div className="flex justify-between items-center">
                              <span className="font-bold">{sc.stationName}</span>
                              <div className="flex items-center gap-2 font-mono">
                                <span className="text-slate-400 font-medium">Расчет:</span>
                                <span className={`font-bold ${sc.estimatedShare > 30 ? 'text-rose-450 font-extrabold' : sc.estimatedShare > 0 ? 'text-amber-400' : 'text-slate-500'}`}>
                                  {Math.round(sc.estimatedValue).toLocaleString()} кВт·ч ({sc.estimatedShare.toFixed(1)}%)
                                </span>
                              </div>
                            </div>
                            <div className="w-full h-2 bg-slate-800 rounded-full overflow-hidden flex">
                              <div className="bg-blue-500 h-full" style={{ width: `${sc.meterShare}%` }}></div>
                              <div className="bg-amber-550 h-full animate-pulse" style={{ width: `${sc.estimatedShare}%` }}></div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Влияние по категориям нагрузок */}
                  <div className={`p-5 rounded-2xl border ${darkTheme ? 'bg-slate-900 border-slate-800' : 'bg-white'} space-y-4 lg:col-span-2`}>
                    <span className="text-xs font-black uppercase text-slate-455 block">Влияние по категориям нагрузок на баланс</span>
                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
                      {categories.slice(0, 12).map(c => {
                        const item = globalCategoryShares[c] || { value: 0, share: 0 };
                        return (
                          <div key={c} className="p-3 bg-slate-850/40 border border-slate-800 rounded-xl flex flex-col justify-between space-y-1.5 h-16">
                            <span className="text-slate-400 text-[10px] uppercase font-bold truncate animate-fadeIn" title={c}>{c}</span>
                            <span className="font-mono text-indigo-400 font-extrabold text-sm text-left">{item.share.toFixed(1)}% доля</span>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                </div>

              </div>
            );
          })()}

          {/* TAB 5: ОТЧЕТЫ (REPORTS - Section 13) */}
          {activeTab === 'reports' && (
            <div className="space-y-6">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b pb-3 border-slate-700/60">
                <div>
                  <h3 className="font-extrabold text-sm">Генерация сводных отчетов</h3>
                  <p className="text-[10px] text-slate-400">Учетный период: {selectedPeriodName} {selectedYear} года</p>
                </div>
                
                <div className="flex flex-wrap items-center gap-3">
                  {/* Переключатель вкладок отчета */}
                  <div className={`p-1 rounded-lg flex items-center border ${darkTheme ? 'bg-slate-900 border-slate-800' : 'bg-slate-100 border-slate-200'}`}>
                    <button
                      onClick={() => setReportSubTab('stations')}
                      className={`px-3 py-1 text-xs font-bold rounded-md transition-all duration-150 ${
                        reportSubTab === 'stations'
                          ? 'bg-blue-600 text-white shadow'
                          : `${darkTheme ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900'}`
                      }`}
                    >
                      Ведомости станций
                    </button>
                    <button
                      onClick={() => {
                        setReportSubTab('categories');
                        if (!selectedReportCategoryId && categories.length > 0) {
                          // Default to some known category
                          const defCat = categories.includes("Освещение горловин") 
                            ? "Освещение горловин" 
                            : categories[0];
                          setSelectedReportCategoryId(defCat);
                        }
                      }}
                      className={`px-3 py-1 text-xs font-bold rounded-md flex items-center gap-1.5 transition-all duration-150 ${
                        reportSubTab === 'categories'
                          ? 'bg-blue-600 text-white shadow'
                          : `${darkTheme ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900'}`
                      }`}
                    >
                      <Tag className="w-3 h-3" />
                      Аналитика по категориям
                    </button>
                    <button
                      onClick={() => setReportSubTab('accounting_methods')}
                      className={`px-3 py-1 text-xs font-bold rounded-md flex items-center gap-1.5 transition-all duration-150 ${
                        reportSubTab === 'accounting_methods'
                          ? 'bg-blue-600 text-white shadow'
                          : `${darkTheme ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900'}`
                      }`}
                    >
                      <Activity className="w-3.5 h-3.5" />
                      Способы расчета
                    </button>
                    <button
                      onClick={() => setReportSubTab('losses')}
                      className={`px-3 py-1 text-xs font-bold rounded-md flex items-center gap-1.5 transition-all duration-150 ${
                        reportSubTab === 'losses'
                          ? 'bg-blue-600 text-white shadow'
                          : `${darkTheme ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900'}`
                      }`}
                    >
                      <ZapOff className="w-3.5 h-3.5 text-rose-500" />
                      Потери сети
                    </button>
                    <button
                      onClick={() => setReportSubTab('slides_presentation')}
                      className={`px-3 py-1 text-xs font-bold rounded-md flex items-center gap-1.5 transition-all duration-150 ${
                        reportSubTab === 'slides_presentation'
                          ? 'bg-blue-600 text-white shadow'
                          : `${darkTheme ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900'}`
                      }`}
                    >
                      <Presentation className="w-3.5 h-3.5 text-amber-500" strokeWidth={2.5} />
                      Слайд-презентация
                    </button>
                    <button
                      onClick={() => setReportSubTab('detailed_analysis')}
                      className={`px-3 py-1 text-xs font-bold rounded-md flex items-center gap-1.5 transition-all duration-150 ${
                        reportSubTab === 'detailed_analysis'
                          ? 'bg-blue-600 text-white shadow'
                          : `${darkTheme ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900'}`
                      }`}
                    >
                      <TrendingUp className="w-3.5 h-3.5 text-emerald-500 animate-pulse" />
                      Подробный анализ
                    </button>
                  </div>

                  <button
                    onClick={reportSubTab === 'detailed_analysis' ? exportDetailedReportToDoc : exportReportToDoc}
                    className="bg-emerald-600 hover:bg-emerald-700 px-3.5 py-1.5 text-xs text-white rounded font-bold transition-all duration-150 flex items-center gap-1.5 cursor-pointer shadow-sm hover:shadow"
                  >
                    <FileText className="w-3.5 h-3.5" />
                    <span>{reportSubTab === 'detailed_analysis' ? 'Скачать сводный подробный отчет (.doc)' : 'Скачать отчет Word (.doc)'}</span>
                  </button>

                  <button
                    onClick={exportReportToPptx}
                    className="bg-amber-600 hover:bg-amber-700 px-3.5 py-1.5 text-xs text-white rounded font-bold transition-all duration-150 flex items-center gap-1.5 cursor-pointer shadow-sm hover:shadow"
                  >
                    <Presentation className="w-3.5 h-3.5" />
                    <span>Скачать слайды PPTX</span>
                  </button>

                  <button onClick={() => window.print()} className="bg-blue-600 px-3.5 py-1.5 text-xs text-white rounded font-bold hover:bg-blue-700 transition-colors">
                    Распечатать в PDF
                  </button>
                </div>
              </div>

              {reportSubTab === 'stations' && (
                <div className="grid grid-cols-1 xl:grid-cols-12 gap-6 items-start">
                  {/* Левая колонка: Ведомость станций */}
                  <div className={`xl:col-span-12 xl:col-span-7 p-6 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                    <div className="text-center space-y-1 mb-5">
                      <h2 className="text-xs uppercase tracking-widest text-slate-400 font-extrabold">Сводная энергетическая ведомость станций железной дороги</h2>
                      <p className="text-[10px] text-slate-400">Нажмите на любую станцию ниже для подробного анализа по точкам и графикам сравнения г/г</p>
                    </div>

                    {/* Интерактивный фильтр способов расчета */}
                    <div className={`flex flex-col sm:flex-row items-center justify-between gap-3 mb-5 p-3 rounded-lg border ${darkTheme ? 'bg-slate-900/40 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
                      <div className="flex items-center gap-1.5 text-[10px] uppercase font-bold tracking-wider text-slate-400">
                        <Filter className="w-3.5 h-3.5 text-blue-500 animate-pulse" />
                        <span>Способ расчета расхода:</span>
                      </div>
                      <div className={`flex rounded-lg p-0.5 border text-[10px] font-bold ${darkTheme ? 'bg-slate-950 border-slate-800' : 'bg-white border-slate-200'}`}>
                        <button
                          onClick={() => setReportCalculationMethodFilter('all')}
                          className={`px-3 py-1 rounded-md transition-all duration-150 ${
                            reportCalculationMethodFilter === 'all'
                              ? 'bg-blue-600 text-white shadow-sm'
                              : `${darkTheme ? 'text-slate-450 hover:text-white' : 'text-slate-600 hover:text-slate-950'}`
                          }`}
                        >
                          📈 Все способы
                        </button>
                        <button
                          onClick={() => setReportCalculationMethodFilter('meter')}
                          className={`px-3 py-1 rounded-md transition-all duration-150 ${
                            reportCalculationMethodFilter === 'meter'
                              ? 'bg-blue-600 text-white shadow-sm'
                              : `${darkTheme ? 'text-slate-450 hover:text-white' : 'text-slate-600 hover:text-slate-950'}`
                          }`}
                        >
                          🎯 Приборы учета
                        </button>
                        <button
                          onClick={() => setReportCalculationMethodFilter('estimated')}
                          className={`px-3 py-1 rounded-md transition-all duration-150 ${
                            reportCalculationMethodFilter === 'estimated'
                              ? 'bg-blue-600 text-white shadow-sm'
                              : `${darkTheme ? 'text-slate-450 hover:text-white' : 'text-slate-600 hover:text-slate-950'}`
                          }`}
                        >
                          🧮 Расчетный способ
                        </button>
                      </div>
                    </div>

                    <div className="overflow-x-auto animate-fadeIn">
                      <table className="w-full text-left text-xs">
                        <thead className={`font-bold ${darkTheme ? 'bg-slate-800 text-slate-400' : 'bg-slate-100 text-slate-600'}`}>
                          <tr>
                            <th className="p-2.5">Станция</th>
                            <th className="p-2.5 text-right font-semibold">Показания (кВт·ч)</th>
                            <th className="p-2.5 text-right font-semibold">Прошлый год (кВт·ч)</th>
                            <th className="p-2.5 text-right font-semibold">Разница</th>
                            <th className="p-2.5 text-right font-semibold">Разница (%)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-800/60">
                          {stations.map(st => {
                            // Filter points for station based on the calculation method filter
                            const rawPoints = supplyPoints.filter(p => p.stationId === st.id && p.isActive);
                            const filteredPoints = rawPoints.filter(p => {
                              if (reportCalculationMethodFilter === 'meter') return p.calculationMethod !== 'estimated';
                              if (reportCalculationMethodFilter === 'estimated') return p.calculationMethod === 'estimated';
                              return true;
                            });
                            const spsIds = filteredPoints.map(p => p.id);

                            const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                            const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                            const diff = curr - prev;
                            const pct = prev > 0 ? (diff / prev) * 100 : 0;
                            const isSelected = selectedReportStationId === st.id;

                            // Calculate breakdown for visual indicators under station name
                            const stationMeterPoints = rawPoints.filter(p => p.calculationMethod !== 'estimated');
                            const stationEstimatedPoints = rawPoints.filter(p => p.calculationMethod === 'estimated');
                            const meterKwh = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && stationMeterPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                            const estimatedKwh = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && stationEstimatedPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                            const totalKwh = meterKwh + estimatedKwh;
                            const meterPct = totalKwh > 0 ? (meterKwh / totalKwh) * 100 : 0;
                            const estimatedPct = totalKwh > 0 ? (estimatedKwh / totalKwh) * 100 : 0;

                            return (
                              <tr
                                key={st.id}
                                onClick={() => {
                                  setSelectedReportStationId(isSelected ? null : st.id);
                                  setSelectedReportSupplyPointId(null);
                                }}
                                className={`cursor-pointer transition-all duration-150 ${
                                  isSelected
                                    ? (darkTheme ? "bg-blue-600/15 text-blue-100 font-bold border-l-2 border-blue-500" : "bg-blue-50 text-blue-900 font-bold border-l-2 border-blue-600")
                                    : (darkTheme ? "hover:bg-slate-700/30" : "hover:bg-slate-50")
                                }`}
                              >
                                <td className="p-2.5 font-bold">
                                  <div className="flex items-start gap-2">
                                    <span className={`transition-transform duration-150 mt-0.5 ${isSelected ? 'rotate-90 text-blue-500' : 'text-slate-500'}`}>
                                      <ChevronRight className="w-3.5 h-3.5" />
                                    </span>
                                    <div>
                                      <span className={`${darkTheme ? 'text-slate-200' : 'text-slate-800'}`}>{st.name}</span>
                                      
                                      {/* Постоянная наглядная расшифровка способов учета по станции */}
                                      <div className="flex flex-col gap-0.5 mt-1.5 font-normal text-[9px] text-slate-400">
                                        <div className="flex items-center gap-1.5">
                                          <span className="inline-block w-1.5 h-1.5 rounded-full bg-blue-500"></span>
                                          <span>Приборы учета: <strong className={darkTheme ? "text-slate-300" : "text-slate-700"}>{Math.round(meterKwh).toLocaleString()} кВт·ч</strong> ({meterPct.toFixed(0)}%)</span>
                                        </div>
                                        <div className="flex items-center gap-1.5">
                                          <span className="inline-block w-1.5 h-1.5 rounded-full bg-amber-500"></span>
                                          <span>Расчетный способ: <strong className={darkTheme ? "text-slate-300" : "text-slate-700"}>{Math.round(estimatedKwh).toLocaleString()} кВт·ч</strong> ({estimatedPct.toFixed(0)}%)</span>
                                        </div>
                                        <div className={`w-28 h-1 rounded-full overflow-hidden flex mt-1 ${darkTheme ? 'bg-slate-800' : 'bg-slate-200'}`}>
                                          <div className="bg-blue-500 h-full" style={{ width: `${meterPct}%` }} title={`Приборы учета: ${meterPct.toFixed(0)}%`} />
                                          <div className="bg-amber-500 h-full" style={{ width: `${estimatedPct}%` }} title={`Расчетный способ: ${estimatedPct.toFixed(0)}%`} />
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                </td>
                                <td className="p-2.5 text-right font-mono font-medium">{curr.toLocaleString()}</td>
                                <td className="p-2.5 text-right font-mono text-slate-400">{prev.toLocaleString()}</td>
                                <td className={`p-2.5 text-right font-mono font-extrabold ${diff > 0 ? "text-rose-500" : diff < 0 ? "text-emerald-500" : "text-slate-400"}`}>
                                  {diff > 0 ? `+${diff.toLocaleString()}` : diff.toLocaleString()}
                                </td>
                                <td className={`p-2.5 text-right font-mono font-extrabold ${diff > 0 ? "text-rose-500" : diff < 0 ? "text-emerald-500" : "text-slate-400"}`}>
                                  {prev > 0 ? (diff > 0 ? `+${pct.toFixed(2)}%` : `${pct.toFixed(2)}%`) : (curr > 0 ? "+100%" : "0.00%")}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Правая колонка: Детализация станций и графики сравнения */}
                  {selectedReportStationId ? (() => {
                    const selectedReportStation = stations.find(st => st.id === selectedReportStationId);
                    if (!selectedReportStation) return null;

                    const reportSupplyPoints = supplyPoints.filter(p => {
                      if (p.stationId !== selectedReportStation.id || !p.isActive) return false;
                      if (reportCalculationMethodFilter === 'meter') return p.calculationMethod !== 'estimated';
                      if (reportCalculationMethodFilter === 'estimated') return p.calculationMethod === 'estimated';
                      return true;
                    });
                    
                    // Считаем помесячные показатели для текущего года и прошлого года
                    const reportSpIds = selectedReportSupplyPointId 
                      ? [selectedReportSupplyPointId] 
                      : reportSupplyPoints.map(p => p.id);

                    const monthlyDataCurr = Array.from({ length: 12 }, (_, idx) => {
                      const m = idx + 1;
                      return readings.filter(r => r.year === selectedYear && r.month === m && reportSpIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    });

                    const monthlyDataPrev = Array.from({ length: 12 }, (_, idx) => {
                      const m = idx + 1;
                      return readings.filter(r => r.year === selectedYear - 1 && r.month === m && reportSpIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    });

                    const maxReportVal = Math.max(...monthlyDataCurr, ...monthlyDataPrev, 100);
                    const reportMaxLimit = Math.ceil(maxReportVal * 1.15);

                    const rpPad = 40;
                    const rpH = 220;
                    const rpW = 540;
                    const rpInnerH = rpH - rpPad * 2;
                    const rpInnerW = rpW - rpPad * 2;

                    const rpXOf = (mIdx: number) => rpPad + (mIdx / 11) * rpInnerW;
                    const rpYOf = (v: number) => rpH - rpPad - (v / reportMaxLimit) * rpInnerH;

                    return (
                      <div className="xl:col-span-12 xl:col-span-5 space-y-6 animate-fadeIn">
                        {/* Карточка с деталями по точкам поставки */}
                        <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                          <div className="flex items-center justify-between mb-3.5">
                            <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400">
                              Точки поставки: <span className={darkTheme ? "text-white" : "text-slate-800"}>{selectedReportStation.name}</span>
                            </h4>
                            {selectedReportSupplyPointId ? (
                              <button 
                                onClick={() => setSelectedReportSupplyPointId(null)}
                                className="text-[9px] bg-red-500/10 text-red-500 hover:bg-red-500/20 px-2 py-0.5 rounded-full font-bold flex items-center gap-1 transition-all duration-150"
                              >
                                <span>Показать все</span>
                                <X className="w-2.5 h-2.5" />
                              </button>
                            ) : (
                              <span className="text-[9px] bg-blue-500/10 text-blue-400 px-2 py-0.5 rounded-full font-bold">
                                {reportSupplyPoints.length} ТП
                              </span>
                            )}
                          </div>

                          <div className="overflow-x-auto max-h-[220px] overflow-y-auto">
                            <table className="w-full text-left text-xs">
                              <thead className={`sticky top-0 font-bold ${darkTheme ? 'bg-slate-900 text-slate-400' : 'bg-slate-100 text-slate-700'}`}>
                                <tr>
                                  <th className="p-2">Название ТП</th>
                                  <th className="p-2 text-right">Текущий (кВт·ч)</th>
                                  <th className="p-2 text-right">Прошлый г.</th>
                                  <th className="p-2 text-right">Разн.</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-800/40">
                                {reportSupplyPoints.map(p => {
                                  const currPt = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
                                  const prevPt = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
                                  const diffPt = currPt - prevPt;
                                  const isPtSelected = selectedReportSupplyPointId === p.id;

                                  return (
                                    <tr 
                                      key={p.id} 
                                      onClick={() => setSelectedReportSupplyPointId(isPtSelected ? null : p.id)}
                                      className={`cursor-pointer transition-all duration-150 ${
                                        isPtSelected
                                          ? (darkTheme ? "bg-amber-500/15 text-amber-200 font-bold border-l-2 border-amber-500" : "bg-amber-50 text-amber-900 font-bold border-l-2 border-amber-600")
                                          : (darkTheme ? "hover:bg-slate-800/30" : "hover:bg-slate-50")
                                      }`}
                                    >
                                      <td className="p-2 font-medium text-[11px] truncate max-w-[145px] flex items-center gap-1.5" title={p.name}>
                                        <span className={`w-1.5 h-1.5 rounded-full transition-all duration-150 ${isPtSelected ? 'bg-amber-500 scale-125' : 'bg-slate-400'}`}></span>
                                        <div className="flex flex-col truncate">
                                          <span className="truncate">{p.name}</span>
                                          <span className={`text-[9px] uppercase font-bold shrink-0 mt-0.5 ${p.calculationMethod === 'estimated' ? 'text-amber-500' : 'text-blue-400'}`}>
                                            {p.calculationMethod === 'estimated' ? 'Расчетный способ' : 'Прибор учета'}
                                          </span>
                                        </div>
                                      </td>
                                      <td className="p-2 text-right font-mono">{currPt.toLocaleString()}</td>
                                      <td className="p-2 text-right font-mono text-slate-400 text-[11px]">{prevPt.toLocaleString()}</td>
                                      <td className={`p-2 text-right font-mono font-bold text-[11px] ${diffPt > 0 ? "text-rose-500" : diffPt < 0 ? "text-emerald-500" : "text-slate-400"}`}>
                                        {diffPt > 0 ? `+${diffPt.toLocaleString()}` : diffPt.toLocaleString()}
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Тренды и сравнение по месяцам (г/г) */}
                        <div className={`p-5 rounded-xl border relative ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                          <div className="flex items-center justify-between mb-2">
                            <div>
                              <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400">
                                {selectedReportSupplyPointId ? "Динамика ТП по месяцам" : "Динамика по месяцам"} ({selectedYear} vs {selectedYear - 1})
                              </h4>
                              <p className="text-[9px] text-slate-500">
                                {selectedReportSupplyPointId 
                                  ? `Точка поставки: "${reportSupplyPoints.find(p => p.id === selectedReportSupplyPointId)?.name}"` 
                                  : "Потребление активными точками станции в кВт·ч"}
                              </p>
                            </div>
                            
                            {/* Интерактивная легенда */}
                            <div className="flex gap-2.5 text-[9px] font-bold">
                              <span className="flex items-center gap-1 text-sky-400">
                                <span className="w-2.5 h-2.5 rounded-full bg-sky-400"></span> {selectedYear - 1} г.
                              </span>
                              <span className="flex items-center gap-1 text-amber-500">
                                <span className="w-2.5 h-2.5 rounded-full bg-amber-500"></span> {selectedYear} г.
                              </span>
                            </div>
                          </div>

                          {/* Интерактивный тултип данных при наведении */}
                          <div className="h-[210px] mt-2 relative">
                            <svg viewBox={`0 0 ${rpW} ${rpH}`} className="w-full h-full select-none">
                              {/* Grid Horizontal Lines */}
                              {[0, 0.25, 0.5, 0.75, 1].map((ratio, index) => {
                                const val = ratio * reportMaxLimit;
                                const yPos = rpYOf(val);
                                return (
                                  <g key={index}>
                                    <line x1={rpPad} y1={yPos} x2={rpW - rpPad} y2={yPos} stroke={darkTheme ? "#334155" : "#e2e8f0"} strokeDasharray="2 3" />
                                    <text x={rpPad - 6} y={yPos + 3} textAnchor="end" fill={darkTheme ? "#64748b" : "#94a3b8"} className="text-[8px] font-mono">
                                      {Math.round(val).toLocaleString()}
                                    </text>
                                  </g>
                                );
                              })}

                              {/* X Axis Months Label */}
                              {RUSSIAN_MONTHS_SHORT.map((mName, idx) => {
                                const xPos = rpXOf(idx);
                                const isCurrentMonthHighlight = selectedMonthsList.includes(idx + 1);
                                return (
                                  <g key={idx}>
                                    <text 
                                      x={xPos} 
                                      y={rpH - rpPad + 14} 
                                      textAnchor="middle" 
                                      fill={isCurrentMonthHighlight ? (darkTheme ? "#60a5fa" : "#2563eb") : "#64748b"} 
                                      className={`text-[8px] ${isCurrentMonthHighlight ? "font-extrabold" : "font-semibold"}`}
                                    >
                                      {mName}
                                    </text>
                                    {isCurrentMonthHighlight && (
                                      <line x1={xPos} y1={rpPad} x2={xPos} y2={rpH - rpPad} stroke={darkTheme ? "rgba(29,78,216,0.2)" : "rgba(219,234,254,0.6)"} strokeWidth="1.5" strokeDasharray="3 3"/>
                                    )}
                                  </g>
                                );
                              })}

                              {/* Path Previous Year */}
                              <path
                                d={monthlyDataPrev.map((val, idx) => `${idx === 0 ? 'M' : 'L'} ${rpXOf(idx)} ${rpYOf(val)}`).join(' ')}
                                fill="none"
                                stroke="#06b6d4"
                                strokeWidth="2"
                                strokeLinecap="round"
                              />
                              {monthlyDataPrev.map((val, idx) => (
                                <circle
                                  key={`prev-${idx}`}
                                  cx={rpXOf(idx)}
                                  cy={rpYOf(val)}
                                  r={selectedMonthsList.includes(idx + 1) ? "5" : "3.5"}
                                  fill={darkTheme ? "#0f172a" : "#ffffff"}
                                  stroke="#06b6d4"
                                  strokeWidth="2"
                                  className="cursor-pointer transition-transform hover:scale-150"
                                  onMouseEnter={() => setReportHoveredPoint({ month: idx + 1, year: selectedYear - 1, value: val })}
                                  onMouseLeave={() => setReportHoveredPoint(null)}
                                />
                              ))}

                              {/* Path Current Year */}
                              <path
                                d={monthlyDataCurr.map((val, idx) => `${idx === 0 ? 'M' : 'L'} ${rpXOf(idx)} ${rpYOf(val)}`).join(' ')}
                                fill="none"
                                stroke="#f59e0b"
                                strokeWidth="2.5"
                                strokeLinecap="round"
                              />
                              {monthlyDataCurr.map((val, idx) => (
                                <circle
                                  key={`curr-${idx}`}
                                  cx={rpXOf(idx)}
                                  cy={rpYOf(val)}
                                  r={selectedMonthsList.includes(idx + 1) ? "5" : "3.5"}
                                  fill={darkTheme ? "#0f172a" : "#ffffff"}
                                  stroke="#f59e0b"
                                  strokeWidth="2.5"
                                  className="cursor-pointer transition-transform hover:scale-150"
                                  onMouseEnter={() => setReportHoveredPoint({ month: idx + 1, year: selectedYear, value: val })}
                                  onMouseLeave={() => setReportHoveredPoint(null)}
                                />
                              ))}
                            </svg>

                            {/* Hover Tooltip display within the chart zone */}
                            {reportHoveredPoint && (
                              <div className={`absolute bottom-3 right-3 p-2 rounded border text-[10px] shadow-lg backdrop-blur-md z-30 transition-all ${
                                darkTheme ? 'bg-slate-900/95 border-slate-700 text-white' : 'bg-white/95 border-slate-200 text-slate-800'
                              }`}>
                                <span className="font-bold text-blue-400 block border-b border-slate-700/50 pb-0.5 mb-1 text-center">
                                  {RUSSIAN_MONTHS[reportHoveredPoint.month - 1]} {reportHoveredPoint.year} г.
                                </span>
                                <div className="flex justify-between gap-4 font-mono">
                                  <span>Потребление:</span>
                                  <span className="font-bold">{reportHoveredPoint.value.toLocaleString()} кВт·ч</span>
                                </div>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })() : (
                    <div className={`xl:col-span-12 xl:col-span-5 p-6 rounded-xl border flex flex-col items-center justify-center text-center h-[400px] border-dashed ${
                      darkTheme ? 'bg-slate-800/10 border-slate-700 text-slate-400' : 'bg-slate-50 border-slate-200 text-slate-500'
                    }`}>
                      <ChevronRight className="w-10 h-10 text-slate-500 mb-3 animate-pulse" />
                      <h4 className="font-bold text-xs uppercase tracking-wider">Детализация не выбрана</h4>
                      <p className="text-[10px] max-w-[240px] mt-1.5 leading-relaxed">
                        Выберите любую станцию в левой сводной таблице, чтобы посмотреть потребление по её точкам и сравнить графики по месяцам.
                      </p>
                    </div>
                  )}
                </div>
              )}

              {reportSubTab === 'categories' && (
                /* Аналитика по категориям (Светосигналы, отопление, бытовые и т.д.) */
                <div className="space-y-6 animate-fadeIn">
                  {/* Группы категорий с индикатором общей динамики */}
                  <div className="flex flex-wrap gap-2">
                    {categories.map(cat => {
                      const catPoints = supplyPoints.filter(p => p.category === cat && p.isActive);
                      const catPtIds = catPoints.map(p => p.id);
                      const catCurrSum = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                      const catPrevSum = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                      const catDiff = catCurrSum - catPrevSum;
                      const catPct = catPrevSum > 0 ? (catDiff / catPrevSum) * 100 : 0;
                      const isCatSelected = selectedReportCategoryId === cat;

                      // Highlight requested key structures
                      const isTargetHighlight = ["Освещение горловин", "Бытовые нагрузки", "Отопление"].includes(cat);

                      return (
                        <button
                          key={cat}
                          onClick={() => setSelectedReportCategoryId(cat)}
                          className={`px-3 py-2 rounded-xl text-xs font-bold border transition-all duration-150 flex items-center gap-2 ${
                            isCatSelected
                              ? 'bg-blue-600 border-blue-500 text-white shadow-md'
                              : isTargetHighlight
                                ? (darkTheme 
                                    ? 'bg-blue-950/40 border-blue-800/80 text-blue-300 hover:bg-blue-900/30' 
                                    : 'bg-blue-50/50 border-blue-200 text-blue-700 hover:bg-blue-100/40')
                                : (darkTheme
                                    ? 'bg-slate-800/60 border-slate-700/50 text-slate-300 hover:bg-slate-800 hover:text-white'
                                    : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-100')
                          }`}
                        >
                          <span className="flex items-center gap-1">
                            {isTargetHighlight && <span className="w-1.5 h-1.5 rounded-full bg-amber-500 inline-block"></span>}
                            {cat}
                          </span>
                          <span className={`text-[10px] px-1.5 py-0.5 rounded-full font-extrabold ${
                            isCatSelected
                              ? 'bg-blue-800 text-blue-100'
                              : catDiff > 0
                                ? (darkTheme ? 'bg-rose-500/10 text-rose-400' : 'bg-rose-50 text-rose-600')
                                : catDiff < 0
                                  ? (darkTheme ? 'bg-emerald-500/10 text-emerald-400' : 'bg-emerald-50 text-emerald-600')
                                  : 'bg-slate-500/10 text-slate-400'
                          }`}>
                            {catPrevSum > 0 
                              ? (catDiff > 0 ? `+${catPct.toFixed(0)}%` : `${catPct.toFixed(0)}%`) 
                              : (catCurrSum > 0 ? "+100%" : "0%")
                            }
                          </span>
                        </button>
                      );
                    })}
                  </div>

                  {selectedReportCategoryId ? (() => {
                    const selectedCat = selectedReportCategoryId;
                    const catPointsTotal = supplyPoints.filter(p => p.category === selectedCat && p.isActive);
                    const catPtIds = catPointsTotal.map(p => p.id);

                    // Apply the calculation method filter to catPoints
                    const catPoints = catPointsTotal.filter(p => {
                      if (categoryCalculationMethodFilter === 'meter') return p.calculationMethod !== 'estimated';
                      if (categoryCalculationMethodFilter === 'estimated') return p.calculationMethod === 'estimated';
                      return true;
                    });

                    // Calculations
                    const currentSum = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    const previousSum = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    const delta = currentSum - previousSum;
                    const pctDelta = previousSum > 0 ? (delta / previousSum) * 100 : 0;

                    // Group by Stations
                    const stationCalculations = stations.map(st => {
                      const stPointsInCat = catPoints.filter(p => p.stationId === st.id);
                      if (stPointsInCat.length === 0) return null;
                      const ptIds = stPointsInCat.map(p => p.id);
                      const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && ptIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                      const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && ptIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                      const diff = curr - prev;
                      const pct = prev > 0 ? (diff / prev) * 100 : 0;
                      return {
                        station: st,
                        pointsCount: stPointsInCat.length,
                        curr,
                        prev,
                        diff,
                        pct
                      };
                    }).filter(Boolean) as { station: Station; pointsCount: number; curr: number; prev: number; diff: number; pct: number }[];

                    // Detailed points in category
                    const pointCalculations = catPoints.map(p => {
                      const st = stations.find(s => s.id === p.stationId);
                      const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
                      const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
                      const diff = curr - prev;
                      const pct = prev > 0 ? (diff / prev) * 100 : 0;
                      return {
                        point: p,
                        stationName: st ? st.name : 'Неизвестная станция',
                        curr,
                        prev,
                        diff,
                        pct
                      };
                    });

                    const stationSavers = stationCalculations.filter(s => s.diff < 0).length;
                    const stationOverspenders = stationCalculations.filter(s => s.diff > 0).length;
                    const pointSavers = pointCalculations.filter(p => p.diff < 0).length;
                    const pointOverspenders = pointCalculations.filter(p => p.diff > 0).length;

                    return (
                      <div className="space-y-6 animate-fadeIn">
                        {/* Карточки сводной аналитики по категории */}
                        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                          <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/30 border-slate-700/85' : 'bg-white border-slate-200'} shadow-sm flex flex-col justify-between`}>
                            <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Общее потребление</span>
                            <div className="mt-2 flex items-baseline gap-1.5">
                              <span className="text-lg font-extrabold font-mono">{currentSum.toLocaleString()}</span>
                              <span className="text-[9px] text-slate-500 font-bold">кВт·ч</span>
                            </div>
                            <span className="text-[8px] text-slate-500 mt-1">Категория: <span className={darkTheme ? "text-slate-300" : "text-slate-750"}>{selectedCat}</span></span>
                          </div>

                          <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/30 border-slate-700/85' : 'bg-white border-slate-200'} shadow-sm flex flex-col justify-between`}>
                            <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Суммарная динамика г/г</span>
                            <div className="mt-2">
                              {delta > 0 ? (
                                <span className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-lg text-xs font-bold ${
                                  darkTheme ? "bg-rose-500/10 text-rose-400 border border-rose-500/25" : "bg-rose-50 text-rose-700 border border-rose-200"
                                }`}>
                                  <TrendingUp className="w-3.5 h-3.5" />
                                  <span>+{delta.toLocaleString()} (+{pctDelta.toFixed(1)}%)</span>
                                </span>
                              ) : delta < 0 ? (
                                <span className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-lg text-xs font-bold ${
                                  darkTheme ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/25" : "bg-emerald-50 text-emerald-700 border border-emerald-200"
                                }`}>
                                  <TrendingDown className="w-3.5 h-3.5" />
                                  <span>{delta.toLocaleString()} ({pctDelta.toFixed(1)}%)</span>
                                </span>
                              ) : (
                                <span className="bg-slate-500/10 text-slate-400 px-2 py-0.5 rounded-md font-bold text-xs">Без изменений</span>
                              )}
                            </div>
                            <span className="text-[8px] text-slate-500 mt-1">Относительно того же месяца 2025 г.</span>
                          </div>

                          <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/30 border-slate-700/85' : 'bg-white border-slate-200'} shadow-sm flex flex-col justify-between`}>
                            <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Баланс по станциям</span>
                            <div className="mt-2 flex items-center gap-3">
                              <div className="text-center">
                                <span className="text-base font-extrabold text-emerald-500 font-mono flex items-center gap-0.5">
                                  <TrendingDown className="w-3 h-3 text-emerald-500" />
                                  {stationSavers}
                                </span>
                                <span className="block text-[8px] text-slate-400 font-semibold">Экономят</span>
                              </div>
                              <div className="text-center border-l pl-3 border-slate-800">
                                <span className="text-base font-extrabold text-rose-500 font-mono flex items-center gap-0.5">
                                  <TrendingUp className="w-3 h-3 text-rose-500" />
                                  {stationOverspenders}
                                </span>
                                <span className="block text-[8px] text-slate-400 font-semibold">Перерасход</span>
                              </div>
                            </div>
                            <span className="text-[8px] text-slate-500 mt-1">Всего задействовано станций: {stationCalculations.length}</span>
                          </div>

                          <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/30 border-slate-700/85' : 'bg-white border-slate-200'} shadow-sm flex flex-col justify-between`}>
                            <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Баланс по точкам (ТП)</span>
                            <div className="mt-2 flex items-center gap-3">
                              <div className="text-center">
                                <span className="text-base font-extrabold text-emerald-400 font-mono flex items-center gap-0.5">
                                  <ChevronRight className="w-3 h-3 rotate-90 text-emerald-400" />
                                  {pointSavers}
                                </span>
                                <span className="block text-[8px] text-slate-400 font-semibold">Экономят</span>
                              </div>
                              <div className="text-center border-l pl-3 border-slate-800">
                                <span className="text-base font-extrabold text-rose-400 font-mono flex items-center gap-0.5">
                                  <ChevronRight className="w-3 h-3 -rotate-90 text-rose-400" />
                                  {pointOverspenders}
                                </span>
                                <span className="block text-[8px] text-slate-400 font-semibold">Перерасход</span>
                              </div>
                            </div>
                            <span className="text-[8px] text-slate-500 mt-1">Активных ТП в категории: {catPoints.length}</span>
                          </div>
                        </div>

                        {/* Фильтр по способам учета в рамках выбранной категории */}
                        <div className={`flex flex-col sm:flex-row items-center justify-between gap-3 p-4 rounded-xl border ${darkTheme ? 'bg-slate-900/60 border-slate-800' : 'bg-slate-50 border-slate-200'} shadow-sm`}>
                          <div className="flex items-center gap-2">
                            <div className="p-2 bg-blue-500/10 text-blue-400 rounded-lg shrink-0">
                              <Filter className="w-4 h-4 text-blue-500 animate-pulse" />
                            </div>
                            <div>
                              <h5 className="text-xs font-bold uppercase tracking-wider text-slate-300">Способ учета в категории "{selectedCat}"</h5>
                              <p className="text-[9px] text-slate-500">Отфильтруйте точки поставки по наличию прибора учета или расчетному способу</p>
                            </div>
                          </div>
                          <div className={`flex rounded-lg p-0.5 border text-[10px] font-bold ${darkTheme ? 'bg-slate-950 border-slate-850' : 'bg-white border-slate-200'}`}>
                            <button
                              onClick={() => setCategoryCalculationMethodFilter('all')}
                              className={`px-3 py-1.5 rounded-md transition-all duration-150 ${
                                categoryCalculationMethodFilter === 'all'
                                  ? 'bg-blue-600 text-white shadow-sm'
                                  : `${darkTheme ? 'text-slate-450 hover:text-white' : 'text-slate-600 hover:text-slate-950'}`
                              }`}
                            >
                              📈 Все точки ({catPointsTotal.length})
                            </button>
                            <button
                              onClick={() => setCategoryCalculationMethodFilter('meter')}
                              className={`px-3 py-1.5 rounded-md transition-all duration-150 ${
                                categoryCalculationMethodFilter === 'meter'
                                  ? 'bg-blue-600 text-white shadow-sm'
                                  : `${darkTheme ? 'text-slate-450 hover:text-white' : 'text-slate-600 hover:text-slate-950'}`
                              }`}
                            >
                              🎯 Приборы учета ({catPointsTotal.filter(p => p.calculationMethod !== 'estimated').length})
                            </button>
                            <button
                              onClick={() => setCategoryCalculationMethodFilter('estimated')}
                              className={`px-3 py-1.5 rounded-md transition-all duration-150 ${
                                categoryCalculationMethodFilter === 'estimated'
                                  ? 'bg-blue-600 text-white shadow-sm'
                                  : `${darkTheme ? 'text-slate-450 hover:text-white' : 'text-slate-600 hover:text-slate-950'}`
                              }`}
                            >
                              🧮 Расчетный способ ({catPointsTotal.filter(p => p.calculationMethod === 'estimated').length})
                            </button>
                          </div>
                        </div>

                        {/* Зона критики и Экспандеры */}
                        <div className={`flex flex-col sm:flex-row items-center justify-between gap-3 p-4 rounded-xl border ${categoryCriticismOnly ? (darkTheme ? 'bg-rose-950/20 border-rose-900/60' : 'bg-rose-50 border-rose-200') : (darkTheme ? 'bg-slate-900/60 border-slate-800' : 'bg-slate-50 border-slate-200')} shadow-sm`}>
                          <div className="flex items-center gap-3">
                            <div className={`p-2 rounded-lg shrink-0 ${categoryCriticismOnly ? 'bg-rose-500/10 text-rose-500' : 'bg-slate-500/10 text-slate-400'}`}>
                              <AlertTriangle className={`w-4 h-4 ${categoryCriticismOnly ? 'animate-bounce text-rose-500' : ''}`} />
                            </div>
                            <div>
                              <h5 className={`text-xs font-bold uppercase tracking-wider ${categoryCriticismOnly ? 'text-rose-400' : 'text-slate-300'}`}>
                                Фильтрация по критическим превышениям (Зона критики)
                              </h5>
                              <p className="text-[9px] text-slate-500">
                                {categoryCriticismOnly 
                                  ? 'Показаны только станции и точки с ростом расхода (перерасходом) г/г. Нажмите на станцию, чтобы раскрыть её.' 
                                  : 'Отображение всех узлов. Переключите, чтобы сфокусироваться только на узлах с перерасходом.'}
                              </p>
                            </div>
                          </div>
                          
                          <div className="flex items-center gap-3">
                            {/* Кнопка сброса раскрытия */}
                            {(Object.keys(expandedCatStations).length > 0 || Object.keys(expandedCatPoints).length > 0) && (
                              <button
                                onClick={() => {
                                  setExpandedCatStations({});
                                  setExpandedCatPoints({});
                                }}
                                className={`px-2.5 py-1.5 text-[10px] font-bold rounded-lg border transition-all ${
                                  darkTheme 
                                    ? 'bg-slate-800 hover:bg-slate-700 text-slate-300 border-slate-700' 
                                    : 'bg-white hover:bg-slate-100 text-slate-600 border-slate-200'
                                }`}
                              >
                                Свернуть все ТП/станции
                              </button>
                            )}

                            <button
                              onClick={() => {
                                setCategoryCriticismOnly(!categoryCriticismOnly);
                                if (!categoryCriticismOnly) {
                                  // Expand all stations with dynamic growth to make them visible immediately
                                  const overspendingStations: Record<string, boolean> = {};
                                  stationCalculations.forEach(sc => {
                                    if (sc.diff > 0) {
                                      overspendingStations[sc.station.id] = true;
                                    }
                                  });
                                  setExpandedCatStations(overspendingStations);
                                }
                              }}
                              className={`px-4 py-2 text-xs font-extrabold rounded-lg border transition-all duration-150 flex items-center gap-1.5 cursor-pointer shadow-sm ${
                                categoryCriticismOnly
                                  ? 'bg-rose-600 hover:bg-rose-700 border-rose-500 text-white'
                                  : darkTheme
                                    ? 'bg-slate-800 hover:bg-slate-750 border-slate-700 text-slate-300'
                                    : 'bg-white hover:bg-slate-50 border-slate-300 text-slate-700'
                              }`}
                            >
                              <Zap className="w-3.5 h-3.5" />
                              <span>{categoryCriticismOnly ? '🔴 Зона критики: ВКЛ' : '🔘 Зона критики: ВЫКЛ'}</span>
                            </button>
                          </div>
                        </div>

                        {/* Split layouts: Stations comparison left, individual supply points right */}
                        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">

                          {/* Left Panel: Stations performance in this category */}
                          <div className={`lg:col-span-5 p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                            <div className="flex items-center justify-between gap-2 mb-4 border-b border-slate-700/20 pb-3">
                              <div>
                                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-300">Станции железной дороги</h4>
                                <p className="text-[9px] text-slate-500">Суммарный баланс по категории "{selectedCat}"</p>
                              </div>
                              <select
                                value={categoryStationSortOrder}
                                onChange={(e) => setCategoryStationSortOrder(e.target.value as any)}
                                className={`text-[10px] font-bold px-2 py-1 rounded border outline-none ${
                                  darkTheme ? 'bg-slate-800 border-slate-700 text-slate-300' : 'bg-slate-50 border-slate-300 text-slate-700'
                                }`}
                              >
                                <option value="worst-first">⚠️ Худшие → Лучшие (по динамике)</option>
                                <option value="best-first">🍏 Лучшие → Худшие (по динамике)</option>
                                <option value="volume-desc">⚡ Потребление ↓</option>
                                <option value="volume-asc">🔌 Потребление ↑</option>
                              </select>
                            </div>

                            <div className="space-y-4">
                              {(() => {
                                const renderMonthlyDetailTable = (pointId: string) => {
                                  return (
                                    <div className={`mt-2.5 p-3 rounded-lg border text-[11px] ${darkTheme ? 'bg-slate-950/80 border-slate-800' : 'bg-slate-100/80 border-slate-200'} animate-slideDown overflow-x-auto`}>
                                      <div className="font-extrabold uppercase text-[9px] text-slate-500 mb-2 tracking-wider flex items-center justify-between">
                                        <span>Помесячный расход (кВт·ч):</span>
                                        <span className="text-slate-400 font-mono text-[8px] normal-case">Период: {selectedPeriodName}</span>
                                      </div>
                                      <table className="w-full border-collapse text-[10px]">
                                        <thead>
                                          <tr className="border-b border-slate-800/40 text-[9px] font-bold text-slate-500">
                                            <th className="text-left pb-1 font-semibold">Месяц</th>
                                            <th className="text-right pb-1 font-semibold">2025 г.</th>
                                            <th className="text-right pb-1 font-semibold">2026 г.</th>
                                            <th className="text-right pb-1 font-semibold">Разница</th>
                                          </tr>
                                        </thead>
                                        <tbody className="divide-y divide-slate-800/20">
                                          {Array.from({ length: 12 }, (_, i) => i + 1).map(m => {
                                            const isActiveMonth = selectedMonthsList.includes(m);
                                            const mValCurr = readings.find(r => r.year === selectedYear && r.month === m && r.supplyPointId === pointId)?.value || 0;
                                            const mValPrev = readings.find(r => r.year === selectedYear - 1 && r.month === m && r.supplyPointId === pointId)?.value || 0;
                                            const mDiff = mValCurr - mValPrev;
                                            const mSaving = mDiff < 0;

                                            if (mValCurr === 0 && mValPrev === 0) return null;

                                            return (
                                              <tr 
                                                key={m} 
                                                className={`transition-colors ${
                                                  isActiveMonth 
                                                    ? (darkTheme ? 'bg-blue-500/10 text-blue-200 font-semibold' : 'bg-blue-100/50 text-blue-900 font-semibold') 
                                                    : 'text-slate-400'
                                                }`}
                                              >
                                                <td className="py-1 flex items-center gap-1">
                                                  {isActiveMonth && <span className="inline-block w-1 h-1 rounded-full bg-blue-500"></span>}
                                                  <span>{RUSSIAN_MONTHS[m - 1]}</span>
                                                </td>
                                                <td className="py-1 text-right font-mono">{mValPrev.toLocaleString()}</td>
                                                <td className={`py-1 text-right font-mono ${isActiveMonth ? (mSaving ? 'text-emerald-400' : mDiff > 0 ? 'text-rose-450' : '') : ''}`}>
                                                  {mValCurr.toLocaleString()}
                                                </td>
                                                <td className={`py-1 text-right font-mono font-bold ${
                                                  mDiff > 0 ? 'text-rose-500' : mDiff < 0 ? 'text-emerald-500' : 'text-slate-500'
                                                }`}>
                                                  {mDiff > 0 ? `+${mDiff.toLocaleString()}` : mDiff.toLocaleString()}
                                                </td>
                                              </tr>
                                            );
                                          })}
                                        </tbody>
                                      </table>
                                    </div>
                                  );
                                };

                                const filteredStationsCalculations = categoryCriticismOnly
                                  ? stationCalculations.filter(sc => sc.diff > 0)
                                  : stationCalculations;

                                if (filteredStationsCalculations.length === 0) {
                                  return (
                                    <div className="text-center py-8 text-slate-500 text-xs">Нет станций с повышенным расходом в выбранном режиме</div>
                                  );
                                }

                                return [...filteredStationsCalculations].sort((a, b) => {
                                  if (categoryStationSortOrder === 'worst-first') return b.diff - a.diff;
                                  if (categoryStationSortOrder === 'best-first') return a.diff - b.diff;
                                  if (categoryStationSortOrder === 'volume-desc') return b.curr - a.curr;
                                  if (categoryStationSortOrder === 'volume-asc') return a.curr - b.curr;
                                  return 0;
                                }).map(sc => {
                                  const isSaving = sc.diff < 0;
                                  const isExpanded = !!expandedCatStations[sc.station.id];
                                  
                                  const stationPoints = catPoints.filter(p => p.stationId === sc.station.id);
                                  const filteredStationPoints = categoryCriticismOnly
                                    ? stationPoints.filter(p => {
                                        const currPt = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
                                        const prevPt = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
                                        return (currPt - prevPt) > 0;
                                      })
                                    : stationPoints;

                                  return (
                                    <div 
                                      key={sc.station.id} 
                                      className={`p-3 rounded-xl border transition-all duration-150 ${
                                        isExpanded 
                                          ? (darkTheme ? 'bg-slate-800/60 border-slate-650 shadow-md' : 'bg-slate-50 border-slate-300 shadow-md')
                                          : (darkTheme ? 'bg-slate-800/20 border-slate-800 hover:bg-slate-800/35 hover:border-slate-700/60' : 'bg-white border-slate-100 hover:bg-slate-55 hover:border-slate-200')
                                      }`}
                                    >
                                      <div 
                                        className="flex items-center justify-between cursor-pointer"
                                        onClick={() => setExpandedCatStations(prev => ({ ...prev, [sc.station.id]: !prev[sc.station.id] }))}
                                      >
                                        <div className="flex items-center gap-2">
                                          <span className={`transition-transform duration-150 ${isExpanded ? 'rotate-90 text-blue-500' : 'text-slate-500'}`}>
                                            <ChevronRight className="w-3.5 h-3.5" />
                                          </span>
                                          <span className="font-bold text-xs">{sc.station.name}</span>
                                          <button
                                            title="Открыть карточку станции"
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              setSelectedStationId(sc.station.id);
                                              setSelectedSupplyPointId(null);
                                              setActiveTab('stations');
                                            }}
                                            className="p-1 rounded text-blue-400 hover:bg-blue-500/10 transition-colors"
                                          >
                                            <ExternalLink className="w-3.5 h-3.5" />
                                          </button>
                                        </div>
                                        <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[10px] font-bold ${
                                          isSaving 
                                            ? (darkTheme ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' : 'bg-emerald-50 text-emerald-600 border border-emerald-100') 
                                            : (darkTheme ? 'bg-rose-500/10 text-rose-400 border border-rose-500/20' : 'bg-rose-50 text-rose-600 border border-rose-100')
                                        }`}>
                                          {isSaving ? <TrendingDown className="w-3 h-3 text-emerald-400" /> : <TrendingUp className="w-3 h-3 text-rose-400" />}
                                          <span>
                                            {sc.diff > 0 ? `+${sc.diff.toLocaleString()}` : sc.diff.toLocaleString()} кВт·ч ({sc.pct > 0 ? `+${sc.pct.toFixed(1)}%` : `${sc.pct.toFixed(1)}%`})
                                          </span>
                                        </span>
                                      </div>

                                      <div className="flex items-center justify-between text-[9px] text-slate-500 font-mono mt-1.5 pl-5">
                                        <span>ТП в категории: <span className="font-bold text-slate-400">{filteredStationPoints.length} / {sc.pointsCount}</span></span>
                                        <div className="flex gap-2">
                                          <span>25г: <span className="font-bold">{sc.prev.toLocaleString()}</span></span>
                                          <span>26г: <span className={`font-bold ${isSaving ? "text-emerald-400" : "text-rose-400"}`}>{sc.curr.toLocaleString()}</span></span>
                                        </div>
                                      </div>

                                      {isExpanded && (
                                        <div className="mt-3 pt-3 border-t border-slate-700/30 space-y-2 animate-fadeIn pl-1">
                                          <div className="text-[9px] font-extrabold uppercase tracking-widest text-slate-450 mb-1.5">
                                            {categoryCriticismOnly ? '⚠️ ТП в зоне критики (перерасход):' : '🔌 Подключенные точки поставки:'}
                                          </div>
                                          {filteredStationPoints.length === 0 ? (
                                            <div className="text-center py-2 text-slate-500 text-[10px]">Нет точек в выбранном режиме</div>
                                          ) : (
                                            filteredStationPoints.map(p => {
                                              const currPt = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
                                              const prevPt = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
                                              const diffPt = currPt - prevPt;
                                              const pctPt = prevPt > 0 ? (diffPt / prevPt) * 100 : 0;
                                              const isPtExpanded = !!expandedCatPoints[p.id];

                                              return (
                                                <div 
                                                  key={p.id} 
                                                  className={`p-2 rounded-lg border text-[11px] transition-all ${
                                                    isPtExpanded
                                                      ? (darkTheme ? 'bg-slate-900 border-amber-500/45 shadow' : 'bg-slate-100 border-amber-600/30 shadow')
                                                      : (darkTheme ? 'bg-slate-900/40 border-slate-800 hover:bg-slate-900/70' : 'bg-slate-50 border-slate-200 hover:bg-slate-100/80')
                                                  }`}
                                                >
                                                  <div 
                                                    className="flex items-center justify-between cursor-pointer"
                                                    onClick={(e) => {
                                                      e.stopPropagation();
                                                      setExpandedCatPoints(prev => ({ ...prev, [p.id]: !prev[p.id] }));
                                                    }}
                                                  >
                                                    <div className="flex items-center gap-1.5 truncate max-w-[150px]">
                                                      <span className={`transition-transform duration-150 ${isPtExpanded ? 'rotate-90 text-amber-500' : 'text-slate-500'}`}>
                                                        <ChevronRight className="w-3 h-3" />
                                                      </span>
                                                      <span className="font-bold truncate" title={p.name}>{p.name}</span>
                                                      <button
                                                        title="Открыть карточку ТП"
                                                        onClick={(e) => {
                                                          e.stopPropagation();
                                                          setSelectedStationId(p.stationId);
                                                          setSelectedSupplyPointId(p.id);
                                                          setActiveTab('stations');
                                                        }}
                                                        className="p-1 rounded text-amber-400 hover:bg-amber-500/10 transition-colors shrink-0"
                                                      >
                                                        <ExternalLink className="w-3 h-3" />
                                                      </button>
                                                      <button
                                                        title="Редактировать показатели ТП"
                                                        onClick={(e) => {
                                                          e.stopPropagation();
                                                          const currK = readings.find(r => r.supplyPointId === p.id && r.year === selectedYear && r.month === selectedMonth)?.value || 0;
                                                          const prevK = readings.find(r => r.supplyPointId === p.id && r.year === selectedYear - 1 && r.month === selectedMonth)?.value || 0;
                                                          setPointModal({
                                                            isOpen: true,
                                                            mode: 'edit',
                                                            id: p.id,
                                                            name: p.name,
                                                            stationId: p.stationId,
                                                            category: p.category,
                                                            note: p.note,
                                                            isActive: p.isActive,
                                                            calculationMethod: p.calculationMethod || 'meter',
                                                            currKwh: currK,
                                                            prevKwh: prevK
                                                          });
                                                        }}
                                                        className="p-1 rounded text-sky-400 hover:bg-sky-500/10 transition-colors shrink-0"
                                                      >
                                                        <Edit2 className="w-3 h-3" />
                                                      </button>
                                                    </div>
                                                    <span className={`font-mono font-bold text-[10px] ${
                                                      diffPt > 0 ? 'text-rose-450' : diffPt < 0 ? 'text-emerald-400' : 'text-slate-400'
                                                    }`}>
                                                      {diffPt > 0 ? `+${diffPt.toLocaleString()}` : diffPt.toLocaleString()} кВт·ч ({pctPt > 0 ? `+${pctPt.toFixed(0)}%` : `${pctPt.toFixed(0)}%`})
                                                    </span>
                                                  </div>

                                                  <div className="flex items-center justify-between text-[8px] text-slate-500 mt-1 pl-4">
                                                    <span className="uppercase font-semibold text-slate-400">{p.calculationMethod === 'estimated' ? 'Расчетный' : 'Прибор учета'}</span>
                                                    <span>25г: {prevPt.toLocaleString()} | 26г: {currPt.toLocaleString()}</span>
                                                  </div>

                                                  {isPtExpanded && renderMonthlyDetailTable(p.id)}
                                                </div>
                                              );
                                            })
                                          )}
                                        </div>
                                      )}
                                    </div>
                                  );
                                });
                              })()}

                              {false && [...stationCalculations].sort((a, b) => {
                                if (categoryStationSortOrder === 'worst-first') return b.diff - a.diff;
                                if (categoryStationSortOrder === 'best-first') return a.diff - b.diff;
                                if (categoryStationSortOrder === 'volume-desc') return b.curr - a.curr;
                                if (categoryStationSortOrder === 'volume-asc') return a.curr - b.curr;
                                return 0;
                              }).map(sc => {
                                const isSaving = sc.diff < 0;
                                const maxVal = Math.max(...stationCalculations.map(s => Math.max(s.curr, s.prev, 1)));
                                const currPctWeb = (sc.curr / maxVal) * 100;
                                const prevPctWeb = (sc.prev / maxVal) * 100;

                                return (
                                  <div key={sc.station.id} className="space-y-1.5 pb-3.5 border-b last:border-b-0 border-slate-700/30">
                                     <div className="flex items-center justify-between">
                                       <span className="font-bold text-xs">{sc.station.name}</span>
                                       <span className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[10px] font-bold ${
                                         isSaving 
                                           ? (darkTheme ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' : 'bg-emerald-50 text-emerald-600 border border-emerald-100') 
                                           : (darkTheme ? 'bg-rose-500/10 text-rose-400 border border-rose-500/20' : 'bg-rose-50 text-rose-600 border border-rose-100')
                                       }`}>
                                         {isSaving ? <TrendingDown className="w-3 h-3 text-emerald-400" /> : <TrendingUp className="w-3 h-3 text-rose-400" />}
                                         <span>
                                           {sc.diff > 0 ? `+${sc.diff.toLocaleString()}` : sc.diff.toLocaleString()} кВт·ч ({sc.pct > 0 ? `+${sc.pct.toFixed(1)}%` : `${sc.pct.toFixed(1)}%`})
                                         </span>
                                       </span>
                                     </div>

                                     <div className="flex items-center justify-between text-[9px] text-slate-500 font-mono">
                                       <span>Подключено ТП: <span className="font-bold text-slate-400">{sc.pointsCount}</span></span>
                                       <div className="flex gap-2">
                                         <span>2025 г: <span className="font-bold">{sc.prev.toLocaleString()}</span></span>
                                         <span>2026 г: <span className={`font-bold ${isSaving ? "text-emerald-400" : "text-rose-400"}`}>{sc.curr.toLocaleString()}</span></span>
                                       </div>
                                     </div>
                                   </div>
                                 );
                               })}
                             </div>
                           </div>

                           {/* Right Panel: Supply points list within this category */}
                           <div className={`lg:col-span-7 p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                             <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-4 border-b border-slate-700/20 pb-3">
                               <div>
                                 <h4 className="text-xs font-bold uppercase tracking-wider text-slate-300">Точки поставки в категории</h4>
                                 <p className="text-[9px] text-slate-500">Детальное сравнение по точкам потребления</p>
                               </div>
                               <select
                                 value={categoryPointSortOrder}
                                 onChange={(e) => setCategoryPointSortOrder(e.target.value as any)}
                                 className={`text-[10px] font-bold px-2 py-1 rounded border outline-none ${
                                   darkTheme ? 'bg-slate-800 border-slate-700 text-slate-300' : 'bg-slate-50 border-slate-300 text-slate-700'
                                 }`}
                               >
                                 <option value="worst-first">⚠️ Худшие → Лучшие (по динамике)</option>
                                 <option value="best-first">🍏 Лучшие → Худшие (по динамике)</option>
                                 <option value="volume-desc">⚡ Потребление ↓</option>
                                 <option value="volume-asc">🔌 Потребление ↑</option>
                               </select>
                             </div>

                             <div className="overflow-y-auto max-h-[400px] space-y-3 pr-1">
                               {(() => {
                                 const filteredPoints = categoryCriticismOnly
                                   ? pointCalculations.filter(pc => pc.diff > 0)
                                   : pointCalculations;

                                 if (filteredPoints.length === 0) {
                                   return (
                                     <div className="text-center py-8 text-slate-500 text-xs">Нет точек в выбранном режиме</div>
                                   );
                                 }

                                 const renderMonthlyDetailTableRight = (pointId: string) => {
                                   return (
                                     <div className={`mt-2 p-3 rounded-lg border text-[10px] ${darkTheme ? 'bg-slate-950/80 border-slate-800' : 'bg-slate-100/80 border-slate-200'} overflow-x-auto`}>
                                       <div className="font-extrabold uppercase text-[8px] text-slate-500 mb-2 tracking-wider flex items-center justify-between">
                                         <span>Помесячный расход (кВт·ч):</span>
                                       </div>
                                       <table className="w-full border-collapse">
                                         <thead>
                                           <tr className="border-b border-slate-800/40 text-[8px] font-bold text-slate-500">
                                             <th className="text-left pb-1">Месяц</th>
                                             <th className="text-right pb-1">2025 г.</th>
                                             <th className="text-right pb-1">2026 г.</th>
                                             <th className="text-right pb-1">Разница</th>
                                           </tr>
                                         </thead>
                                         <tbody className="divide-y divide-slate-800/20">
                                           {Array.from({ length: 12 }, (_, i) => i + 1).map(m => {
                                             const isActiveMonth = selectedMonthsList.includes(m);
                                             const mValCurr = readings.find(r => r.year === selectedYear && r.month === m && r.supplyPointId === pointId)?.value || 0;
                                             const mValPrev = readings.find(r => r.year === selectedYear - 1 && r.month === m && r.supplyPointId === pointId)?.value || 0;
                                             const mDiff = mValCurr - mValPrev;
                                             const mSaving = mDiff < 0;

                                             if (mValCurr === 0 && mValPrev === 0) return null;

                                             return (
                                               <tr key={m} className={isActiveMonth ? (darkTheme ? 'bg-blue-500/10 font-semibold text-blue-300' : 'bg-blue-100/50 font-semibold text-blue-900') : 'text-slate-400'}>
                                                 <td className="py-1">{RUSSIAN_MONTHS[m - 1]}</td>
                                                 <td className="py-1 text-right font-mono">{mValPrev.toLocaleString()}</td>
                                                 <td className={`py-1 text-right font-mono ${isActiveMonth ? (mSaving ? 'text-emerald-400' : mDiff > 0 ? 'text-rose-450' : '') : ''}`}>
                                                   {mValCurr.toLocaleString()}
                                                 </td>
                                                 <td className={`py-1 text-right font-mono font-bold ${
                                                   mDiff > 0 ? 'text-rose-500' : mDiff < 0 ? 'text-emerald-500' : 'text-slate-500'
                                                 }`}>
                                                   {mDiff > 0 ? `+${mDiff.toLocaleString()}` : mDiff.toLocaleString()}
                                                 </td>
                                               </tr>
                                             );
                                           })}
                                         </tbody>
                                       </table>
                                     </div>
                                   );
                                 };

                                 return [...filteredPoints].sort((a, b) => {
                                   if (categoryPointSortOrder === 'worst-first') return b.diff - a.diff;
                                   if (categoryPointSortOrder === 'best-first') return a.diff - b.diff;
                                   if (categoryPointSortOrder === 'volume-desc') return b.curr - a.curr;
                                   if (categoryPointSortOrder === 'volume-asc') return a.curr - b.curr;
                                   return 0;
                                 }).map(pc => {
                                   const isSaving = pc.diff < 0;
                                   const isPtExpanded = !!expandedCatPoints[pc.point.id];

                                   return (
                                      <div 
                                        key={pc.point.id} 
                                        className={`p-3 rounded-lg border flex flex-col gap-2 transition-all cursor-pointer ${
                                          isPtExpanded
                                            ? (darkTheme ? 'bg-slate-900 border-amber-500/40 shadow-md' : 'bg-slate-100/80 border-amber-500/30 shadow-md')
                                            : (darkTheme ? 'bg-slate-900/40 border-slate-800/60 hover:bg-slate-900/70' : 'bg-slate-50 border-slate-200 hover:bg-slate-100/60')
                                        }`}
                                        onClick={() => setExpandedCatPoints(prev => ({ ...prev, [pc.point.id]: !prev[pc.point.id] }))}
                                      >
                                        <div className="flex items-start justify-between gap-3">
                                          <div>
                                            <div className="text-xs font-bold flex items-center gap-1.5">
                                              <span className={`transition-transform duration-150 ${isPtExpanded ? 'rotate-90 text-amber-500' : 'text-slate-500'}`}>
                                                <ChevronRight className="w-3.5 h-3.5" />
                                              </span>
                                              <span className={`w-1.5 h-1.5 rounded-full ${pc.point.isActive ? 'bg-emerald-500' : 'bg-slate-500'}`}></span>
                                              <span>{pc.point.name}</span>
                                              <button
                                                title="Открыть карточку ТП"
                                                onClick={(e) => {
                                                  e.stopPropagation();
                                                  setSelectedStationId(pc.point.stationId);
                                                  setSelectedSupplyPointId(pc.point.id);
                                                  setActiveTab('stations');
                                                }}
                                                className="p-1 rounded text-amber-400 hover:bg-amber-500/10 transition-colors shrink-0"
                                              >
                                                <ExternalLink className="w-3.5 h-3.5" />
                                              </button>
                                              <button
                                                title="Редактировать показатели ТП"
                                                onClick={(e) => {
                                                  e.stopPropagation();
                                                  const currK = readings.find(r => r.supplyPointId === pc.point.id && r.year === selectedYear && r.month === selectedMonth)?.value || 0;
                                                  const prevK = readings.find(r => r.supplyPointId === pc.point.id && r.year === selectedYear - 1 && r.month === selectedMonth)?.value || 0;
                                                  setPointModal({
                                                    isOpen: true,
                                                    mode: 'edit',
                                                    id: pc.point.id,
                                                    name: pc.point.name,
                                                    stationId: pc.point.stationId,
                                                    category: pc.point.category,
                                                    note: pc.point.note,
                                                    isActive: pc.point.isActive,
                                                    calculationMethod: pc.point.calculationMethod || 'meter',
                                                    currKwh: currK,
                                                    prevKwh: prevK
                                                  });
                                                }}
                                                className="p-1 rounded text-sky-400 hover:bg-sky-500/10 transition-colors shrink-0"
                                              >
                                                <Edit2 className="w-3.5 h-3.5" />
                                              </button>
                                            </div>
                                            <div className="text-[9px] text-slate-500 font-semibold uppercase mt-0.5 pl-4 flex items-center gap-1">
                                              <span>{pc.stationName}</span>
                                              <button
                                                title="Открыть карточку станции"
                                                onClick={(e) => {
                                                  e.stopPropagation();
                                                  setSelectedStationId(pc.point.stationId);
                                                  setSelectedSupplyPointId(null);
                                                  setActiveTab('stations');
                                                }}
                                                className="p-0.5 rounded text-blue-400 hover:bg-blue-500/10 transition-colors"
                                              >
                                                <ExternalLink className="w-2.5 h-2.5" />
                                              </button>
                                            </div>
                                          </div>
                                          <span className={`text-[8px] font-extrabold uppercase px-1.5 py-0.5 rounded ${
                                            pc.point.calculationMethod === 'estimated' 
                                              ? 'bg-amber-500/10 text-amber-500' 
                                              : 'bg-blue-500/10 text-blue-400'
                                          }`}>
                                            {pc.point.calculationMethod === 'estimated' ? 'Расчетный' : 'Прибор учета'}
                                          </span>
                                        </div>

                                        <div className="grid grid-cols-3 gap-2 border-t border-slate-700/20 pt-2 text-[10px] font-mono pl-4">
                                          <div>
                                            <span className="block text-[8px] text-slate-500 font-semibold uppercase">2025 г.</span>
                                            <span className="font-bold">{pc.prev.toLocaleString()} кВт·ч</span>
                                          </div>
                                          <div>
                                            <span className="block text-[8px] text-slate-500 font-semibold uppercase">2026 г.</span>
                                            <span className={`font-bold ${isSaving ? 'text-emerald-400' : 'text-rose-400'}`}>{pc.curr.toLocaleString()} кВт·ч</span>
                                          </div>
                                          <div>
                                            <span className="block text-[8px] text-slate-500 font-semibold uppercase">Динамика</span>
                                            <span className={`font-bold flex items-center ${isSaving ? 'text-emerald-400' : 'text-rose-400'}`}>
                                              {isSaving ? <TrendingDown className="w-3 h-3 mr-0.5" /> : <TrendingUp className="w-3 h-3 mr-0.5" />}
                                              {pc.diff > 0 ? `+${pc.diff.toLocaleString()}` : pc.diff.toLocaleString()} ({pc.pct > 0 ? `+${pc.pct.toFixed(0)}%` : `${pc.pct.toFixed(0)}%`})
                                            </span>
                                          </div>
                                        </div>

                                        {isPtExpanded && (
                                          <div onClick={(e) => e.stopPropagation()}>
                                            {renderMonthlyDetailTableRight(pc.point.id)}
                                          </div>
                                        )}
                                      </div>
                                    );
                                  });
                               })()}

                               {false && [...pointCalculations].sort((a, b) => {
                                 if (categoryPointSortOrder === 'worst-first') return b.diff - a.diff;
                                 if (categoryPointSortOrder === 'best-first') return a.diff - b.diff;
                                 if (categoryPointSortOrder === 'volume-desc') return b.curr - a.curr;
                                 if (categoryPointSortOrder === 'volume-asc') return a.curr - b.curr;
                                 return 0;
                               }).map(pc => {
                                 const isSaving = pc.diff < 0;
                                   return (
                                     <div key={pc.point.id} className={`p-3 rounded-lg border flex flex-col gap-2 ${darkTheme ? 'bg-slate-900/40 border-slate-800/60' : 'bg-slate-50 border-slate-200'}`}>
                                       <div className="flex items-start justify-between gap-3">
                                         <div>
                                           <div className="text-xs font-bold flex items-center gap-1.5">
                                             <span className={`w-1.5 h-1.5 rounded-full ${pc.point.isActive ? 'bg-emerald-500' : 'bg-slate-500'}`}></span>
                                             {pc.point.name}
                                           </div>
                                           <div className="text-[9px] text-slate-500 font-semibold uppercase mt-0.5">
                                             {pc.stationName}
                                           </div>
                                         </div>
                                         <span className={`text-[8px] font-extrabold uppercase px-1.5 py-0.5 rounded ${
                                           pc.point.calculationMethod === 'estimated' 
                                             ? 'bg-amber-500/10 text-amber-500' 
                                             : 'bg-blue-500/10 text-blue-400'
                                         }`}>
                                           {pc.point.calculationMethod === 'estimated' ? 'Расчетный' : 'Прибор учета'}
                                         </span>
                                       </div>

                                       <div className="grid grid-cols-3 gap-2 border-t border-slate-700/20 pt-2 text-[10px] font-mono">
                                         <div>
                                           <span className="block text-[8px] text-slate-500 font-semibold uppercase">2025 г.</span>
                                           <span className="font-bold">{pc.prev.toLocaleString()} кВт·ч</span>
                                         </div>
                                         <div>
                                           <span className="block text-[8px] text-slate-500 font-semibold uppercase">2026 г.</span>
                                           <span className={`font-bold ${isSaving ? 'text-emerald-400' : 'text-rose-400'}`}>{pc.curr.toLocaleString()} кВт·ч</span>
                                         </div>
                                         <div>
                                           <span className="block text-[8px] text-slate-500 font-semibold uppercase">Динамика</span>
                                           <span className={`font-bold flex items-center ${isSaving ? 'text-emerald-400' : 'text-rose-400'}`}>
                                             {isSaving ? <TrendingDown className="w-3 h-3 mr-0.5" /> : <TrendingUp className="w-3 h-3 mr-0.5" />}
                                             {pc.diff > 0 ? `+${pc.diff.toLocaleString()}` : pc.diff.toLocaleString()} ({pc.pct > 0 ? `+${pc.pct.toFixed(0)}%` : `${pc.pct.toFixed(0)}%`})
                                           </span>
                                         </div>
                                       </div>
                                     </div>
                                   );
                                 })
                               }
                             </div>
                           </div>
                         </div>
                       </div>
                     );
                   })() : (
                     <div className="text-center py-8 text-slate-500 text-xs">Выберите категорию из предложенных выше</div>
                   )}
                 </div>
               )}

              {reportSubTab === 'accounting_methods' && (
                <div className="space-y-6 animate-fadeIn">
                  {/* Заголовок и вводная часть */}
                  <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                    <div className="flex items-start gap-4">
                      <div className="p-3 bg-amber-500/10 text-amber-500 rounded-lg shrink-0">
                        <Activity className="w-6 h-6" />
                      </div>
                      <div className="space-y-1">
                        <h4 className="text-sm font-bold">Сравнительный анализ приборного учета и расчетного способа определения расхода</h4>
                        <p className={`text-xs leading-relaxed ${darkTheme ? 'text-slate-400' : 'text-slate-600'}`}>
                          Системный аудит безприборных подключений (расчетный способ по установленной мощности) в сравнении с фактическими показаниями коммерческих приборов учета. 
                          Основной целью энергетического аудита ОАО РЖД является перевод расчетных точек на цифровые приборы учета для минимизации сверхнормативных потерь.
                        </p>
                      </div>
                    </div>
                  </div>

                  {/* Карточки KPI для способов учета */}
                  {(() => {
                    // Calculate overall global totals for calculation methods
                    const meterPoints = supplyPoints.filter(p => p.isActive && p.calculationMethod !== 'estimated');
                    const estimatedPoints = supplyPoints.filter(p => p.isActive && p.calculationMethod === 'estimated');

                    const meterCurr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && meterPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    const meterPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && meterPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    const meterDiff = meterCurr - meterPrev;
                    const meterPct = meterPrev > 0 ? (meterDiff / meterPrev) * 100 : 0;

                    const estCurr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && estimatedPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    const estPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && estimatedPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    const estDiff = estCurr - estPrev;
                    const estPct = estPrev > 0 ? (estDiff / estPrev) * 100 : 0;

                    const totalKwh = meterCurr + estCurr;
                    const meterRatio = totalKwh > 0 ? (meterCurr / totalKwh) * 100 : 0;
                    const estRatio = totalKwh > 0 ? (estCurr / totalKwh) * 100 : 0;

                    return (
                      <>
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                          {/* KPI 1: Приборы учета */}
                          <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm space-y-3`}>
                            <div className="flex items-center justify-between">
                              <span className="text-[10px] font-bold uppercase tracking-wider text-blue-400">По приборам учета (ПУ)</span>
                              <span className="text-[9px] font-extrabold bg-blue-500/10 text-blue-400 px-2 py-0.5 rounded-full">{meterPoints.length} ТП</span>
                            </div>
                            <div className="space-y-1">
                              <div className="text-lg font-mono font-extrabold">{meterCurr.toLocaleString()} кВт·ч</div>
                              <p className={`text-[10px] ${darkTheme ? 'text-slate-400' : 'text-slate-550'}`}>Доля в общем расходе: <strong className="text-blue-400">{meterRatio.toFixed(1)}%</strong></p>
                            </div>
                            <div className="flex items-center gap-1.5 pt-1.5 border-t border-slate-800/40 text-[10px]">
                              {meterDiff > 0 ? (
                                <span className="flex items-center text-rose-500 font-bold"><TrendingUp className="w-3.5 h-3.5 mr-0.5" /> +{meterPct.toFixed(2)}%</span>
                              ) : meterDiff < 0 ? (
                                <span className="flex items-center text-emerald-500 font-bold"><TrendingDown className="w-3.5 h-3.5 mr-0.5" /> {meterPct.toFixed(2)}%</span>
                              ) : (
                                <span className="text-slate-400">Без изменений</span>
                              )}
                              <span className="text-slate-550">к г/г</span>
                            </div>
                          </div>

                          {/* KPI 2: Расчетный способ */}
                          <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm space-y-3`}>
                            <div className="flex items-center justify-between">
                              <span className="text-[10px] font-bold uppercase tracking-wider text-amber-500">Расчетный способ (уст. мощность)</span>
                              <span className="text-[9px] font-extrabold bg-amber-500/10 text-amber-400 px-2 py-0.5 rounded-full">{estimatedPoints.length} ТП</span>
                            </div>
                            <div className="space-y-1">
                              <div className="text-lg font-mono font-extrabold">{estCurr.toLocaleString()} кВт·ч</div>
                              <p className={`text-[10px] ${darkTheme ? 'text-slate-400' : 'text-slate-550'}`}>Доля в общем расходе: <strong className="text-amber-500">{estRatio.toFixed(1)}%</strong></p>
                            </div>
                            <div className="flex items-center gap-1.5 pt-1.5 border-t border-slate-800/40 text-[10px]">
                              {estDiff > 0 ? (
                                <span className="flex items-center text-rose-500 font-bold"><TrendingUp className="w-3.5 h-3.5 mr-0.5" /> +{estPct.toFixed(2)}%</span>
                              ) : estDiff < 0 ? (
                                <span className="flex items-center text-emerald-500 font-bold"><TrendingDown className="w-3.5 h-3.5 mr-0.5" /> {estPct.toFixed(2)}%</span>
                              ) : (
                                <span className="text-slate-400">Без изменений</span>
                              )}
                              <span className="text-slate-550">к г/г</span>
                            </div>
                          </div>

                          {/* KPI 3: Качество учета на участке */}
                          <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm flex flex-col justify-between`}>
                            <div className="space-y-1">
                              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Индекс эффективности приборного учета</span>
                              <h5 className={`text-xs font-bold mt-1 ${darkTheme ? 'text-slate-200' : 'text-slate-800'}`}>
                                {meterRatio >= 90 ? '🟢 Отличный уровень' : meterRatio >= 75 ? '🟡 Приемлемый уровень' : '🔴 Требуется модернизация'}
                              </h5>
                            </div>
                            <div className="space-y-1.5 mt-2">
                              <div className="flex items-center justify-between text-[10px] font-mono">
                                <span>Доля ПУ: {meterRatio.toFixed(1)}%</span>
                                <span>Цель: 100%</span>
                              </div>
                              <div className={`w-full h-2 rounded-full overflow-hidden ${darkTheme ? 'bg-slate-800' : 'bg-slate-200'}`}>
                                <div className="bg-emerald-500 h-full rounded-full" style={{ width: `${meterRatio}%` }} />
                              </div>
                            </div>
                            <p className="text-[9px] text-slate-400 leading-normal mt-2">
                              {estRatio > 20 
                                ? '⚠️ Высокий процент расчетного способа. Рекомендуется установить умные счетчики ПУ на ст. Александров-1.' 
                                : 'Высокий уровень прозрачности учета. Ситуация находится под контролем.'}
                            </p>
                          </div>
                        </div>

                        {/* Разделенные таблицы */}
                        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                          {/* Слева: Сводная ведомость станций по способам учета */}
                          <div className={`lg:col-span-7 p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                            <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-4">Сравнение способов учета по станциям железной дороги</h4>
                            
                            <div className="overflow-x-auto">
                              <table className="w-full text-left text-xs">
                                <thead className={`font-bold ${darkTheme ? 'bg-slate-900/60 text-slate-400' : 'bg-slate-100 text-slate-700'}`}>
                                  <tr>
                                    <th className="p-2">Станция</th>
                                    <th className="p-2 text-right">По приборам (кВт·ч)</th>
                                    <th className="p-2 text-right">По расчетным (кВт·ч)</th>
                                    <th className="p-2 text-right">Доля расчетных (%)</th>
                                    <th className="p-2 text-right">Статус</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-800/40">
                                  {stations.map(st => {
                                    const stPoints = supplyPoints.filter(p => p.stationId === st.id && p.isActive);
                                    const stMeterPoints = stPoints.filter(p => p.calculationMethod !== 'estimated');
                                    const stEstPoints = stPoints.filter(p => p.calculationMethod === 'estimated');

                                    const stMeterVal = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && stMeterPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                                    const stEstVal = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && stEstPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                                    const stTotal = stMeterVal + stEstVal;
                                    const stEstPct = stTotal > 0 ? (stEstVal / stTotal) * 100 : 0;

                                    return (
                                      <tr key={st.id} className={darkTheme ? "hover:bg-slate-800/25" : "hover:bg-slate-50"}>
                                        <td className="p-2 font-bold">{st.name}</td>
                                        <td className="p-2 text-right font-mono text-blue-400">{stMeterVal.toLocaleString()}</td>
                                        <td className="p-2 text-right font-mono text-amber-500">{stEstVal.toLocaleString()}</td>
                                        <td className="p-2 text-right font-mono font-semibold">
                                          {stEstPct > 0 ? `${stEstPct.toFixed(1)}%` : '0%'}
                                        </td>
                                        <td className="p-2 text-right">
                                          <span className={`inline-block px-2 py-0.5 text-[8px] font-extrabold uppercase rounded ${
                                            stEstPct === 0 
                                              ? 'bg-emerald-500/10 text-emerald-400' 
                                              : stEstPct > 35 
                                                ? 'bg-rose-500/10 text-rose-400' 
                                                : 'bg-amber-500/10 text-amber-400'
                                          }`}>
                                            {stEstPct === 0 ? '100% ПУ' : stEstPct > 35 ? 'Критично' : 'Частично'}
                                          </span>
                                        </td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            </div>
                          </div>

                          {/* Справа: Список расчетных ТП (без приборов учета) */}
                          <div className={`lg:col-span-5 p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm space-y-4`}>
                            <div>
                              <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400">Точки поставки на расчетном способе</h4>
                              <p className="text-[10px] text-slate-500">Данные точки поставки не имеют приборов учета и подлежат дооснащению ПУ</p>
                            </div>

                            <div className="overflow-y-auto max-h-[350px] space-y-2.5 pr-1">
                              {estimatedPoints.map(p => {
                                const st = stations.find(s => s.id === p.stationId);
                                const currVal = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
                                const prevVal = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((sum, r) => sum + r.value, 0);
                                const diff = currVal - prevVal;

                                return (
                                  <div key={p.id} className={`p-3 rounded-lg border flex flex-col gap-1.5 ${darkTheme ? 'bg-slate-900/60 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
                                    <div className="flex items-start justify-between gap-2">
                                      <div>
                                        <div className="text-xs font-bold">{p.name}</div>
                                        <div className="text-[9px] text-slate-500 font-semibold uppercase">{st?.name || 'Вне станций'} • {p.category}</div>
                                      </div>
                                      <span className="text-[8px] font-extrabold uppercase px-1.5 py-0.5 rounded bg-amber-500/10 text-amber-500">
                                        Расчетный
                                      </span>
                                    </div>
                                    <div className="flex items-center justify-between font-mono text-[10px] border-t border-slate-800/35 pt-1.5">
                                      <div>
                                        <span className="text-slate-550">Расход:</span> <strong className={darkTheme ? "text-slate-300" : "text-slate-800"}>{currVal.toLocaleString()} кВт·ч</strong>
                                      </div>
                                      <div className={`font-semibold ${diff > 0 ? 'text-rose-500' : diff < 0 ? 'text-emerald-500' : 'text-slate-450'}`}>
                                        {diff > 0 ? `+${diff.toLocaleString()}` : diff.toLocaleString()} к г/г
                                      </div>
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        </div>
                      </>
                    );
                  })()}
                </div>
              )}

              {reportSubTab === 'losses' && (
                <div className="space-y-6 animate-fadeIn">
                  {/* Заголовок и описание */}
                  <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                    <div className="flex items-start gap-4">
                      <div className="p-3 bg-rose-500/10 text-rose-500 rounded-lg shrink-0">
                        <ZapOff className="w-6 h-6" />
                      </div>
                      <div className="space-y-1">
                        <h4 className="text-sm font-bold">Аналитический аудит технологических потерь электрических сетей участка</h4>
                        <p className={`text-xs leading-relaxed ${darkTheme ? 'text-slate-400' : 'text-slate-600'}`}>
                          Сквозной контроль технологических потерь в силовых трансформаторах, контактных сетях и распределительных фидерах.
                          Данный модуль сопоставляет фактические показатели с нормативами для определения зон повышенного утечки энергии и неэффективного использования мощностей ОАО РЖД.
                        </p>
                      </div>
                    </div>
                  </div>

                  {/* Карточки KPI */}
                  {(() => {
                    const allLossObjectsWithValues = lossObjects.map(lo => {
                      const curr = lossReadings.filter(r => r.lossObjectId === lo.id && r.year === selectedYear && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
                      const prev = lossReadings.filter(r => r.lossObjectId === lo.id && r.year === selectedYear - 1 && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
                      const diff = curr - prev;
                      const pct = prev > 0 ? (diff / prev) * 100 : 0;
                      return { lo, curr, prev, diff, pct };
                    });

                    const totalLossCurr = allLossObjectsWithValues.reduce((sum, x) => sum + x.curr, 0);
                    const totalLossPrev = allLossObjectsWithValues.reduce((sum, x) => sum + x.prev, 0);
                    const totalLossDiff = totalLossCurr - totalLossPrev;
                    const totalLossPct = totalLossPrev > 0 ? (totalLossDiff / totalLossPrev) * 100 : 0;

                    const totalConsumptionCurr = supplyPoints.filter(p => p.isActive).reduce((sum, p) => {
                      return sum + readings.filter(r => r.supplyPointId === p.id && r.year === selectedYear && selectedMonthsList.includes(r.month)).reduce((s, r) => s + r.value, 0);
                    }, 0);
                    const lossRatio = (totalConsumptionCurr + totalLossCurr) > 0 
                      ? (totalLossCurr / (totalConsumptionCurr + totalLossCurr)) * 100 
                      : 0;

                    const maxLossNode = allLossObjectsWithValues.length > 0 
                      ? [...allLossObjectsWithValues].sort((a, b) => b.curr - a.curr)[0] 
                      : null;

                    const anomalousLossNodesCount = allLossObjectsWithValues.filter(x => x.diff > 0).length;

                    // Monthly losses trends for SVG
                    const monthlyLossDataCurr = Array.from({ length: 12 }, (_, i) => {
                      const m = i + 1;
                      return lossReadings
                        .filter(r => r.year === selectedYear && r.month === m)
                        .reduce((sum, r) => sum + r.value, 0);
                    });

                    const monthlyLossDataPrev = Array.from({ length: 12 }, (_, i) => {
                      const m = i + 1;
                      return lossReadings
                        .filter(r => r.year === selectedYear - 1 && r.month === m)
                        .reduce((sum, r) => sum + r.value, 0);
                    });

                    const lossesMaxVal = Math.max(...monthlyLossDataCurr, ...monthlyLossDataPrev, 10);
                    const lossesMaxLimit = Math.ceil(lossesMaxVal * 1.15 / 100) * 100;

                    const lW = 800;
                    const lH = 180;
                    const lPad = 40;

                    const lXOf = (idx: number) => lPad + (idx * (lW - lPad * 2) / 11);
                    const lYOf = (val: number) => {
                      if (lossesMaxLimit <= 0) return lH - lPad;
                      return lH - lPad - (val / lossesMaxLimit) * (lH - lPad * 2);
                    };

                    return (
                      <>
                        {/* KPI Grid */}
                        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                          {/* Карточка 1: Общий объем потерь */}
                          <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm space-y-2`}>
                            <span className="text-[10px] font-bold uppercase tracking-wider text-rose-450 block">Всего потерь сети</span>
                            <div className="text-lg font-mono font-extrabold text-rose-500">{totalLossCurr.toLocaleString()} кВт·ч</div>
                            <div className="flex items-center gap-1.5 text-[10px]">
                              {totalLossDiff > 0 ? (
                                <span className="flex items-center text-rose-500 font-bold"><TrendingUp className="w-3.5 h-3.5 mr-0.5" /> +{totalLossPct.toFixed(2)}%</span>
                              ) : totalLossDiff < 0 ? (
                                <span className="flex items-center text-emerald-500 font-bold"><TrendingDown className="w-3.5 h-3.5 mr-0.5" /> {totalLossPct.toFixed(2)}%</span>
                              ) : (
                                <span className="text-slate-450 font-semibold">Без изменений</span>
                              )}
                              <span className={darkTheme ? "text-slate-400" : "text-slate-500"}>к г/г</span>
                            </div>
                          </div>

                          {/* Карточка 2: Доля потерь */}
                          <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm space-y-2`}>
                            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block">Доля технологических потерь</span>
                            <div className="text-lg font-mono font-extrabold text-amber-500">{lossRatio.toFixed(2)}%</div>
                            <p className="text-[10px] text-slate-500">От общего энергопотребления участка</p>
                          </div>

                          {/* Карточка 3: Узлы с ростом потерь */}
                          <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm space-y-2`}>
                            <span className="text-[10px] font-bold uppercase tracking-wider text-rose-450 block">Рост потерь</span>
                            <div className="text-lg font-mono font-extrabold text-rose-500">{anomalousLossNodesCount} узел(ла)</div>
                            <p className="text-[10px] text-slate-500">С повышенными технологическими потерями год-к-году.</p>
                          </div>

                          {/* Карточка 4: Самый критический узел */}
                          <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm space-y-2`}>
                            <span className="text-[10px] font-bold uppercase tracking-wider text-purple-450 block">Максимальный источник потерь</span>
                            {maxLossNode ? (
                              <>
                                <div className="text-xs font-bold truncate text-slate-300" title={maxLossNode.lo.name}>{maxLossNode.lo.name}</div>
                                <div className="text-[10px] font-mono text-purple-400">
                                  <strong>{maxLossNode.curr.toLocaleString()} кВт·ч</strong> ({((maxLossNode.curr / (totalLossCurr || 1)) * 100).toFixed(1)}% от всех)
                                </div>
                              </>
                            ) : (
                              <div className="text-xs text-slate-500">Данные отсутствуют</div>
                            )}
                          </div>
                        </div>

                        {/* Интерактивный годовой график потерь */}
                        <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                          <div className="flex items-center justify-between gap-4 mb-4">
                            <div>
                              <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400">Динамика суммарных потерь сети (кВт·ч) по месяцам</h4>
                              <p className="text-[10px] text-slate-500">Сравнение годовых трендов потерь для оценки эффективности мероприятий по энергосбережению</p>
                            </div>
                            <div className="flex gap-2.5 text-[9px] font-bold">
                              <span className="flex items-center gap-1 text-sky-400">
                                <span className="w-2 rounded-full h-2 bg-sky-400" /> {selectedYear - 1} г.
                              </span>
                              <span className="flex items-center gap-1 text-rose-500">
                                <span className="w-2 rounded-full h-2 bg-rose-500" /> {selectedYear} г.
                              </span>
                            </div>
                          </div>

                          <div className="h-[210px] relative">
                            <svg viewBox={`0 0 ${lW} ${lH}`} className="w-full h-full select-none">
                              {/* Горизонтальные линии сетки */}
                              {[0, 0.25, 0.5, 0.75, 1].map((ratio, index) => {
                                const val = ratio * lossesMaxLimit;
                                const yPos = lYOf(val);
                                return (
                                  <g key={index}>
                                    <line x1={lPad} y1={yPos} x2={lW - lPad} y2={yPos} stroke={darkTheme ? "#334155" : "#e2e8f0"} strokeDasharray="2 3" />
                                    <text x={lPad - 6} y={yPos + 3} textAnchor="end" fill={darkTheme ? "#64748b" : "#94a3b8"} className="text-[8px] font-mono">
                                      {Math.round(val).toLocaleString()}
                                    </text>
                                  </g>
                                );
                              })}

                              {/* Подписи месяцев по оси X */}
                              {RUSSIAN_MONTHS.map((mName, idx) => {
                                const isCurrentMonthHighlight = selectedMonthsList.includes(idx + 1);
                                return (
                                  <text
                                    key={idx}
                                    x={lXOf(idx)}
                                    y={lH - 12}
                                    textAnchor="middle"
                                    fill={isCurrentMonthHighlight ? (darkTheme ? "#60a5fa" : "#2563eb") : (darkTheme ? "#64748b" : "#94a3b8")}
                                    className={`text-[8px] ${isCurrentMonthHighlight ? "font-extrabold" : "font-semibold"}`}
                                  >
                                    {mName.substring(0, 3)}
                                  </text>
                                );
                              })}

                              {/* Линия прошлого года */}
                              <path
                                d={monthlyLossDataPrev.map((val, idx) => `${idx === 0 ? 'M' : 'L'} ${lXOf(idx)} ${lYOf(val)}`).join(' ')}
                                fill="none"
                                stroke="#38bdf8"
                                strokeWidth="2"
                                strokeDasharray="4 3"
                                strokeLinecap="round"
                              />

                              {/* Линия текущего года */}
                              <path
                                d={monthlyLossDataCurr.map((val, idx) => `${idx === 0 ? 'M' : 'L'} ${lXOf(idx)} ${lYOf(val)}`).join(' ')}
                                fill="none"
                                stroke="#f43f5e"
                                strokeWidth="2.5"
                                strokeLinecap="round"
                              />
                              {monthlyLossDataCurr.map((val, idx) => (
                                <circle
                                  key={`lcurr-${idx}`}
                                  cx={lXOf(idx)}
                                  cy={lYOf(val)}
                                  r={selectedMonthsList.includes(idx + 1) ? "5" : "3.5"}
                                  fill={darkTheme ? "#0f172a" : "#ffffff"}
                                  stroke="#f43f5e"
                                  strokeWidth="2.5"
                                  className="cursor-pointer transition-transform hover:scale-150"
                                  onMouseEnter={() => setReportHoveredPoint({ month: idx + 1, year: selectedYear, value: val })}
                                  onMouseLeave={() => setReportHoveredPoint(null)}
                                />
                              ))}
                            </svg>

                            {/* Тултип графика */}
                            {reportHoveredPoint && (
                              <div className={`absolute bottom-3 right-3 p-2 rounded border text-[10px] shadow-lg backdrop-blur-md z-30 transition-all ${
                                darkTheme ? 'bg-slate-900/95 border-slate-700 text-white' : 'bg-white/95 border-slate-200 text-slate-800'
                              }`}>
                                <span className="font-bold text-rose-400 block border-b border-slate-700/50 pb-0.5 mb-1 text-center">
                                  {RUSSIAN_MONTHS[reportHoveredPoint.month - 1]} {reportHoveredPoint.year} г.
                                </span>
                                <div className="font-mono flex justify-between gap-4">
                                  <span>Потери:</span>
                                  <span className="font-bold text-slate-300">{reportHoveredPoint.value.toLocaleString()} кВт·ч</span>
                                </div>
                              </div>
                            )}
                          </div>
                        </div>

                        {/* Аналитика по Станциям и Узлам */}
                        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                          
                          {/* Слева: Сводная потерь по Станциям */}
                          <div className={`lg:col-span-5 p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm flex flex-col`}>
                            <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-2">Ведомость потерь по станциям</h4>
                            <p className="text-[10px] text-slate-500 mb-4">Суммированные технологические потери по всем объектам в границах станции</p>
                            
                            <div className="overflow-x-auto grow">
                              <table className="w-full text-left text-xs">
                                <thead className={`font-bold ${darkTheme ? 'bg-slate-900/60 text-slate-400' : 'bg-slate-100 text-slate-700'}`}>
                                  <tr>
                                    <th className="p-2">Станция</th>
                                    <th className="p-2 text-right">Потери (кВт·ч)</th>
                                    <th className="p-2 text-right">Динамика</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-800/40">
                                  {(() => {
                                    const stationsLossesCalculations = stations.map(st => {
                                      const stLossObjects = lossObjects.filter(lo => lo.stationId === st.id);
                                      const curr = stLossObjects.reduce((sum, lo) => {
                                        return sum + lossReadings.filter(r => r.lossObjectId === lo.id && r.year === selectedYear && selectedMonthsList.includes(r.month)).reduce((s, r) => s + r.value, 0);
                                      }, 0);
                                      const prev = stLossObjects.reduce((sum, lo) => {
                                        return sum + lossReadings.filter(r => r.lossObjectId === lo.id && r.year === selectedYear - 1 && selectedMonthsList.includes(r.month)).reduce((s, r) => s + r.value, 0);
                                      }, 0);
                                      const diff = curr - prev;
                                      const pct = prev > 0 ? (diff / prev) * 100 : 0;
                                      return { station: st, curr, prev, diff, pct };
                                    }).filter(item => item.curr > 0 || item.prev > 0);

                                    if (stationsLossesCalculations.length === 0) {
                                      return (
                                        <tr>
                                          <td colSpan={3} className="p-4 text-center text-slate-500 italic">
                                            Объекты потерь не привязаны к станциям
                                          </td>
                                        </tr>
                                      );
                                    }

                                    return stationsLossesCalculations.map(sc => {
                                      const isSaving = sc.diff <= 0;
                                      const maxStationLossVal = Math.max(...stationsLossesCalculations.map(s => s.curr), 1);
                                      const pctOfMax = (sc.curr / maxStationLossVal) * 100;

                                      return (
                                        <tr key={sc.station.id} className={darkTheme ? "hover:bg-slate-800/25" : "hover:bg-slate-50"}>
                                          <td className="p-2">
                                            <span className="font-bold block text-[12px]">{sc.station.name}</span>
                                            {/* Sparkbar */}
                                            <div className="w-24 bg-slate-800 h-1 rounded-full overflow-hidden mt-1">
                                              <div className="bg-rose-500 h-full rounded-full" style={{ width: `${pctOfMax}%` }} />
                                            </div>
                                          </td>
                                          <td className="p-2 text-right font-mono font-bold text-rose-400">
                                            {sc.curr.toLocaleString()}
                                            <span className="block text-[8px] text-slate-500 font-semibold">Прошлый: {sc.prev.toLocaleString()}</span>
                                          </td>
                                          <td className="p-2 text-right">
                                            <span className={`inline-flex px-1.5 py-0.5 rounded text-[8.5px] font-bold ${
                                              isSaving ? 'bg-emerald-500/10 text-emerald-400' : 'bg-rose-500/10 text-rose-400'
                                            }`}>
                                              {sc.diff > 0 ? `+${sc.diff.toLocaleString()}` : sc.diff.toLocaleString()} кВт·ч
                                            </span>
                                            <span className={`block text-[8px] font-semibold mt-0.5 ${isSaving ? 'text-emerald-500' : 'text-rose-500'}`}>
                                              {sc.pct > 0 ? `+${sc.pct.toFixed(1)}%` : `${sc.pct.toFixed(1)}%`}
                                            </span>
                                          </td>
                                        </tr>
                                      );
                                    });
                                  })()}
                                </tbody>
                              </table>
                            </div>
                          </div>

                          {/* Справа: Список объектов потерь с Поиском и Сортировкой */}
                          <div className={`lg:col-span-7 p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
                              <div>
                                <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400">Реестр объектов потерь сети</h4>
                                <p className="text-[9px] text-slate-500">Глубокий анализ каждого зарегистрированного узла потерь</p>
                              </div>
                              <div className="flex gap-2">
                                <input
                                  type="text"
                                  value={reportLossSearch}
                                  onChange={(e) => setReportLossSearch(e.target.value)}
                                  placeholder="Поиск по названию/участку..."
                                  className={`text-[10px] px-2.5 py-1 rounded border outline-none w-36 sm:w-44 ${
                                    darkTheme ? 'bg-slate-800 border-slate-700 text-slate-300' : 'bg-slate-50 border-slate-300 text-slate-700'
                                  }`}
                                />
                                <select
                                  value={reportLossSortOrder}
                                  onChange={(e) => setReportLossSortOrder(e.target.value as any)}
                                  className={`text-[10px] font-bold px-2 py-1 rounded border outline-none ${
                                    darkTheme ? 'bg-slate-800 border-slate-700 text-slate-300' : 'bg-slate-50 border-slate-300 text-slate-700'
                                  }`}
                                >
                                  <option value="worst-first">⚠️ Худшие → Лучшие (по росту %)</option>
                                  <option value="best-first">🍏 Лучшие → Худшие (по снижению %)</option>
                                  <option value="volume-desc">⚡ Объем потерь ↓</option>
                                  <option value="volume-asc">🔌 Объем потерь ↑</option>
                                </select>
                              </div>
                            </div>

                            <div className="overflow-x-auto max-h-[400px] overflow-y-auto">
                              <table className="w-full text-left text-xs">
                                <thead className={`sticky top-0 z-10 font-bold ${darkTheme ? 'bg-slate-900 text-slate-400' : 'bg-slate-100 text-slate-600'}`}>
                                  <tr>
                                    <th className="p-2.5">Название узла контроля / Станция</th>
                                    <th className="p-2.5 text-right font-mono">Текущие</th>
                                    <th className="p-2.5 text-right font-mono">Прошлый г.</th>
                                    <th className="p-2.5 text-right font-mono">Отклонение</th>
                                    <th className="p-2.5 text-center">Статус</th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-800/30">
                                  {(() => {
                                    const filteredLossObjects = lossObjects.filter(lo => {
                                      if (!reportLossSearch) return true;
                                      const q = reportLossSearch.toLowerCase();
                                      const stationName = stations.find(s => s.id === lo.stationId)?.name || '';
                                      return (
                                        lo.name.toLowerCase().includes(q) ||
                                        (lo.section && lo.section.toLowerCase().includes(q)) ||
                                        stationName.toLowerCase().includes(q)
                                      );
                                    });

                                    const lossObjectsCalculations = filteredLossObjects.map(lo => {
                                      const curr = lossReadings.filter(r => r.lossObjectId === lo.id && r.year === selectedYear && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
                                       const prev = lossReadings.filter(r => r.lossObjectId === lo.id && r.year === selectedYear - 1 && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
                                      const diff = curr - prev;
                                      const pct = prev > 0 ? (diff / prev) * 100 : 0;
                                      const station = stations.find(s => s.id === lo.stationId);
                                      return { lo, curr, prev, diff, pct, stationName: station?.name || 'Вне станций' };
                                    });

                                    const sortedLossCalculations = [...lossObjectsCalculations].sort((a, b) => {
                                      if (reportLossSortOrder === 'worst-first') return b.diff - a.diff;
                                      if (reportLossSortOrder === 'best-first') return a.diff - b.diff;
                                      if (reportLossSortOrder === 'volume-desc') return b.curr - a.curr;
                                      if (reportLossSortOrder === 'volume-asc') return a.curr - b.curr;
                                      return 0;
                                    });

                                    if (sortedLossCalculations.length === 0) {
                                      return (
                                        <tr>
                                          <td colSpan={5} className="p-8 text-center text-slate-500 italic">
                                            Узлы потерь не найдены по заданным критериям
                                          </td>
                                        </tr>
                                      );
                                    }

                                    return sortedLossCalculations.map((lc, idx) => {
                                      const isSaving = lc.diff <= 0;
                                      const isHighGrowth = lc.diff > 100 && lc.pct > 10;
                                      const isSelected = reportSelectedLossObjectId === lc.lo.id || (!reportSelectedLossObjectId && idx === 0);
                                      return (
                                        <tr 
                                          key={lc.lo.id} 
                                          className={`cursor-pointer transition-all border-l-2 ${
                                            isSelected 
                                              ? (darkTheme ? "bg-indigo-500/10 border-indigo-500" : "bg-indigo-50/70 border-indigo-500")
                                              : (darkTheme ? "hover:bg-slate-700/20 border-transparent" : "hover:bg-slate-50 border-transparent")
                                          }`}
                                          onClick={() => setReportSelectedLossObjectId(lc.lo.id)}
                                        >
                                          <td className="p-2.5">
                                            <div className="flex items-center gap-1.5">
                                               <span className={`text-[10px] ${isSelected ? 'text-indigo-400' : 'text-slate-500'}`}>📊</span>
                                               <div className={`font-bold text-[11px] ${isSelected ? (darkTheme ? 'text-indigo-300 font-extrabold' : 'text-indigo-900 font-extrabold') : (darkTheme ? 'text-slate-200' : 'text-slate-800')}`}>{lc.lo.name}</div>
                                             </div>
                                            <div className="flex items-center flex-wrap gap-1.5 mt-0.5 text-[9px] text-slate-500 uppercase font-semibold">
                                              <span>{lc.stationName}</span>
                                              <span>•</span>
                                              <span className="text-amber-500/90">{lc.lo.section || 'Технологические'}</span>
                                            </div>
                                          </td>
                                          <td className="p-2.5 text-right font-mono font-medium text-rose-400">{lc.curr.toLocaleString()}</td>
                                          <td className="p-2.5 text-right font-mono text-slate-500">{lc.prev.toLocaleString()}</td>
                                          <td className={`p-2.5 text-right font-mono font-extrabold ${isSaving ? 'text-emerald-400' : 'text-rose-400'}`}>
                                            {lc.diff > 0 ? `+${lc.diff.toLocaleString()}` : lc.diff.toLocaleString()}
                                            <span className="block text-[8px] opacity-75 font-semibold">
                                              {lc.pct > 0 ? `+${lc.pct.toFixed(1)}%` : `${lc.pct.toFixed(1)}%`}
                                            </span>
                                          </td>
                                          <td className="p-2.5 text-center">
                                            <span className={`inline-flex px-2 py-0.5 rounded-full text-[8px] font-extrabold uppercase ${
                                              isHighGrowth 
                                                ? 'bg-rose-500/20 text-rose-400 border border-rose-500/30' 
                                                : isSaving 
                                                  ? 'bg-emerald-500/10 text-emerald-400 border border-emerald-500/20' 
                                                  : 'bg-slate-800 text-slate-400'
                                            }`}>
                                              {isHighGrowth ? '⚠️ Критично' : isSaving ? 'Экономия' : 'Стабильно'}
                                            </span>
                                          </td>
                                        </tr>
                                      );
                                    });
                                  })()}
                                </tbody>
                              </table>
                            </div>
                          </div>

                        </div>
                         {/* Детальный график по выбранной потере */}
                         {(() => {
                           const activeLo = lossObjects.find(lo => lo.id === reportSelectedLossObjectId) || lossObjects[0];
                           if (!activeLo) return null;

                           const singleLossDataCurr = Array.from({ length: 12 }, (_, i) => {
                             const m = i + 1;
                             return lossReadings
                               .filter(r => r.lossObjectId === activeLo.id && r.year === selectedYear && r.month === m)
                               .reduce((sum, r) => sum + r.value, 0);
                           });

                           const singleLossDataPrev = Array.from({ length: 12 }, (_, i) => {
                             const m = i + 1;
                             return lossReadings
                               .filter(r => r.lossObjectId === activeLo.id && r.year === selectedYear - 1 && r.month === m)
                               .reduce((sum, r) => sum + r.value, 0);
                           });

                           const singleLossMaxVal = Math.max(...singleLossDataCurr, ...singleLossDataPrev, 10);
                           const singleLossMaxLimit = Math.ceil(singleLossMaxVal * 1.15 / 10) * 10;

                           const sW = 800;
                           const sH = 180;
                           const sPad = 40;

                           const sXOf = (idx: number) => sPad + (idx * (sW - sPad * 2) / 11);
                           const sYOf = (val: number) => {
                             if (singleLossMaxLimit <= 0) return sH - sPad;
                             return sH - sPad - (val / singleLossMaxLimit) * (sH - sPad * 2);
                           };

                           const activeStationName = stations.find(s => s.id === activeLo.stationId)?.name || 'Вне станций';

                           const singleCurrSum = singleLossDataCurr.reduce((a, b) => a + b, 0);
                           const singlePrevSum = singleLossDataPrev.reduce((a, b) => a + b, 0);
                           const singleDiff = singleCurrSum - singlePrevSum;
                           const singlePct = singlePrevSum > 0 ? (singleDiff / singlePrevSum) * 100 : 0;
                           const isSaving = singleDiff <= 0;

                           return (
                             <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm animate-fadeIn mt-6`}>
                               <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-4">
                                 <div className="space-y-1">
                                   <div className="flex items-center gap-2">
                                     <span className="p-1.5 bg-indigo-500/10 text-indigo-400 rounded-md">
                                       <TrendingUp className="w-4 h-4 text-indigo-400" />
                                     </span>
                                     <h4 className="text-xs font-bold uppercase tracking-wider text-indigo-400">
                                       График динамики потерь по выбранному узлу
                                     </h4>
                                   </div>
                                   <h3 className={`text-sm font-bold mt-1 ${darkTheme ? 'text-slate-200' : 'text-slate-800'}`}>
                                     {activeLo.name} <span className="text-[11px] font-normal text-slate-500 uppercase tracking-wide">({activeStationName})</span>
                                   </h3>
                                 </div>
                                 
                                 <div className="flex flex-wrap items-center gap-4 text-xs">
                                   <div className={`px-3 py-1.5 rounded-lg border ${darkTheme ? 'bg-slate-800/60 border-slate-700/40' : 'bg-slate-50 border-slate-250/50'}`}>
                                     <span className="text-[9px] text-slate-500 uppercase font-bold block">Сумма за {selectedYear} г.</span>
                                     <span className="font-mono font-bold text-indigo-500">{singleCurrSum.toLocaleString()} кВт·ч</span>
                                   </div>
                                   <div className={`px-3 py-1.5 rounded-lg border ${darkTheme ? 'bg-slate-800/60 border-slate-700/40' : 'bg-slate-50 border-slate-250/50'}`}>
                                     <span className="text-[9px] text-slate-500 uppercase font-bold block">Отклонение к {selectedYear - 1} г.</span>
                                     <span className={`font-mono font-bold ${isSaving ? 'text-emerald-500' : 'text-rose-500'}`}>
                                       {singleDiff > 0 ? `+${singleDiff.toLocaleString()}` : singleDiff.toLocaleString()} ({singlePct > 0 ? `+${singlePct.toFixed(1)}%` : `${singlePct.toFixed(1)}%`})
                                     </span>
                                   </div>
                                   <div className="flex gap-2.5 text-[9px] font-bold md:self-center">
                                     <span className="flex items-center gap-1 text-teal-500">
                                       <span className="w-2 rounded-full h-2 bg-teal-400" /> {selectedYear - 1} г.
                                     </span>
                                     <span className="flex items-center gap-1 text-indigo-500">
                                       <span className="w-2 rounded-full h-2 bg-indigo-500" /> {selectedYear} г.
                                     </span>
                                   </div>
                                 </div>
                               </div>

                               <div className="h-[210px] relative">
                                 <svg viewBox={`0 0 ${sW} ${sH}`} className="w-full h-full select-none">
                                   {/* Горизонтальные линии сетки */}
                                   {[0, 0.25, 0.5, 0.75, 1].map((ratio, index) => {
                                     const val = ratio * singleLossMaxLimit;
                                     const yPos = sYOf(val);
                                     return (
                                       <g key={index}>
                                         <line x1={sPad} y1={yPos} x2={sW - sPad} y2={yPos} stroke={darkTheme ? "#334155" : "#e2e8f0"} strokeDasharray="2 3" />
                                         <text x={sPad - 6} y={yPos + 3} textAnchor="end" fill={darkTheme ? "#64748b" : "#94a3b8"} className="text-[8px] font-mono">
                                           {Math.round(val).toLocaleString()}
                                         </text>
                                       </g>
                                     );
                                   })}

                                   {/* Подписи месяцев по оси X */}
                                   {RUSSIAN_MONTHS.map((mName, idx) => {
                                     const isCurrentMonthHighlight = selectedMonthsList.includes(idx + 1);
                                     return (
                                       <text
                                         key={idx}
                                         x={sXOf(idx)}
                                         y={sH - 12}
                                         textAnchor="middle"
                                         fill={isCurrentMonthHighlight ? (darkTheme ? "#818cf8" : "#4f46e5") : (darkTheme ? "#64748b" : "#94a3b8")}
                                         className={`text-[8px] ${isCurrentMonthHighlight ? "font-extrabold" : "font-semibold"}`}
                                       >
                                         {mName.substring(0, 3)}
                                       </text>
                                     );
                                   })}

                                   {/* Линия прошлого года */}
                                   <path
                                     d={singleLossDataPrev.map((val, idx) => `${idx === 0 ? 'M' : 'L'} ${sXOf(idx)} ${sYOf(val)}`).join(' ')}
                                     fill="none"
                                     stroke="#2dd4bf"
                                     strokeWidth="2"
                                     strokeDasharray="4 3"
                                     strokeLinecap="round"
                                   />

                                   {/* Линия текущего года */}
                                   <path
                                     d={singleLossDataCurr.map((val, idx) => `${idx === 0 ? 'M' : 'L'} ${sXOf(idx)} ${sYOf(val)}`).join(' ')}
                                     fill="none"
                                     stroke="#6366f1"
                                     strokeWidth="2.5"
                                     strokeLinecap="round"
                                   />
                                   {singleLossDataCurr.map((val, idx) => (
                                     <circle
                                       key={`lsingle-${idx}`}
                                       cx={sXOf(idx)}
                                       cy={sYOf(val)}
                                       r={selectedMonthsList.includes(idx + 1) ? "5" : "3.5"}
                                       fill={darkTheme ? "#0f172a" : "#ffffff"}
                                       stroke="#6366f1"
                                       strokeWidth="2.5"
                                       className="cursor-pointer transition-transform hover:scale-150"
                                       onMouseEnter={() => setReportHoveredSingleLossPoint({ month: idx + 1, year: selectedYear, value: val })}
                                       onMouseLeave={() => setReportHoveredSingleLossPoint(null)}
                                     />
                                   ))}
                                 </svg>

                                 {/* Тултип графика */}
                                 {reportHoveredSingleLossPoint && (
                                   <div className={`absolute bottom-3 right-3 p-2 rounded border text-[10px] shadow-lg backdrop-blur-md z-30 transition-all ${
                                     darkTheme ? 'bg-slate-900/95 border-slate-700 text-white' : 'bg-white/95 border-slate-200 text-slate-800'
                                   }`}>
                                     <span className="font-bold text-indigo-400 block border-b border-slate-700/50 pb-0.5 mb-1 text-center">
                                       {RUSSIAN_MONTHS[reportHoveredSingleLossPoint.month - 1]} {reportHoveredSingleLossPoint.year} г.
                                     </span>
                                     <div className="font-mono flex justify-between gap-4">
                                       <span>Потери:</span>
                                       <span className={darkTheme ? "font-bold text-slate-300" : "font-bold text-slate-700"}>{reportHoveredSingleLossPoint.value.toLocaleString()} кВт·ч</span>
                                     </div>
                                   </div>
                                 )}
                               </div>
                             </div>
                           );
                         })()}

                       </>
                    );
                  })()}
                </div>
              )}

              {reportSubTab === 'slides_presentation' && (
                <div className="space-y-6 animate-fadeIn text-slate-100">
                  {/* Executive Header */}
                  <div className={`p-6 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-6`}>
                    <div className="flex items-start gap-4">
                      <div className="p-3 bg-amber-500/10 text-amber-500 rounded-lg shrink-0">
                        <Presentation className="w-6 h-6 animate-pulse" />
                      </div>
                      <div className="space-y-1">
                        <h4 className={`text-sm font-bold ${darkTheme ? 'text-white' : 'text-slate-800'}`}>Интерактивный конструктор исполнительной презентации PPTX</h4>
                        <p className="text-xs text-slate-400">
                          Экспортируйте результаты сопоставительного энергоаудита в готовые презентационные слайды Microsoft PowerPoint.
                        </p>
                      </div>
                    </div>
                    <button
                      onClick={exportReportToPptx}
                      className="bg-amber-600 hover:bg-amber-700 text-xs px-5 py-2.5 rounded-lg text-white font-bold font-mono transition-all duration-150 flex items-center gap-2 self-start md:self-center shadow-lg hover:shadow shadow-amber-600/10 shrink-0 cursor-pointer border-none outline-none"
                    >
                      <Presentation className="w-4 h-4" strokeWidth={2.5} />
                      Сформировать презентацию (.PPTX)
                    </button>
                  </div>

                  {/* Storyboard Layout Grid */}
                  <div className="space-y-4">
                    <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">Макет слайд-шоу (Раскадровка презентации: 7 слайдов)</h3>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
                      
                      {/* Slide 1: Title */}
                      <div className={`rounded-xl border p-4 flex flex-col justify-between h-[220px] relative overflow-hidden transition-all duration-150 hover:border-slate-500 ${darkTheme ? 'bg-slate-900 border-slate-850' : 'bg-slate-50 border-slate-200'} shadow-sm`}>
                        <div className="space-y-2">
                          <span className="text-[9px] bg-slate-800 text-slate-400 font-bold px-1.5 py-0.5 rounded uppercase">Слайд 1 • Титульный</span>
                          <h5 className="text-[10px] uppercase font-bold text-blue-400 mt-2 font-mono">ОАО РЖД • ЦЕНТР СБЕРЕЖЕНИЯ ЭНЕРГОРЕСУРСОВ</h5>
                          <h4 className={`text-xs font-extrabold leading-snug line-clamp-2 ${darkTheme ? 'text-white' : 'text-slate-850'}`}>ЭНЕРГЕТИЧЕСКИЙ АУДИТ Ж/Д ИНФРАСТРУКТУРЫ УЧАСТКА</h4>
                        </div>
                        <div className="text-[9px] text-slate-500 italic mt-4 font-mono">
                          Учетный период: {selectedPeriodName} {selectedYear} г.
                        </div>
                        <div className="absolute top-2 right-2 text-slate-700/50">
                          <Presentation className="w-12 h-12 stroke-1" />
                        </div>
                      </div>

                      {/* Slide 2: KPI Dashboard */}
                      <div className={`rounded-xl border p-4 flex flex-col justify-between h-[220px] relative overflow-hidden transition-all duration-150 hover:border-slate-500 ${darkTheme ? 'bg-slate-900 border-slate-850' : 'bg-slate-50 border-slate-200'} shadow-sm`}>
                        <div className="space-y-2">
                          <span className="text-[9px] bg-slate-800 text-slate-400 font-bold px-1.5 py-0.5 rounded uppercase">Слайд 2 • Сводные KPI</span>
                          <h4 className={`text-xs font-bold leading-snug mt-2 ${darkTheme ? 'text-white' : 'text-slate-850'}`}>АНАЛИТИКА ЭНЕРГОПОТРЕБЛЕНИЯ УЧАСТКА</h4>
                          <div className="grid grid-cols-2 gap-2 mt-1 font-mono text-[9px]">
                            <div className={`p-1.5 rounded ${darkTheme ? 'bg-slate-850' : 'bg-slate-200/55'}`}>
                              <span className="text-slate-500 block">Расход {selectedYear} г.</span>
                              <span className={`font-bold ${darkTheme ? 'text-white' : 'text-slate-800'}`}>
                                {stations.reduce((sum, st) => {
                                  const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
                                  return sum + readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                                }, 0).toLocaleString()} кВт·ч
                              </span>
                            </div>
                            <div className={`p-1.5 rounded ${darkTheme ? 'bg-slate-850' : 'bg-slate-200/55'}`}>
                              <span className="text-slate-500 block">Прошлый год</span>
                              <span className={`font-bold ${darkTheme ? 'text-slate-300' : 'text-slate-650'}`}>
                                {stations.reduce((sum, st) => {
                                  const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
                                  return sum + readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                                }, 0).toLocaleString()} кВт·ч
                              </span>
                            </div>
                          </div>
                        </div>
                        <div className="text-[9px] text-slate-400 line-clamp-2 mt-1 leading-relaxed">
                          Динамика: <span className={
                            (() => {
                              const curr = stations.reduce((sum, st) => {
                                const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
                                return sum + readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                              }, 0);
                              const prev = stations.reduce((sum, st) => {
                                const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
                                return sum + readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                              }, 0);
                              return curr - prev > 0 ? 'text-rose-500 font-bold' : 'text-emerald-500 font-bold';
                            })()
                          }>
                            {(() => {
                              const curr = stations.reduce((sum, st) => {
                                const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
                                return sum + readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                              }, 0);
                              const prev = stations.reduce((sum, st) => {
                                const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
                                return sum + readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                              }, 0);
                              const diff = curr - prev;
                              const pct = prev > 0 ? (diff / prev) * 100 : 0;
                              return `${diff > 0 ? '+' : ''}${diff.toLocaleString()} кВт·ч (${diff > 0 ? '+' : ''}${pct.toFixed(2)}%)`;
                            })()}
                          </span>
                        </div>
                      </div>

                      {/* Slide 3: Worst Performing Stations */}
                      <div className={`rounded-xl border p-4 flex flex-col justify-between h-[220px] relative overflow-hidden transition-all duration-150 hover:border-slate-500 ${darkTheme ? 'bg-slate-900 border-slate-850' : 'bg-slate-50 border-slate-200'} shadow-sm`}>
                        <div className="space-y-1">
                          <span className="text-[9px] bg-slate-800 text-slate-400 font-bold px-1.5 py-0.5 rounded uppercase">Слайд 3 • Лидеры роста</span>
                          <h4 className="text-xs font-bold text-rose-400 uppercase font-mono mt-1">ЛИДЕРЫ РОСТА РАСХОДА (Worst 5)</h4>
                          
                          <div className="mt-2 space-y-1">
                            {(() => {
                              const diffs = stations.map(st => {
                                const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
                                const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                                const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                                return { name: st.name, diff: curr - prev };
                              }).filter(s => s.diff > 0).sort((a, b) => b.diff - a.diff).slice(0, 3);

                              if (diffs.length === 0) {
                                return <span className="text-[10px] text-emerald-500 block italic">Превышений не обнаружено</span>;
                              }
                              return diffs.map((d, idx) => (
                                <div key={idx} className="flex justify-between items-center text-[9px] border-b border-slate-800/60 pb-1">
                                  <span className={`truncate max-w-[120px] ${darkTheme ? 'text-slate-400' : 'text-slate-600'}`}>{d.name}</span>
                                  <span className="text-rose-500 font-mono">+{d.diff.toLocaleString()} кВт·ч</span>
                                </div>
                              ));
                            })()}
                          </div>
                        </div>
                        <div className="text-[9px] text-slate-500 mt-2">
                          Сортировка: от худших к лучшим. Макс 5 объектов.
                        </div>
                      </div>

                      {/* Slide 4: Best Performing Stations */}
                      <div className={`rounded-xl border p-4 flex flex-col justify-between h-[220px] relative overflow-hidden transition-all duration-150 hover:border-slate-500 ${darkTheme ? 'bg-slate-900 border-slate-850' : 'bg-slate-50 border-slate-200'} shadow-sm`}>
                        <div className="space-y-1">
                          <span className="text-[9px] bg-slate-800 text-slate-400 font-bold px-1.5 py-0.5 rounded uppercase">Слайд 4 • Лидеры экономии</span>
                          <h4 className="text-xs font-bold text-emerald-400 uppercase font-mono mt-1">ЛИДЕРЫ СНИЖЕНИЯ (Best 5)</h4>
                          
                          <div className="mt-2 space-y-1">
                            {(() => {
                              const diffs = stations.map(st => {
                                const spsIds = supplyPoints.filter(p => p.stationId === st.id && p.isActive).map(p => p.id);
                                const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                                const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                                return { name: st.name, diff: curr - prev };
                              }).filter(s => s.diff < 0).sort((a, b) => a.diff - b.diff).slice(0, 3);

                              if (diffs.length === 0) {
                                return <span className="text-[10px] text-slate-500 block italic">Положительной динамики нет</span>;
                              }
                              return diffs.map((d, idx) => (
                                <div key={idx} className="flex justify-between items-center text-[9px] border-b border-slate-800/60 pb-1">
                                  <span className={`truncate max-w-[120px] ${darkTheme ? 'text-slate-400' : 'text-slate-600'}`}>{d.name}</span>
                                  <span className="text-emerald-500 font-mono">{d.diff.toLocaleString()} кВт·ч</span>
                                </div>
                              ));
                            })()}
                          </div>
                        </div>
                        <div className="text-[9px] text-slate-500 mt-2">
                          Сортировка: от лучших к худшим. Макс 5 объектов.
                        </div>
                      </div>

                      {/* Slide 5: Segment Analysis */}
                      <div className={`rounded-xl border p-4 flex flex-col justify-between h-[220px] relative overflow-hidden transition-all duration-150 hover:border-slate-500 ${darkTheme ? 'bg-slate-900 border-slate-850' : 'bg-slate-50 border-slate-200'} shadow-sm`}>
                        <div className="space-y-1">
                          <span className="text-[9px] bg-slate-800 text-slate-400 font-bold px-1.5 py-0.5 rounded uppercase">Слайд 5 • Сегментный анализ</span>
                          <h4 className="text-xs font-bold text-indigo-400 uppercase font-mono mt-1">ТЕХНОЛОГИЧЕСКИЕ КАТЕГОРИИ</h4>
                          
                          <div className="mt-2 space-y-1 font-mono text-[9px]">
                            {["Освещение горловин", "Бытовые нагрузки", "Отопление"].map((cat, idx) => {
                              const catPoints = supplyPoints.filter(p => p.category === cat && p.isActive);
                              const curr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && catPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                              const prev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && catPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                              const diff = curr - prev;
                              return (
                                <div key={idx} className="flex justify-between items-center pb-1 border-b border-slate-800/40">
                                  <span className={`truncate max-w-[100px] ${darkTheme ? 'text-slate-400' : 'text-slate-600'}`}>{cat}</span>
                                  <span className={diff > 0 ? "text-rose-500" : "text-emerald-500"}>
                                    {diff > 0 ? `+${diff.toLocaleString()}` : diff.toLocaleString()}
                                  </span>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                        <div className="text-[9px] text-slate-500 mt-2">
                          Сравнение технологических групп потребителей участка.
                        </div>
                      </div>

                      {/* Slide 6: Network Loss Analysis */}
                      <div className={`rounded-xl border p-4 flex flex-col justify-between h-[220px] relative overflow-hidden transition-all duration-150 hover:border-slate-500 ${darkTheme ? 'bg-slate-900 border-slate-850' : 'bg-slate-50 border-slate-200'} shadow-sm`}>
                        <div className="space-y-1">
                          <span className="text-[9px] bg-slate-800 text-slate-400 font-bold px-1.5 py-0.5 rounded uppercase">Слайд 6 • Сетевые потери</span>
                          <h4 className="text-xs font-bold text-slate-400 uppercase font-mono mt-1">АНАЛИЗ ТЕХНОЛОГИЧЕСКИХ ПОТЕРЬ</h4>
                          
                          <div className="mt-2 text-[9px] font-mono space-y-1 bg-slate-850 p-2 rounded">
                            <div className="flex justify-between">
                              <span className="text-slate-500">Всего потерь:</span>
                              <span className="font-bold text-white">
                                {lossReadings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0).toLocaleString()} кВт·ч
                              </span>
                            </div>
                            <div className="flex justify-between">
                              <span className="text-slate-500">Динамика:</span>
                              {(() => {
                                const cur = lossReadings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
                                const prev = lossReadings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month)).reduce((sum, r) => sum + r.value, 0);
                                const diff = cur - prev;
                                return (
                                  <span className={diff > 0 ? 'text-rose-500 font-bold' : 'text-emerald-500 font-bold'}>
                                    {diff > 0 ? `+${diff.toLocaleString()}` : diff.toLocaleString()}
                                  </span>
                                );
                              })()}
                            </div>
                          </div>
                        </div>
                        <div className="text-[9px] text-slate-500 mt-2">
                          Сводная оценка потерь в ЛЭП и понижающих силовых КТП.
                        </div>
                      </div>

                      {/* Slide 7: Technical recommendations */}
                      <div className={`rounded-xl border p-4 flex flex-col justify-between h-[220px] relative overflow-hidden transition-all duration-150 hover:border-slate-500 ${darkTheme ? 'bg-slate-900 border-slate-850' : 'bg-slate-50 border-slate-200'} shadow-sm`}>
                        <div className="space-y-1">
                          <span className="text-[9px] bg-slate-800 text-slate-400 font-bold px-1.5 py-0.5 rounded uppercase">Слайд 7 • Рекомендации</span>
                          <h4 className="text-xs font-bold text-amber-500 uppercase font-mono mt-1">РЕГЛАМЕНТ ДЕЙСТВИЙ</h4>
                          
                          <ul className="text-[8px] text-slate-400 mt-2 list-disc list-inside space-y-1 font-mono">
                            <li>Автоматизация съема телеметрии</li>
                            <li>Регулировка астрономических реле</li>
                            <li>Ограничение отопления постов ЭЦ</li>
                            <li>Тепловизионная диагностика КТП</li>
                          </ul>
                        </div>
                        <div className="text-[9px] text-slate-500 mt-2">
                          Комплексные адресно-технические предписания.
                        </div>
                      </div>

                    </div>
                  </div>
                </div>
              )}

              {reportSubTab === 'detailed_analysis' && (
                <div className="space-y-6 animate-fadeIn">
                  {/* Executive summary block */}
                  <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white border-slate-200'} shadow-sm`}>
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                      <div className="flex items-start gap-4">
                        <div className="p-3 bg-indigo-500/10 text-indigo-500 rounded-lg shrink-0">
                          <Activity className="w-6 h-6 animate-pulse text-indigo-500" />
                        </div>
                        <div className="space-y-1">
                          <h4 className={`text-sm font-bold ${darkTheme ? 'text-white' : 'text-slate-800'}`}>Подробный сводный анализ энергоэффективности региона</h4>
                          <p className="text-xs text-slate-400">
                            Полное иерархическое дерево: Регион → Станции → Категории → Точки поставки. 
                            Включает накопительные итоги с начала года (YTD), сравнение способов учета и прямое редактирование показателей.
                          </p>
                        </div>
                      </div>
                      <button
                        onClick={exportDetailedReportToDoc}
                        className="bg-emerald-600 hover:bg-emerald-700 text-xs px-4 py-2.5 rounded-lg text-white font-bold transition-all flex items-center gap-2 cursor-pointer shadow-lg self-start md:self-center shrink-0 border-none outline-none"
                      >
                        <FileText className="w-4 h-4" />
                        Сформировать сводный отчет Word (.DOC)
                      </button>
                    </div>
                  </div>

                  {(() => {
                    const activePtIds = supplyPoints.filter(p => p.isActive).map(p => p.id);
                    const totalCurrent = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && activePtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                    const totalPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && activePtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                    const totalDiff = totalCurrent - totalPrev;
                    const totalPct = totalPrev > 0 ? (totalDiff / totalPrev) * 100 : 0;

                    const ytdMonthsList = Array.from({ length: selectedMonth }, (_, i) => i + 1);
                    const totalCurrentYtd = readings.filter(r => r.year === selectedYear && ytdMonthsList.includes(r.month) && activePtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                    const totalPrevYtd = readings.filter(r => r.year === selectedYear - 1 && ytdMonthsList.includes(r.month) && activePtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                    const totalYtdDiff = totalCurrentYtd - totalPrevYtd;
                    const totalYtdPct = totalPrevYtd > 0 ? (totalYtdDiff / totalPrevYtd) * 100 : 0;

                    // Calculation methods
                    const meterPoints = supplyPoints.filter(p => p.isActive && p.calculationMethod !== 'estimated');
                    const estimatedPoints = supplyPoints.filter(p => p.isActive && p.calculationMethod === 'estimated');

                    const meterCurr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && meterPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    const meterPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && meterPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    const meterDiff = meterCurr - meterPrev;
                    const meterPct = meterPrev > 0 ? (meterDiff / meterPrev) * 100 : 0;

                    const estCurr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && estimatedPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    const estPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && estimatedPoints.map(p => p.id).includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
                    const estDiff = estCurr - estPrev;
                    const estPct = estPrev > 0 ? (estDiff / estPrev) * 100 : 0;

                    const totalKwh = meterCurr + estCurr;
                    const meterRatio = totalKwh > 0 ? (meterCurr / totalKwh) * 100 : 0;

                    return (
                      <>
                        {/* Overall KPIs */}
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                          {/* KPI 1: Текущий период */}
                          <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/30 border-slate-700/80' : 'bg-white border-slate-200'} shadow-sm flex flex-col justify-between h-[130px]`}>
                            <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Текущий период ({selectedPeriodName})</span>
                            <div className="mt-2 flex items-baseline gap-1.5">
                              <span className="text-xl font-extrabold font-mono">{totalCurrent.toLocaleString()}</span>
                              <span className="text-[10px] text-slate-500 font-bold">кВт·ч</span>
                            </div>
                            <div className="mt-2 flex items-center justify-between text-xs">
                              <span className="text-slate-400">г/г:</span>
                              {totalDiff > 0 ? (
                                <span className="text-rose-500 font-bold flex items-center gap-0.5"><TrendingUp className="w-3.5 h-3.5" /> +{totalDiff.toLocaleString()} (+{totalPct.toFixed(1)}%)</span>
                              ) : (
                                <span className="text-emerald-500 font-bold flex items-center gap-0.5"><TrendingDown className="w-3.5 h-3.5" /> {totalDiff.toLocaleString()} ({totalPct.toFixed(1)}%)</span>
                              )}
                            </div>
                          </div>

                          {/* KPI 2: Накопительно YTD */}
                          <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/30 border-slate-700/80' : 'bg-white border-slate-200'} shadow-sm flex flex-col justify-between h-[130px]`}>
                            <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">С начала года (YTD накопительный)</span>
                            <div className="mt-2 flex items-baseline gap-1.5">
                              <span className="text-xl font-extrabold font-mono">{totalCurrentYtd.toLocaleString()}</span>
                              <span className="text-[10px] text-slate-500 font-bold">кВт·ч</span>
                            </div>
                            <div className="mt-2 flex items-center justify-between text-xs">
                              <span className="text-slate-400">YTD динамика г/г:</span>
                              {totalYtdDiff > 0 ? (
                                <span className="text-rose-500 font-bold flex items-center gap-0.5"><TrendingUp className="w-3.5 h-3.5" /> +{totalYtdDiff.toLocaleString()} (+{totalYtdPct.toFixed(1)}%)</span>
                              ) : (
                                <span className="text-emerald-500 font-bold flex items-center gap-0.5"><TrendingDown className="w-3.5 h-3.5" /> {totalYtdDiff.toLocaleString()} ({totalYtdPct.toFixed(1)}%)</span>
                              )}
                            </div>
                          </div>

                          {/* KPI 3: Качество учета */}
                          <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-800/30 border-slate-700/80' : 'bg-white border-slate-200'} shadow-sm flex flex-col justify-between h-[130px]`}>
                            <span className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">Индекс коммерческого приборного учета</span>
                            <div className="mt-2">
                              <div className="flex items-center justify-between text-xs font-bold mb-1">
                                <span>Доля ПУ: {meterRatio.toFixed(1)}%</span>
                                <span className={meterRatio >= 90 ? 'text-emerald-500' : 'text-amber-500'}>
                                  {meterRatio >= 90 ? 'Отлично' : 'Требуется модернизация'}
                                </span>
                              </div>
                              <div className={`w-full h-2 rounded-full overflow-hidden ${darkTheme ? 'bg-slate-800' : 'bg-slate-200'}`}>
                                <div className="bg-emerald-500 h-full rounded-full" style={{ width: `${meterRatio}%` }} />
                              </div>
                            </div>
                            <span className="text-[9px] text-slate-500">Доля расчетных потерь: {(100 - meterRatio).toFixed(1)}%</span>
                          </div>
                        </div>

                        {/* Calculation Methods grid */}
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                          <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-900/60 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
                            <div className="flex items-center justify-between mb-3 border-b border-slate-800/50 pb-2">
                              <h5 className="text-xs font-bold uppercase tracking-wider text-blue-400">📊 ПО ПРИБОРАМ УЧЕТА (ПУ)</h5>
                              <span className="text-[10px] text-slate-400">Активных точек: {meterPoints.length}</span>
                            </div>
                            <div className="grid grid-cols-2 gap-4 text-xs font-mono">
                              <div>
                                <span className="text-[10px] block text-slate-500 uppercase font-sans">Текущий расход</span>
                                <strong className="text-sm text-slate-200">{meterCurr.toLocaleString()} кВт·ч</strong>
                              </div>
                              <div>
                                <span className="text-[10px] block text-slate-500 uppercase font-sans">Динамика г/г</span>
                                <strong className={meterDiff > 0 ? 'text-rose-500 text-sm' : 'text-emerald-500 text-sm'}>
                                  {meterDiff > 0 ? `+${meterDiff.toLocaleString()}` : meterDiff.toLocaleString()} ({meterPct.toFixed(1)}%)
                                </strong>
                              </div>
                            </div>
                          </div>

                          <div className={`p-5 rounded-xl border ${darkTheme ? 'bg-slate-900/60 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
                            <div className="flex items-center justify-between mb-3 border-b border-slate-800/50 pb-2">
                              <h5 className="text-xs font-bold uppercase tracking-wider text-amber-500">🧮 РАСЧЕТНЫЙ СПОСОБ (УСТ. МОЩНОСТЬ)</h5>
                              <span className="text-[10px] text-slate-400">Активных точек: {estimatedPoints.length}</span>
                            </div>
                            <div className="grid grid-cols-2 gap-4 text-xs font-mono">
                              <div>
                                <span className="text-[10px] block text-slate-500 uppercase font-sans">Текущий расход</span>
                                <strong className="text-sm text-slate-200">{estCurr.toLocaleString()} кВт·ч</strong>
                              </div>
                              <div>
                                <span className="text-[10px] block text-slate-500 uppercase font-sans">Динамика г/г</span>
                                <strong className={estDiff > 0 ? 'text-rose-500 text-sm' : 'text-emerald-500 text-sm'}>
                                  {estDiff > 0 ? `+${estDiff.toLocaleString()}` : estDiff.toLocaleString()} ({estPct.toFixed(1)}%)
                                </strong>
                              </div>
                            </div>
                          </div>
                        </div>

                        {/* Hierarchical Stations Breakdown */}
                        <div className="space-y-4">
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-2 border-b border-slate-700/60">
                            <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">Иерархический список станций ({stations.length})</h3>
                            
                            <div className="flex items-center gap-3">
                              <button
                                onClick={() => {
                                  const allExp: Record<string, boolean> = {};
                                  stations.forEach(s => allExp[s.id] = true);
                                  setExpandedDetailedStations(allExp);
                                }}
                                className={`px-2.5 py-1.5 text-[10px] font-bold rounded border ${
                                  darkTheme ? 'bg-slate-800 hover:bg-slate-750 text-slate-300 border-slate-700' : 'bg-white hover:bg-slate-50 border-slate-200 text-slate-700'
                                }`}
                              >
                                Раскрыть все станции
                              </button>
                              <button
                                onClick={() => setExpandedDetailedStations({})}
                                className={`px-2.5 py-1.5 text-[10px] font-bold rounded border ${
                                  darkTheme ? 'bg-slate-800 hover:bg-slate-750 text-slate-300 border-slate-700' : 'bg-white hover:bg-slate-50 border-slate-200 text-slate-700'
                                }`}
                              >
                                Свернуть все
                              </button>
                            </div>
                          </div>

                          <div className="space-y-4">
                            {stations.map(st => {
                              const isExpanded = expandedDetailedStations[st.id];
                              const stPoints = supplyPoints.filter(p => p.stationId === st.id && p.isActive);
                              const stPtIds = stPoints.map(p => p.id);

                              const stCurr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && stPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                              const stPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && stPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                              const stDiff = stCurr - stPrev;
                              const stPct = stPrev > 0 ? (stDiff / stPrev) * 100 : 0;

                              const stCurrYtd = readings.filter(r => r.year === selectedYear && ytdMonthsList.includes(r.month) && stPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                              const stPrevYtd = readings.filter(r => r.year === selectedYear - 1 && ytdMonthsList.includes(r.month) && stPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                              const stDiffYtd = stCurrYtd - stPrevYtd;
                              const stPctYtd = stPrevYtd > 0 ? (stDiffYtd / stPrevYtd) * 100 : 0;

                              return (
                                <div 
                                  key={st.id} 
                                  className={`rounded-xl border transition-all ${
                                    isExpanded 
                                      ? (darkTheme ? 'bg-slate-900/85 border-slate-700' : 'bg-slate-50/50 border-slate-300 shadow-sm')
                                      : (darkTheme ? 'bg-slate-900/30 border-slate-800 hover:bg-slate-900/55' : 'bg-white border-slate-200 hover:bg-slate-50')
                                  }`}
                                >
                                  {/* Header card of station */}
                                  <div 
                                    onClick={() => setExpandedDetailedStations(prev => ({ ...prev, [st.id]: !prev[st.id] }))}
                                    className="p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4 cursor-pointer select-none"
                                  >
                                    <div className="flex items-center gap-3">
                                      <div className={`p-2 rounded-lg ${stDiff > 0 ? 'bg-rose-500/10 text-rose-400' : 'bg-emerald-500/10 text-emerald-400'}`}>
                                        <Train className="w-5 h-5" />
                                      </div>
                                      <div>
                                        <h4 className="font-extrabold text-sm">{st.name}</h4>
                                        <p className="text-[10px] text-slate-500">Участок контроля: <span className="font-bold">{st.section}</span> • Точек: {stPoints.length}</p>
                                      </div>
                                    </div>

                                    <div className="flex flex-wrap items-center gap-6 text-xs font-mono">
                                      <div>
                                        <span className="text-[9px] block text-slate-500 uppercase font-sans">За период ({selectedPeriodName})</span>
                                        <span className="font-bold">{stCurr.toLocaleString()} кВт·ч</span>
                                        <span className={`inline-block ml-1.5 text-[10px] font-bold ${stDiff > 0 ? 'text-rose-500' : 'text-emerald-500'}`}>
                                          {stDiff > 0 ? `+${stPct.toFixed(1)}%` : `${stPct.toFixed(1)}%`}
                                        </span>
                                      </div>

                                      <div className="border-l border-slate-800/60 pl-6">
                                        <span className="text-[9px] block text-slate-500 uppercase font-sans">Накопительно YTD</span>
                                        <span className="font-bold">{stCurrYtd.toLocaleString()} кВт·ч</span>
                                        <span className={`inline-block ml-1.5 text-[10px] font-bold ${stDiffYtd > 0 ? 'text-rose-500' : 'text-emerald-500'}`}>
                                          {stDiffYtd > 0 ? `+${stPctYtd.toFixed(1)}%` : `${stPctYtd.toFixed(1)}%`}
                                        </span>
                                      </div>

                                      <div className="flex items-center justify-center w-8 h-8 rounded-full hover:bg-slate-800/40 shrink-0 transition-colors">
                                        <ChevronRight className={`w-5 h-5 text-slate-400 transition-transform duration-200 ${isExpanded ? 'rotate-90 text-blue-400' : ''}`} />
                                      </div>
                                    </div>
                                  </div>

                                  {/* Expanded Area: Categories & Points details */}
                                  {isExpanded && (
                                    <div className={`p-4 border-t ${darkTheme ? 'border-slate-800' : 'border-slate-200'} space-y-4`}>
                                      {/* Category sub-table summary */}
                                      <div className="space-y-2">
                                        <h5 className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Свод по технологическим категориям</h5>
                                        <div className="overflow-x-auto rounded-lg border border-slate-800/40">
                                          <table className="w-full text-left text-[11px] font-mono">
                                            <thead className={darkTheme ? 'bg-slate-950 text-slate-400' : 'bg-slate-100 text-slate-700'}>
                                              <tr>
                                                <th className="p-2 font-sans font-bold">Категория</th>
                                                <th className="p-2 text-right">Текущий период (кВт·ч)</th>
                                                <th className="p-2 text-right">Прошлый период (кВт·ч)</th>
                                                <th className="p-2 text-right">Изменение</th>
                                                <th className="p-2 text-right font-sans font-bold">Динамика (%)</th>
                                              </tr>
                                            </thead>
                                            <tbody className="divide-y divide-slate-800/30">
                                              {categories.map(cat => {
                                                const catPoints = stPoints.filter(p => p.category === cat);
                                                if (catPoints.length === 0) return null;
                                                const catPtIds = catPoints.map(p => p.id);

                                                const cCurr = readings.filter(r => r.year === selectedYear && selectedMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                                                const cPrev = readings.filter(r => r.year === selectedYear - 1 && selectedMonthsList.includes(r.month) && catPtIds.includes(r.supplyPointId)).reduce((s, r) => s + r.value, 0);
                                                const cDiff = cCurr - cPrev;
                                                const cPct = cPrev > 0 ? (cDiff / cPrev) * 100 : 0;

                                                return (
                                                  <tr key={cat} className={darkTheme ? 'hover:bg-slate-800/20' : 'hover:bg-slate-50'}>
                                                    <td className="p-2 font-sans font-bold text-slate-300">{cat}</td>
                                                    <td className="p-2 text-right">{cCurr.toLocaleString()}</td>
                                                    <td className="p-2 text-right text-slate-500">{cPrev.toLocaleString()}</td>
                                                    <td className={`p-2 text-right font-semibold ${cDiff > 0 ? 'text-rose-400' : 'text-emerald-400'}`}>
                                                      {cDiff > 0 ? `+${cDiff.toLocaleString()}` : cDiff.toLocaleString()}
                                                    </td>
                                                    <td className="p-2 text-right font-sans font-bold">
                                                      <span className={`px-2 py-0.5 rounded text-[10px] ${
                                                        cDiff > 0 
                                                          ? 'bg-rose-500/10 text-rose-400' 
                                                          : 'bg-emerald-500/10 text-emerald-400'
                                                      }`}>
                                                        {cDiff > 0 ? `+${cPct.toFixed(1)}%` : `${cPct.toFixed(1)}%`}
                                                      </span>
                                                    </td>
                                                  </tr>
                                                );
                                              }).filter(Boolean)}
                                            </tbody>
                                          </table>
                                        </div>
                                      </div>

                                      {/* Points list detailed table */}
                                      <div className="space-y-2">
                                        <h5 className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Детализация по точкам поставки (ТП) и управление показателями</h5>
                                        <div className="overflow-x-auto rounded-lg border border-slate-800/40">
                                          <table className="w-full text-left text-[11px] font-mono">
                                            <thead className={darkTheme ? 'bg-slate-950 text-slate-400' : 'bg-slate-100 text-slate-700'}>
                                              <tr>
                                                <th className="p-2.5 font-sans font-bold">Точка поставки (ТП)</th>
                                                <th className="p-2.5 font-sans font-bold">Категория</th>
                                                <th className="p-2.5 font-sans font-bold">Учет</th>
                                                <th className="p-2.5 text-right">Период (кВт·ч)</th>
                                                <th className="p-2.5 text-right">Прошлый г.</th>
                                                <th className="p-2.5 text-right">YTD Текущий</th>
                                                <th className="p-2.5 text-right">YTD Прошлый</th>
                                                <th className="p-2.5 text-right">YTD Динамика</th>
                                                <th className="p-2.5 text-center font-sans font-bold">Действия</th>
                                              </tr>
                                            </thead>
                                            <tbody className="divide-y divide-slate-800/30">
                                              {stPoints.map(p => {
                                                const ptCurr = readings.find(r => r.supplyPointId === p.id && r.year === selectedYear && r.month === selectedMonth)?.value || 0;
                                                const ptPrev = readings.find(r => r.supplyPointId === p.id && r.year === selectedYear - 1 && r.month === selectedMonth)?.value || 0;

                                                const ptCurrYtd = readings.filter(r => r.year === selectedYear && ytdMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((s, r) => s + r.value, 0);
                                                const ptPrevYtd = readings.filter(r => r.year === selectedYear - 1 && ytdMonthsList.includes(r.month) && r.supplyPointId === p.id).reduce((s, r) => s + r.value, 0);
                                                const ptDiffYtd = ptCurrYtd - ptPrevYtd;
                                                const ptPctYtd = ptPrevYtd > 0 ? (ptDiffYtd / ptPrevYtd) * 100 : 0;

                                                const calcMethodLabel = p.calculationMethod === 'estimated' ? 'Расчетный' : 'ПУ';

                                                return (
                                                  <tr key={p.id} className={darkTheme ? 'hover:bg-slate-800/20' : 'hover:bg-slate-50'}>
                                                    <td className="p-2.5 font-sans font-bold text-slate-200">{p.name}</td>
                                                    <td className="p-2.5 font-sans text-slate-400">{p.category}</td>
                                                    <td className="p-2.5">
                                                      <span className={`px-1.5 py-0.5 rounded text-[9px] font-bold uppercase ${
                                                        p.calculationMethod === 'estimated' 
                                                          ? 'bg-amber-500/15 text-amber-500' 
                                                          : 'bg-blue-500/15 text-blue-400'
                                                      }`}>
                                                        {calcMethodLabel}
                                                      </span>
                                                    </td>
                                                    <td className="p-2.5 text-right font-bold text-slate-100">{ptCurr.toLocaleString()}</td>
                                                    <td className="p-2.5 text-right text-slate-500">{ptPrev.toLocaleString()}</td>
                                                    <td className="p-2.5 text-right">{ptCurrYtd.toLocaleString()}</td>
                                                    <td className="p-2.5 text-right text-slate-500">{ptPrevYtd.toLocaleString()}</td>
                                                    <td className={`p-2.5 text-right font-extrabold ${ptDiffYtd > 0 ? 'text-rose-400' : 'text-emerald-400'}`}>
                                                      {ptDiffYtd > 0 ? `+${ptDiffYtd.toLocaleString()}` : ptDiffYtd.toLocaleString()}
                                                      <span className="block text-[8px] opacity-75 font-semibold">
                                                        {ptDiffYtd > 0 ? `+${ptPctYtd.toFixed(1)}%` : `${ptPctYtd.toFixed(1)}%`}
                                                      </span>
                                                    </td>
                                                    <td className="p-2.5 text-center font-sans">
                                                      <button
                                                        onClick={() => handleEditPointFromDetailed(p)}
                                                        className="px-2 py-1 bg-blue-600/15 text-blue-400 hover:bg-blue-600 hover:text-white border border-blue-500/20 hover:border-transparent rounded font-bold text-[10px] transition-all cursor-pointer"
                                                      >
                                                        Редактировать
                                                      </button>
                                                    </td>
                                                  </tr>
                                                );
                                              })}
                                            </tbody>
                                          </table>
                                        </div>
                                      </div>
                                    </div>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </>
                    );
                  })()}
                </div>
              )}

            </div>
          )}

          {/* TAB 6: СПРАВОЧНИКИ (REF BOOKS - Section 7 & 8) */}
          {activeTab === 'directories' && (
            <div className="space-y-6">
              
              <div className="space-y-3">
                <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400">Все точки поставки (Глобальный реестр)</h3>
                <input
                  type="text"
                  value={pointSearch}
                  onChange={(e) => setPointSearch(e.target.value)}
                  placeholder="Фильтр фидеров..."
                  className={`w-full text-xs px-3 py-2 rounded-lg border outline-none ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
                />

                <div className="overflow-x-auto rounded-lg border border-slate-800">
                  <table className="w-full text-xs text-left">
                    <thead className="bg-slate-800 text-slate-400">
                      <tr>
                        <th className="p-2.5">Название ТП</th>
                        <th className="p-2.5">Станция</th>
                        <th className="p-2.5">Категория нагрузок</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-850">
                      {pointsListFiltered.map(p => (
                        <tr 
                          key={p.id} 
                          onClick={() => {
                            setSelectedStationId(p.stationId);
                            setSelectedSupplyPointId(p.id);
                            setActiveTab('stations');
                          }}
                          className={`cursor-pointer transition-colors ${
                            p.isActive 
                              ? (darkTheme ? 'hover:bg-slate-800' : 'hover:bg-slate-100') 
                              : 'opacity-45 hover:opacity-100'
                          }`}
                        >
                          <td className="p-2.5 font-bold">
                            <span className="hover:text-blue-400 transition-colors">{p.name}</span>
                          </td>
                          <td className="p-2.5 text-slate-350">{stations.find(s => s.id === p.stationId)?.name || '—'}</td>
                          <td className="p-2.5"><span className="bg-slate-900 border border-slate-700/50 px-2 py-0.5 rounded text-blue-400 font-bold">{p.category}</span></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Editable categories */}
              <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white'} space-y-3`}>
                <span className="text-xs font-bold uppercase text-slate-400 block">Пользовательские группы нагрузок</span>
                
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={categoryInput}
                    onChange={(e) => setCategoryInput(e.target.value)}
                    placeholder="Новая категория нагрузок..."
                    className={`text-xs px-3 py-1.5 rounded-lg border grow outline-none ${darkTheme ? 'bg-slate-900 border-slate-700 text-white' : 'bg-white'}`}
                  />
                  <button onClick={handleAddCategory} className="bg-blue-600 px-4 text-xs text-white rounded font-bold">Добавить</button>
                </div>

                <div className="flex flex-wrap gap-2 pt-1">
                  {categories.map(c => (
                    <div key={c} className="bg-slate-900 px-2.5 py-1 rounded text-[11px] text-slate-300 border border-slate-800 flex items-center gap-1.5 font-bold">
                      <span>{c}</span>
                      <button 
                        type="button"
                        title="Удалить категорию нагрузки"
                        onClick={() => handleDeleteCategory(c)} 
                        className="text-rose-400 hover:text-rose-300 font-extrabold cursor-pointer bg-slate-800 hover:bg-slate-700 rounded-full w-3.5 h-3.5 flex items-center justify-center text-[8px] transition-all"
                      >
                        ✕
                      </button>
                    </div>
                  ))}
                </div>
              </div>

            </div>
          )}

          {/* TAB 7: НАСТРОЙКИ (SETTINGS - Section 15 & 16) */}
          {activeTab === 'settings' && (
            <div className="space-y-6 max-w-2xl">
              <h3 className="text-xs font-bold uppercase tracking-wider text-slate-400">Резервное копирование, экспорт и сброс</h3>
              
              <div className={`p-5 rounded-xl border space-y-3.5 ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white'}`}>
                <div className="flex items-center justify-between border-b pb-3 border-slate-700/60">
                  <div>
                    <h4 className="text-xs font-bold uppercase text-slate-300">Резервное копирование файлов</h4>
                    <p className="text-[10px] text-slate-400 mt-0.5">Автоматическая запись слепков БД при любых изменениях показателей ЖД</p>
                  </div>
                  <label className="relative inline-flex items-center cursor-pointer">
                    <input
                      type="checkbox"
                      checked={autoBackupEnabled}
                      onChange={(e) => {
                        const val = e.target.checked;
                        setAutoBackupEnabled(val);
                        localStorage.setItem('st_auto_backup', val ? 'true' : 'false');
                      }}
                      className="sr-only peer"
                    />
                    <div className="w-9 h-5 bg-slate-700 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-blue-600"></div>
                    <span className="ml-2.5 text-xs font-semibold text-slate-300">{autoBackupEnabled ? 'ВКЛ' : 'ВЫКЛ'}</span>
                  </label>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  
                  <button onClick={downloadReportTemplate} className="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold py-2.5 px-4 rounded-lg flex items-center gap-2 justify-center">
                    <Download className="w-4 h-4" />
                    <span>Скачать шаблон ведомости .XLSX</span>
                  </button>

                  <button onClick={downloadDbDump} className="bg-blue-600 hover:bg-blue-500 text-white text-xs font-bold py-2.5 px-4 rounded-lg flex items-center gap-2 justify-center">
                    <Download className="w-4 h-4" />
                    <span>Скачать резервный JSON-файл</span>
                  </button>

                  <button onClick={handleCreateManualBackup} className="bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold py-2.5 px-4 rounded-lg flex items-center gap-2 justify-center col-span-1">
                    <span>➕ Создать резервную копию</span>
                  </button>

                  <label className="bg-slate-800 hover:bg-slate-700 border border-slate-700 text-white text-xs font-bold py-2.5 px-4 rounded-lg flex items-center gap-2 justify-center cursor-pointer text-center">
                    <span>📂 Восстановить из файла .JSON</span>
                    <input type="file" accept=".json" onChange={handleImportBackupJsonFile} className="hidden" />
                  </label>

                  <button onClick={handleClearAllReadings} className="border border-amber-500/25 text-amber-400 bg-amber-500/10 hover:bg-amber-500/20 text-xs font-bold py-2.5 px-4 rounded-lg flex items-center gap-2 justify-center md:col-span-2">
                    <Trash2 className="w-4 h-4" />
                    <span>Стереть все показания (оставить только нули на графиках)</span>
                  </button>

                  <button onClick={handleDbReset} className="border border-rose-500/25 text-rose-400 bg-rose-500/10 hover:bg-rose-500/20 text-xs font-bold py-2.5 px-4 rounded-lg flex items-center gap-2 justify-center md:col-span-2">
                    <RefreshCw className="w-4 h-4" />
                    <span>Сбросить локальные изменения к демонстрационным</span>
                  </button>

                </div>
              </div>

              {/* Backups List UI Panel */}
              <div className={`p-5 rounded-xl border space-y-3.5 ${darkTheme ? 'bg-slate-800/40 border-slate-700' : 'bg-white'}`}>
                <h4 className="text-xs font-bold uppercase text-slate-350">Точки локального восстановления</h4>
                <div className="max-h-[220px] overflow-y-auto space-y-2 pr-1 divide-y divide-slate-800">
                  {backupsList.length > 0 ? (
                    backupsList.map((bk) => (
                      <div key={bk.id} className="flex items-center justify-between py-2 text-xs gap-4 first:pt-0">
                        <div className="space-y-0.5">
                          <span className="font-bold block text-slate-205">{bk.name}</span>
                          <span className="text-[10px] text-slate-450 block font-mono">Сохранено: {bk.timestamp}</span>
                        </div>
                        <div className="flex items-center gap-2 shrink-0">
                          <button
                            onClick={() => handleRestoreFromBackup(bk)}
                            className="bg-amber-600 hover:bg-amber-500 text-white text-[10px] font-bold px-2.5 py-1 rounded"
                          >
                            Восстановить
                          </button>
                          <button
                            onClick={() => handleDeleteBackup(bk.id)}
                            className="bg-slate-850 hover:bg-rose-900 hover:text-rose-200 border border-slate-800 text-rose-405 text-[10px] px-2 py-1 rounded"
                          >
                            Удалить
                          </button>
                        </div>
                      </div>
                    ))
                  ) : (
                    <p className="text-xs text-slate-450 italic py-2 text-center">Точки резервного копирования отсутствуют. Измените показания или сделайте резервную копию вручную.</p>
                  )}
                </div>
              </div>

              <div className="p-4 bg-slate-900 text-[10px] text-slate-400 rounded-lg space-y-1.5 leading-relaxed font-mono">
                <span className="text-white font-bold block uppercase mb-1">Спецификация Huawei MatePad Workspace</span>
                <span>• Оптимизация экрана: Горизонтальный планшет (16:10 / 11.5")</span>
                <span>• Поддержка СУБД: Offline (LocalStorage HTML5)</span>
                <span>• Лимиты: Конфигурация локальных ведомостей ЖД без ограничений по КБ</span>
              </div>
            </div>
          )}

        </main>

      </div>

      {/* FOOTER */}
      <footer className={`py-2 text-center text-[10px] border-t ${darkTheme ? 'bg-[#1e293b]/40 border-slate-800 text-slate-500' : 'bg-white border-slate-200 text-slate-500'}`}>
        <span>ОАО «РЖД» • Департамент Электроэнергетики ж/д • 2026 г.</span>
      </footer>

      {/* --- ADD/EDIT STATIONS MODAL DIALOG --- */}
      {stationModal?.isOpen && (
        <div className="fixed inset-0 bg-black/75 flex items-center justify-center p-4 z-50 backdrop-blur-sm">
          <form onSubmit={handleSaveStation} className={`w-full max-w-xs rounded-xl border p-4 space-y-3.5 ${darkTheme ? 'bg-slate-900 border-slate-800' : 'bg-white text-slate-850'}`}>
            <h3 className="font-bold text-xs border-b pb-1.5 border-slate-700">
              {stationModal.mode === 'add' ? 'Создать станцию' : 'Параметры станции'}
            </h3>
            <div className="space-y-1">
              <label className="text-[10px] uppercase font-bold text-slate-400">Название станции:</label>
              <input
                type="text"
                required
                value={stationModal.name}
                onChange={(e) => setStationModal({ ...stationModal, name: e.target.value })}
                className={`w-full text-xs px-2.5 py-1.5 border outline-none rounded ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
              />
            </div>
            <div className="space-y-1">
              <label className="text-[10px] uppercase font-bold text-slate-400">Категория нагрузки:</label>
              <select
                value={stationModal.section}
                onChange={(e) => setStationModal({ ...stationModal, section: e.target.value })}
                className={`w-full text-xs px-2.5 py-1.5 border outline-none rounded ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
              >
                <option value="">-- Выберите категорию нагрузки --</option>
                {categories.map(c => (
                  <option key={c} value={c}>{c}</option>
                ))}
              </select>
            </div>
            <div className="space-y-1">
              <label className="text-[10px] uppercase font-bold text-slate-400">Примечания:</label>
              <textarea
                value={stationModal.note}
                onChange={(e) => setStationModal({ ...stationModal, note: e.target.value })}
                rows={2}
                className={`w-full text-xs px-2.5 py-1.5 border outline-none rounded resize-none ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
              />
            </div>
            <div className="flex gap-2 justify-end text-xs pt-1">
              <button type="button" onClick={() => setStationModal(null)} className="px-3 py-1 bg-slate-800 rounded font-semibold text-slate-350">Отмена</button>
              <button type="submit" className="px-3 py-1 bg-blue-600 hover:bg-blue-500 rounded font-bold text-white">Записать</button>
            </div>
          </form>
        </div>
      )}

      {/* --- ADD/EDIT SUPPLY POINTS MODAL DIALOG --- */}
      {pointModal?.isOpen && (
        <div className="fixed inset-0 bg-black/75 flex items-center justify-center p-4 z-50 backdrop-blur-sm">
          <form onSubmit={handleSavePoint} className={`w-full max-w-xs rounded-xl border p-4 space-y-3.5 ${darkTheme ? 'bg-slate-900 border-slate-800' : 'bg-white text-slate-850'}`}>
            <h3 className="font-bold text-xs border-b pb-1.5 border-slate-700">
              {pointModal.mode === 'add' ? 'Создать точку ТП' : 'Параметры ТП'}
            </h3>
            <div className="space-y-1">
              <label className="text-[10px] uppercase font-bold text-slate-400">Название ТП:</label>
              <input
                type="text"
                required
                value={pointModal.name}
                onChange={(e) => setPointModal({ ...pointModal, name: e.target.value })}
                className={`w-full text-xs px-2.5 py-1.5 border outline-none rounded ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
              />
            </div>
            <div className="space-y-1">
              <label className="text-[10px] uppercase font-bold text-slate-400">Инженерная группа нагрузок:</label>
              <select
                value={pointModal.category}
                onChange={(e) => setPointModal({ ...pointModal, category: e.target.value })}
                className={`w-full text-xs px-2 py-1.5 border outline-none rounded ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
              >
                {categories.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            <div className="space-y-1">
              <label className="text-[10px] uppercase font-bold text-slate-400">Способ расчета:</label>
              <select
                value={pointModal.calculationMethod || 'meter'}
                onChange={(e) => setPointModal({ ...pointModal, calculationMethod: e.target.value as 'meter' | 'estimated' })}
                className={`w-full text-xs px-2 py-1.5 border outline-none rounded ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
              >
                <option value="meter">📈 По прибору учета</option>
                <option value="estimated">🧮 Расчетный способ</option>
              </select>
            </div>
            
            <div className="grid grid-cols-2 gap-2">
              <div className="space-y-1">
                <label className="text-[9px] uppercase font-bold text-slate-450">Расход {selectedYear} г. ({RUSSIAN_MONTHS[selectedMonth - 1]}):</label>
                <input
                  type="number"
                  required
                  value={pointModal.currKwh !== undefined ? pointModal.currKwh : ''}
                  onChange={(e) => setPointModal({ ...pointModal, currKwh: Number(e.target.value) })}
                  className={`w-full text-xs px-2 py-1.5 border outline-none rounded font-mono ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[9px] uppercase font-bold text-slate-455">Расход {selectedYear - 1} г. ({RUSSIAN_MONTHS[selectedMonth - 1]}):</label>
                <input
                  type="number"
                  required
                  value={pointModal.prevKwh !== undefined ? pointModal.prevKwh : ''}
                  onChange={(e) => setPointModal({ ...pointModal, prevKwh: Number(e.target.value) })}
                  className={`w-full text-xs px-2 py-1.5 border outline-none rounded font-mono ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
                />
              </div>
            </div>

            <div className="space-y-1">
              <label className="text-[10px] uppercase font-bold text-slate-400">Заметки / Описание ТП:</label>
              <input
                type="text"
                value={pointModal.note}
                onChange={(e) => setPointModal({ ...pointModal, note: e.target.value })}
                className={`w-full text-xs px-2.5 py-1.5 border outline-none rounded ${darkTheme ? 'bg-slate-800 border-slate-700 text-white' : 'bg-white'}`}
              />
            </div>
            <div className="flex gap-2 justify-end text-xs pt-1">
              <button type="button" onClick={() => setPointModal(null)} className="px-3 py-1 bg-slate-800 rounded font-semibold text-slate-350">Отмена</button>
              <button type="submit" className="px-3 py-1 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded">Внести</button>
            </div>
          </form>
        </div>
      )}
      {(() => {
        // Defined helper for trend SVG charts
        const renderCustomSvgTrendChart = (targetId: string, type: 'station' | 'supply_point' | 'loss') => {
          let currValues: number[] = Array(12).fill(0);
          let prevValues: number[] = Array(12).fill(0);

          if (type === 'station') {
            const spsIds = supplyPoints.filter(p => p.stationId === targetId).map(p => p.id);
            for (let m = 1; m <= 12; m++) {
              currValues[m - 1] = readings.filter(r => r.year === selectedYear && r.month === m && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
              prevValues[m - 1] = readings.filter(r => r.year === selectedYear - 1 && r.month === m && spsIds.includes(r.supplyPointId)).reduce((sum, r) => sum + r.value, 0);
            }
          } else if (type === 'supply_point') {
            for (let m = 1; m <= 12; m++) {
              const currR = readings.find(r => r.supplyPointId === targetId && r.year === selectedYear && r.month === m);
              const prevR = readings.find(r => r.supplyPointId === targetId && r.year === selectedYear - 1 && r.month === m);
              currValues[m - 1] = currR ? currR.value : 0;
              prevValues[m - 1] = prevR ? prevR.value : 0;
            }
          } else if (type === 'loss') {
            for (let m = 1; m <= 12; m++) {
              const currR = lossReadings.find(r => r.lossObjectId === targetId && r.year === selectedYear && r.month === m);
              const prevR = lossReadings.find(r => r.lossObjectId === targetId && r.year === selectedYear - 1 && r.month === m);
              currValues[m - 1] = currR ? currR.value : 0;
              prevValues[m - 1] = prevR ? prevR.value : 0;
            }
          }

          const allValues = [...currValues, ...prevValues];
          const maxVal = Math.max(...allValues, 100) * 1.15;
          const minVal = Math.max(0, Math.min(...allValues) * 0.85);
          const range = maxVal - minVal;

          const width = 400;
          const height = 150;
          const padding = { top: 15, right: 15, bottom: 25, left: 45 };

          const getX = (index: number) => {
            return padding.left + (index * (width - padding.left - padding.right)) / 11;
          };

          const getY = (val: number) => {
            if (range === 0) return padding.top + (height - padding.top - padding.bottom) / 2;
            return height - padding.bottom - ((val - minVal) * (height - padding.top - padding.bottom)) / range;
          };

          const currPath = currValues.map((val, idx) => `${idx === 0 ? 'M' : 'L'} ${getX(idx)} ${getY(val)}`).join(' ');
          const prevPath = prevValues.map((val, idx) => `${idx === 0 ? 'M' : 'L'} ${getX(idx)} ${getY(val)}`).join(' ');

          return (
            <div className="w-full h-full flex flex-col justify-between">
              <svg className="w-full h-full min-h-[160px]" viewBox={`0 0 ${width} ${height}`}>
                {/* Background Grid */}
                {[0, 0.25, 0.5, 0.75, 1].map((p, idx) => {
                  const yVal = minVal + p * range;
                  const y = getY(yVal);
                  return (
                    <g key={`grid-${idx}`}>
                      <line
                        x1={padding.left}
                        y1={y}
                        x2={width - padding.right}
                        y2={y}
                        stroke={darkTheme ? "#334155" : "#e2e8f0"}
                        strokeDasharray="3,3"
                        strokeWidth="1"
                      />
                      <text
                        x={padding.left - 8}
                        y={y + 3}
                        textAnchor="end"
                        className="text-[9px] font-mono fill-slate-400 font-bold"
                      >
                        {Math.round(yVal).toLocaleString()}
                      </text>
                    </g>
                  );
                })}

                {/* X Axis Labels */}
                {RUSSIAN_MONTHS_SHORT.map((m, idx) => {
                  const x = getX(idx);
                  const isSelected = selectedMonth === (idx + 1);
                  return (
                    <text
                      key={`x-lbl-${idx}`}
                      x={x}
                      y={height - 8}
                      textAnchor="middle"
                      className={`text-[9px] font-bold ${
                        isSelected 
                          ? "fill-blue-500 font-black" 
                          : "fill-slate-400"
                      }`}
                    >
                      {m}
                    </text>
                  );
                })}

                {/* Previous Year Path */}
                <path
                  d={prevPath}
                  fill="none"
                  stroke={darkTheme ? "rgba(148, 163, 184, 0.35)" : "rgba(100, 116, 139, 0.25)"}
                  strokeWidth="2"
                  strokeDasharray="4,4"
                />

                {/* Current Year Path */}
                <path
                  d={currPath}
                  fill="none"
                  stroke="#3b82f6"
                  strokeWidth="3"
                />

                {/* Previous Year Dots */}
                {prevValues.map((val, idx) => (
                  <circle
                    key={`dot-prev-${idx}`}
                    cx={getX(idx)}
                    cy={getY(val)}
                    r="3"
                    fill={darkTheme ? "#1e293b" : "#ffffff"}
                    stroke={darkTheme ? "rgba(148, 163, 184, 0.6)" : "rgba(100, 116, 139, 0.4)"}
                    strokeWidth="1.5"
                  />
                ))}

                {/* Current Year Dots */}
                {currValues.map((val, idx) => {
                  const isSelected = selectedMonth === (idx + 1);
                  return (
                    <circle
                      key={`dot-curr-${idx}`}
                      cx={getX(idx)}
                      cy={getY(val)}
                      r={isSelected ? "5" : "3.5"}
                      fill={isSelected ? "#3b82f6" : (darkTheme ? "#0f172a" : "#ffffff")}
                      stroke="#3b82f6"
                      strokeWidth={isSelected ? "3" : "2"}
                    />
                  );
                })}
              </svg>
              <div className="flex justify-center gap-6 mt-1 text-[10px] font-bold">
                <span className="flex items-center gap-1 text-slate-400">
                  <span className="w-2.5 h-0.5 border-t border-dashed border-slate-400"></span>
                  <span>{selectedYear - 1} г.</span>
                </span>
                <span className="flex items-center gap-1 text-blue-500">
                  <span className="w-2.5 h-0.5 bg-blue-500"></span>
                  <span>{selectedYear} г.</span>
                </span>
              </div>
            </div>
          );
        };

        (globalThis as any).renderCustomSvgTrendChart = renderCustomSvgTrendChart;
        return null;
      })()}

      {anomaliesModalOpen && (
        <div id="anomalies-list-modal" className="fixed inset-0 z-[90] flex items-center justify-center bg-black/70 backdrop-blur-sm p-4 animate-fade-in">
          <div 
            className={`w-full max-w-3xl rounded-2xl border p-6 shadow-2xl flex flex-col max-h-[90vh] ${
              darkTheme ? 'bg-slate-900 border-slate-700 text-slate-100 shadow-slate-950/50' : 'bg-white border-slate-200 text-slate-800'
            }`}
          >
            {/* Header */}
            <div className="flex justify-between items-center border-b border-slate-700/50 pb-4">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-amber-500/10 text-amber-500 rounded-lg">
                  <AlertTriangle className="w-5 h-5 animate-bounce" />
                </div>
                <div>
                  <h3 className="text-base font-black tracking-tight flex items-center gap-2">
                    <span>Критические аномалии энергопотребления</span>
                    <span className="text-xs px-2.5 py-0.5 bg-rose-500/10 text-rose-400 border border-rose-500/20 rounded-full font-bold font-mono">
                      {currentAnomalies.length}
                    </span>
                  </h3>
                  <p className="text-[10px] text-slate-450 font-semibold uppercase tracking-wider mt-0.5">
                    Автоматический аудит за {selectedPeriodName} {selectedYear} г.
                  </p>
                </div>
              </div>
              <button 
                onClick={() => {
                  setAnomaliesModalOpen(false);
                  setAnomalySearchQuery('');
                  setAnomalySeverityFilter('all');
                  setAnomalyTypeFilter('all');
                }} 
                className={`p-1.5 rounded-lg transition-colors ${darkTheme ? 'hover:bg-slate-800 text-slate-400 hover:text-white' : 'hover:bg-slate-100 text-slate-500 hover:text-slate-900'}`}
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {selectedAnomalyForDynamics ? (
              /* --- DETAILED DYNAMICS VIEW FOR SELECTED ANOMALY --- */
              <div className="flex-1 flex flex-col overflow-hidden space-y-4">
                {/* Back button and item title */}
                <div className="flex items-center justify-between border-b border-slate-800/50 pb-3 mt-2">
                  <button
                    type="button"
                    onClick={() => setSelectedAnomalyForDynamics(null)}
                    className="flex items-center gap-1.5 text-xs font-bold text-sky-400 hover:text-sky-300 transition-colors"
                  >
                    <span>← Вернуться к списку аномалий</span>
                  </button>
                  <div className="flex items-center gap-2">
                    <span className="text-[10px] uppercase font-bold tracking-widest text-slate-400">
                      {selectedAnomalyForDynamics.targetType === 'station' ? 'Станция' : selectedAnomalyForDynamics.targetType === 'supply_point' ? 'Точка ТП' : 'Объект потерь'}
                    </span>
                    {selectedAnomalyForDynamics.severity === 'high' ? (
                      <span className="text-[9px] px-2 py-0.5 rounded bg-rose-500/15 text-rose-400 border border-rose-500/30 font-black uppercase">Высокая 🔴</span>
                    ) : selectedAnomalyForDynamics.severity === 'medium' ? (
                      <span className="text-[9px] px-2 py-0.5 rounded bg-amber-500/15 text-amber-400 border border-amber-500/30 font-black uppercase">Средняя 🟡</span>
                    ) : (
                      <span className="text-[9px] px-2 py-0.5 rounded bg-emerald-500/15 text-emerald-450 border border-emerald-500/30 font-black uppercase">Низкая 🟢</span>
                    )}
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-12 gap-5 flex-1 overflow-y-auto pr-1">
                  {/* Left Column: Dynamics chart & metadata */}
                  <div className="md:col-span-7 space-y-4">
                    <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-950/50 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
                      <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-3 flex items-center gap-1.5">
                        <Activity className="w-3.5 h-3.5 text-blue-500" />
                        <span>График динамики электропотребления (кВт·ч)</span>
                      </h4>
                      <div className="h-44 flex items-center justify-center">
                        <CustomSvgTrendChart targetId={selectedAnomalyForDynamics.targetId} type={selectedAnomalyForDynamics.targetType} />
                      </div>
                    </div>

                    <div className={`p-4 rounded-xl border ${darkTheme ? 'bg-slate-950/45 border-slate-800' : 'bg-slate-50 border-slate-200'} space-y-2`}>
                      <div className="text-xs font-extrabold text-blue-500 uppercase tracking-widest">Описание аномалии:</div>
                      <h3 className="text-sm font-extrabold">{selectedAnomalyForDynamics.targetName}</h3>
                      <p className="text-xs text-slate-400 leading-relaxed">{selectedAnomalyForDynamics.description}</p>
                      <div className="pt-2 flex justify-between items-center">
                        <span className="text-[10px] text-slate-500 font-bold uppercase font-mono">Метрика отклонения: {selectedAnomalyForDynamics.metric}</span>
                        <button
                          type="button"
                          onClick={() => {
                            if (selectedAnomalyForDynamics.targetType === 'station') {
                              setSelectedStationId(selectedAnomalyForDynamics.targetId);
                              setActiveTab('stations');
                            } else if (selectedAnomalyForDynamics.targetType === 'supply_point') {
                              const point = supplyPoints.find(p => p.id === selectedAnomalyForDynamics.targetId);
                              if (point) {
                                setSelectedStationId(point.stationId);
                                setSelectedSupplyPointId(point.id);
                                setActiveTab('stations');
                              }
                            } else if (selectedAnomalyForDynamics.targetType === 'loss') {
                              setSelectedLossId(selectedAnomalyForDynamics.targetId);
                              setActiveTab('losses');
                            }
                            setAnomaliesModalOpen(false);
                            setSelectedAnomalyForDynamics(null);
                          }}
                          className="text-xs font-bold text-sky-400 hover:underline flex items-center gap-1"
                        >
                          Перейти к разделу в системе <ExternalLink className="w-3 h-3" />
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Right Column: Interactive indicators editor */}
                  <div className="md:col-span-5 space-y-4">
                    <div className={`p-4 rounded-xl border flex flex-col h-full ${darkTheme ? 'bg-slate-950/45 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
                      <h4 className="text-xs font-bold uppercase tracking-wider text-slate-400 mb-2">
                        ✍️ Корректировка показателей
                      </h4>
                      <p className="text-[10px] text-slate-500 mb-3 font-semibold uppercase">
                        {selectedAnomalyForDynamics.targetType === 'station'
                          ? 'Суммарные показатели станции (изменяются через ТП):'
                          : `Редактирование расхода помесячно за ${selectedYear-1} / ${selectedYear} гг.`}
                      </p>

                      {selectedAnomalyForDynamics.targetType === 'station' ? (
                        <div className="space-y-2.5 flex-1 overflow-y-auto max-h-[300px]">
                          <p className="text-[11px] text-slate-400 italic">Показатели железнодорожной станции складываются из расходов её точек поставки. Отредактируйте подключенные ТП ниже:</p>
                          {supplyPoints.filter(p => p.stationId === selectedAnomalyForDynamics.targetId).map(p => {
                            const pCurr = readings.find(r => r.supplyPointId === p.id && r.year === selectedYear && r.month === selectedMonth)?.value || 0;
                            const pPrev = readings.find(r => r.supplyPointId === p.id && r.year === selectedYear - 1 && r.month === selectedMonth)?.value || 0;
                            return (
                              <div key={p.id} className={`p-2 rounded-lg border flex items-center justify-between gap-2 text-xs ${darkTheme ? 'bg-slate-900/60 border-slate-800' : 'bg-white border-slate-200'}`}>
                                <div className="truncate">
                                  <div className="font-bold truncate" title={p.name}>{p.name}</div>
                                  <div className="text-[9px] text-slate-500 font-semibold font-mono">25г: {pPrev.toLocaleString()} | 26г: {pCurr.toLocaleString()}</div>
                                </div>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setPointModal({
                                      isOpen: true,
                                      mode: 'edit',
                                      id: p.id,
                                      name: p.name,
                                      stationId: p.stationId,
                                      category: p.category,
                                      note: p.note,
                                      isActive: p.isActive,
                                      calculationMethod: p.calculationMethod || 'meter',
                                      currKwh: pCurr,
                                      prevKwh: pPrev
                                    });
                                  }}
                                  className="px-2.5 py-1 bg-sky-500/10 hover:bg-sky-500/20 text-sky-400 rounded text-[10px] font-bold"
                                >
                                  Править ТП
                                </button>
                              </div>
                            );
                          })}
                        </div>
                      ) : (
                        <div className="flex-1 overflow-y-auto space-y-2 pr-1 max-h-[320px]">
                          <table className="w-full text-xs">
                            <thead>
                              <tr className="text-[9px] text-slate-500 uppercase font-bold border-b border-slate-800/60">
                                <th className="text-left pb-1">Месяц</th>
                                <th className="text-right pb-1">{selectedYear - 1} г.</th>
                                <th className="text-right pb-1">{selectedYear} г.</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-800/30">
                              {Array.from({ length: 12 }, (_, i) => i + 1).map(m => {
                                const isCurrentActiveMonth = selectedMonth === m;
                                let valCurr = 0;
                                let valPrev = 0;

                                if (selectedAnomalyForDynamics.targetType === 'supply_point') {
                                  valCurr = readings.find(r => r.supplyPointId === selectedAnomalyForDynamics.targetId && r.year === selectedYear && r.month === m)?.value || 0;
                                  valPrev = readings.find(r => r.supplyPointId === selectedAnomalyForDynamics.targetId && r.year === selectedYear - 1 && r.month === m)?.value || 0;
                                } else {
                                  valCurr = lossReadings.find(r => r.lossObjectId === selectedAnomalyForDynamics.targetId && r.year === selectedYear && r.month === m)?.value || 0;
                                  valPrev = lossReadings.find(r => r.lossObjectId === selectedAnomalyForDynamics.targetId && r.year === selectedYear - 1 && r.month === m)?.value || 0;
                                }

                                return (
                                  <tr key={m} className={`hover:bg-slate-900/10 ${isCurrentActiveMonth ? 'bg-blue-500/10 font-semibold text-blue-300' : ''}`}>
                                    <td className="py-1 text-[11px]">{RUSSIAN_MONTHS[m - 1]}</td>
                                    <td className="py-1 text-right">
                                      <input
                                        type="number"
                                        defaultValue={valPrev || ''}
                                        onBlur={(e) => {
                                          const v = Number(e.target.value);
                                          if (selectedAnomalyForDynamics.targetType === 'supply_point') {
                                            let updatedRds = [...readings];
                                            const idx = updatedRds.findIndex(r => r.supplyPointId === selectedAnomalyForDynamics.targetId && r.year === selectedYear - 1 && r.month === m);
                                            if (idx > -1) {
                                              updatedRds[idx] = { ...updatedRds[idx], value: v };
                                            } else {
                                              updatedRds.push({ supplyPointId: selectedAnomalyForDynamics.targetId, year: selectedYear - 1, month: m, value: v });
                                            }
                                            syncDatabase(stations, supplyPoints, categories, updatedRds, lossObjects, lossReadings);
                                          } else {
                                            let updatedLossRds = [...lossReadings];
                                            const idx = updatedLossRds.findIndex(r => r.lossObjectId === selectedAnomalyForDynamics.targetId && r.year === selectedYear - 1 && r.month === m);
                                            if (idx > -1) {
                                              updatedLossRds[idx] = { ...updatedLossRds[idx], value: v };
                                            } else {
                                              updatedLossRds.push({ lossObjectId: selectedAnomalyForDynamics.targetId, year: selectedYear - 1, month: m, value: v });
                                            }
                                            syncDatabase(stations, supplyPoints, categories, readings, lossObjects, updatedLossRds);
                                          }
                                        }}
                                        className={`w-16 text-right px-1.5 py-0.5 text-[11px] rounded outline-none border font-mono ${darkTheme ? 'bg-slate-950 border-slate-800 text-white focus:border-blue-500' : 'bg-white border-slate-300 focus:border-blue-500'}`}
                                      />
                                    </td>
                                    <td className="py-1 text-right">
                                      <input
                                        type="number"
                                        defaultValue={valCurr || ''}
                                        onBlur={(e) => {
                                          const v = Number(e.target.value);
                                          if (selectedAnomalyForDynamics.targetType === 'supply_point') {
                                            let updatedRds = [...readings];
                                            const idx = updatedRds.findIndex(r => r.supplyPointId === selectedAnomalyForDynamics.targetId && r.year === selectedYear && r.month === m);
                                            if (idx > -1) {
                                              updatedRds[idx] = { ...updatedRds[idx], value: v };
                                            } else {
                                              updatedRds.push({ supplyPointId: selectedAnomalyForDynamics.targetId, year: selectedYear, month: m, value: v });
                                            }
                                            syncDatabase(stations, supplyPoints, categories, updatedRds, lossObjects, lossReadings);
                                          } else {
                                            let updatedLossRds = [...lossReadings];
                                            const idx = updatedLossRds.findIndex(r => r.lossObjectId === selectedAnomalyForDynamics.targetId && r.year === selectedYear && r.month === m);
                                            if (idx > -1) {
                                              updatedLossRds[idx] = { ...updatedLossRds[idx], value: v };
                                            } else {
                                              updatedLossRds.push({ lossObjectId: selectedAnomalyForDynamics.targetId, year: selectedYear, month: m, value: v });
                                            }
                                            syncDatabase(stations, supplyPoints, categories, readings, lossObjects, updatedLossRds);
                                          }
                                        }}
                                        className={`w-16 text-right px-1.5 py-0.5 text-[11px] rounded outline-none border font-mono ${darkTheme ? 'bg-slate-950 border-slate-800 text-white focus:border-blue-500' : 'bg-white border-slate-300 focus:border-blue-500'}`}
                                      />
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            ) : (
              <>
                {/* Controls / Filters */}
                <div className="grid grid-cols-1 md:grid-cols-3 gap-3 my-4">
                  <div className="relative">
                    <Search className="w-4 h-4 text-slate-400 absolute left-3 top-3" />
                    <input
                      type="text"
                      placeholder="Поиск по объекту или описанию..."
                      value={anomalySearchQuery}
                      onChange={(e) => setAnomalySearchQuery(e.target.value)}
                      className={`w-full text-xs pl-9 pr-3 py-2.5 border outline-none rounded-lg font-medium ${
                        darkTheme ? 'bg-slate-950 border-slate-800 text-white placeholder-slate-500' : 'bg-slate-50 border-slate-300 text-slate-800'
                      }`}
                    />
                  </div>

                  <div>
                    <select
                      value={anomalyTypeFilter}
                      onChange={(e) => setAnomalyTypeFilter(e.target.value as any)}
                      className={`w-full text-xs px-3 py-2.5 border outline-none rounded-lg font-bold ${
                        darkTheme ? 'bg-slate-950 border-slate-800 text-slate-300' : 'bg-slate-50 border-slate-300 text-slate-800'
                      }`}
                    >
                      <option value="all">⚡ Все типы объектов</option>
                      <option value="station">🚉 Станции</option>
                      <option value="supply_point">🔌 Точки поставки (ТП)</option>
                      <option value="loss">📉 Технологические потери</option>
                    </select>
                  </div>

                  <div>
                    <select
                      value={anomalySeverityFilter}
                      onChange={(e) => setAnomalySeverityFilter(e.target.value as any)}
                      className={`w-full text-xs px-3 py-2.5 border outline-none rounded-lg font-bold ${
                        darkTheme ? 'bg-slate-950 border-slate-800 text-slate-300' : 'bg-slate-50 border-slate-300 text-slate-800'
                      }`}
                    >
                      <option value="all">🔥 Все уровни критичности</option>
                      <option value="high">🔴 Высокая критичность</option>
                      <option value="medium">🟡 Средняя критичность</option>
                      <option value="low">🟢 Низкая критичность</option>
                    </select>
                  </div>
                </div>

                {/* List */}
                <div className="flex-1 overflow-y-auto min-h-[300px] max-h-[50vh] pr-1 space-y-3">
                  {(() => {
                    const filteredAnoms = currentAnomalies.filter(anom => {
                      if (anomalySearchQuery) {
                        const q = anomalySearchQuery.toLowerCase();
                        const matchName = anom.targetName.toLowerCase().includes(q);
                        const matchDesc = anom.description.toLowerCase().includes(q);
                        const matchMetric = anom.metric.toLowerCase().includes(q);
                        if (!matchName && !matchDesc && !matchMetric) return false;
                      }
                      if (anomalySeverityFilter !== 'all' && anom.severity !== anomalySeverityFilter) return false;
                      if (anomalyTypeFilter !== 'all' && anom.targetType !== anomalyTypeFilter) return false;
                      return true;
                    });

                    if (filteredAnoms.length === 0) {
                      return (
                        <div className="text-center py-12 space-y-2">
                          <div className="text-3xl">🕊️</div>
                          <p className="text-sm font-bold text-slate-400">Аномалий с выбранными фильтрами не обнаружено</p>
                          <p className="text-xs text-slate-500">Попробуйте сбросить фильтры поиска или выбрать другой период.</p>
                        </div>
                      );
                    }

                    return filteredAnoms.map(anom => {
                      const getSeverityBadge = (sev: string) => {
                        if (sev === 'high') {
                          return <span className="text-[8px] px-2 py-0.5 rounded-full font-black uppercase tracking-wider bg-rose-500/10 text-rose-400 border border-rose-500/25">Высокая 🔴</span>;
                        }
                        if (sev === 'medium') {
                          return <span className="text-[8px] px-2 py-0.5 rounded-full font-black uppercase tracking-wider bg-amber-500/10 text-amber-400 border border-amber-500/25">Средняя 🟡</span>;
                        }
                        return <span className="text-[8px] px-2 py-0.5 rounded-full font-black uppercase tracking-wider bg-emerald-500/10 text-emerald-450 border border-emerald-500/25">Низкая 🟢</span>;
                      };

                      const getTargetLabel = (type: string) => {
                        if (type === 'station') return 'Железнодорожная станция';
                        if (type === 'supply_point') return 'Точка ТП';
                        return 'Объект потерь сети';
                      };

                      return (
                        <div 
                          key={anom.id}
                          onClick={() => {
                            setSelectedAnomalyForDynamics(anom);
                          }}
                          className={`p-4 rounded-xl border cursor-pointer transition-all duration-150 active:scale-[0.99] group ${
                            darkTheme 
                              ? 'bg-slate-950/45 border-slate-800 hover:border-blue-500 hover:bg-slate-900/60' 
                              : 'bg-slate-50 border-slate-200 hover:border-blue-400 hover:bg-white'
                          }`}
                        >
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
                            <div className="space-y-1">
                              <div className="flex flex-wrap items-center gap-2">
                                <span className="text-[9px] font-extrabold text-slate-450 uppercase tracking-widest">{getTargetLabel(anom.targetType)}</span>
                                {getSeverityBadge(anom.severity)}
                              </div>
                              <h4 className={`text-sm font-extrabold tracking-tight ${darkTheme ? 'text-white' : 'text-slate-900'} group-hover:text-blue-500 transition-colors`}>
                                {anom.targetName}
                              </h4>
                              <p className="text-xs text-slate-400 leading-relaxed max-w-xl">{anom.description}</p>
                            </div>
                            <div className="shrink-0 flex sm:flex-col items-end gap-1.5 justify-between">
                              <span className="text-xs font-mono font-bold px-3 py-1 bg-red-500/10 text-rose-400 border border-red-500/20 rounded-lg">
                                {anom.metric}
                              </span>
                              <span className="text-[9px] font-bold text-blue-500 group-hover:underline flex items-center gap-0.5">
                                Смотреть динамику / Править →
                              </span>
                            </div>
                          </div>
                        </div>
                      );
                    });
                  })()}
                </div>
              </>
            )}

            {/* Footer */}
            <div className="flex justify-between items-center border-t border-slate-800/50 pt-4 mt-4 text-xs font-semibold text-slate-450">
              <span>Всего в этом месяце выявлено: {currentAnomalies.length} отклонений</span>
              <button 
                onClick={() => {
                  setAnomaliesModalOpen(false);
                  setAnomalySearchQuery('');
                  setAnomalySeverityFilter('all');
                  setAnomalyTypeFilter('all');
                  setSelectedAnomalyForDynamics(null);
                }}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-lg transition-colors"
              >
                Закрыть
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
