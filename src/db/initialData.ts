/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface Station {
  id: string;
  name: string;
  note: string;
  section?: string; // участок контроля
}

export interface SupplyPoint {
  id: string;
  name: string;
  stationId: string;
  category: string;
  note: string;
  isActive: boolean;
  meterNumber?: string;
  calculationMethod?: 'meter' | 'estimated';
}

export interface Reading {
  supplyPointId: string;
  year: number;
  month: number; // 1-12
  value: number; // kWh
}

export interface LossObject {
  id: string;
  name: string;
  note: string;
  stationId?: string; // ассоциированный ИД станции для потерь
  section?: string; // участок контроля
}

export interface LossReading {
  lossObjectId: string;
  year: number;
  month: number;
  value: number; // kWh
}

export const INITIAL_CATEGORIES = [
  "Освещение горловин",
  "Освещение парков",
  "Освещение платформ",
  "Освещение территории",
  "Отопление",
  "Обогрев стрелок",
  "Бытовые нагрузки",
  "Пост ЭЦ",
  "СЦБ и связь",
  "Компрессорная",
  "Насосная",
  "Гараж",
  "Мастерская",
  "Вокзал",
  "Прочие"
];

// Seed list of stations
export const INITIAL_STATIONS: Station[] = [
  { id: "st-alexandrov", name: "Станция Александров-1", section: "Московско-Ярославский диспетчерский участок", note: "Узловой логистический узел, высокая нагрузка освещения парков" },
  { id: "st-bologoe", name: "Станция Бологое-Московское", section: "Бологовский участок Октябрьской ж.д.", note: "Крупная узловая станция Октябрьской дороги. Зимний обогрев стрелок" },
  { id: "st-vladimir", name: "Станция Владимир-Пассажирский", section: "Владимирский участок Горьковской ж.д.", note: "Высокий расход электроэнергии на отопление вокзального комплекса" },
  { id: "st-gryazi", name: "Станция Грязи-Воронежские", section: "Грязинский узел Юго-Восточной ж.д.", note: "Сортировочная станция, крупная компрессорная и ремонтные мастерские" },
  { id: "st-petushki", name: "Станция Петушки", section: "Московско-Курский пригородный сектор", note: "Обслуживание пригородного сообщения, повышенные нагрузки СЦБ" }
];

// Seed list of supply points
export const INITIAL_SUPPLY_POINTS: SupplyPoint[] = [
  // Александров-1
  { id: "tp-alex-1", name: "ТП-1 Фидер №1 (Пост ЭЦ)", stationId: "st-alexandrov", category: "Пост ЭЦ", note: "Питание релейной ЭЦ", isActive: true },
  { id: "tp-alex-2", name: "ТП-1 Фидер №5 (Освещение)", stationId: "st-alexandrov", category: "Освещение парков", note: "Осветительные мачты нечетной горловины", isActive: true },
  { id: "tp-alex-3", name: "ТП-2 Тяговая (Обогрев стрелок)", stationId: "st-alexandrov", category: "Обогрев стрелок", note: "Пневматический и электрообогрев", isActive: true },
  { id: "tp-alex-4", name: "ТП-2 Фидер №6 (Пассажирские платформы)", stationId: "st-alexandrov", category: "Освещение платформ", note: "Платформы 1, 2 и 3 пути", isActive: true },
  { id: "tp-alex-5", name: "ТП-3 Бытовой фидер (Мастерские)", stationId: "st-alexandrov", category: "Мастерская", note: "Инструментальный цех ПЧ-12", isActive: true },

  // Бологое
  { id: "tp-bol-1", name: "ТП-7 (Пост ЭЦ Главный)", stationId: "st-bologoe", category: "Пост ЭЦ", note: "Резервное питание электроснабжения", isActive: true },
  { id: "tp-bol-2", name: "ТП-10 Освещение (Парк А)", stationId: "st-bologoe", category: "Освещение парков", note: "Высотное ригельное освещение", isActive: true },
  { id: "tp-bol-3", name: "ТП-10 Освещение (Парк Б)", stationId: "st-bologoe", category: "Освещение парков", note: "Прожекторные мачты", isActive: true },
  { id: "tp-bol-4", name: "ТП-12 Электрообогрев 1-16", stationId: "st-bologoe", category: "Обогрев стрелок", note: "Релейный шкаф ЭОС", isActive: true },
  { id: "tp-bol-5", name: "ТП-15 Вокзал Главный", stationId: "st-bologoe", category: "Вокзал", note: "Основное здание вокзала, кассы", isActive: true },

  // Владимир
  { id: "tp-vlad-1", name: "ТП-3 Фидер Вокзал (Отопление)", stationId: "st-vladimir", category: "Отопление", note: "Калориферы центрального зала", isActive: true },
  { id: "tp-vlad-2", name: "ТП-3 Пассажирский конкорс", stationId: "st-vladimir", category: "Освещение платформ", note: "Пешеходный мост и платформы", isActive: true },
  { id: "tp-vlad-3", name: "ТП-4 СЦБ и Связь", stationId: "st-vladimir", category: "СЦБ и связь", note: "Связевой шкаф ШЦН", isActive: true },
  { id: "tp-vlad-4", name: "ТП-4 Насосная станция", stationId: "st-vladimir", category: "Насосная", note: "Дренажная и пожарная система", isActive: true },
  { id: "tp-vlad-5", name: "ТП-5 Вагонное депо", stationId: "st-vladimir", category: "Мастерская", note: "Цех технического обслуживания вагонов", isActive: true },

  // Грязи
  { id: "tp-gry-1", name: "ТП-1 Компрессорная Сортировки", stationId: "st-gryazi", category: "Компрессорная", note: "Ресиверы парковых пневмоочисток", isActive: true },
  { id: "tp-gry-2", name: "ТП-1 Горловина парка Грязи-В.", stationId: "st-gryazi", category: "Освещение горловин", note: "Освещение стрелочных переводов", isActive: true },
  { id: "tp-gry-3", name: "ТП-2 Тяговая подстанция (Бытовые)", stationId: "st-gryazi", category: "Бытовые нагрузки", note: "Бытовой корпус ст. Грязи", isActive: true },
  { id: "tp-gry-4", name: "ТП-2 Депо веерное (Гараж)", stationId: "st-gryazi", category: "Гараж", note: "Гараж дрезин и спецтехники ШЧ", isActive: true },
  { id: "tp-gry-5", name: "ТП-3 Отопление бытовок ЭЧ", stationId: "st-gryazi", category: "Отопление", note: "Электрообогреватели цеха службы пути", isActive: true },

  // Петушки
  { id: "tp-pet-1", name: "ТП-Петушки 1 Ф.1 (СЦБ)", stationId: "st-petushki", category: "СЦБ и связь", note: "Питание автоблокировки перегона", isActive: true },
  { id: "tp-pet-2", name: "ТП-Петушки 1 Ф.5 (Платформы)", stationId: "st-petushki", category: "Освещение платформ", note: "Освещение платформы электропоездов", isActive: true },
  { id: "tp-pet-3", name: "ТП-Петушки 2 Вокзальный модуль", stationId: "st-petushki", category: "Вокзал", note: "Зал ожидания, билетные автоматы", isActive: true },
  { id: "tp-pet-4", name: "ТП-Петушки 2 Территория ШЧ", stationId: "st-petushki", category: "Освещение территории", note: "Резервное освещение склада материалов", isActive: true },
  { id: "tp-pet-5", name: "ТП-Петушки 3 Сварочный пост", stationId: "st-petushki", category: "Мастерская", note: "Сварочные трансформаторы ПЧ", isActive: true }
];

// Seed list of standalone loss objects
export const INITIAL_LOSS_OBJECTS: LossObject[] = [
  { id: "loss-trans-1", name: "Потери 1 (Трансформатор Т-1 Александров)", stationId: "st-alexandrov", note: "Потери холостого хода и нагрузочные потери трансформатора Т-1" },
  { id: "loss-line-2", name: "Потери 2 (ЛЭП 10кВ ф.8 Бологое)", stationId: "st-bologoe", note: "Потери в линии электропередачи протяженностью 8.5 км" },
  { id: "loss-cable-3", name: "Потери 3 (Кабельные сети Владимир)", stationId: "st-vladimir", note: "Технологические распределительные потери" },
  { id: "loss-reactive-4", name: "Потери 4 (Реактивная энергия Грязи)", stationId: "st-gryazi", note: "Потери из-за циркуляции реактивного тока" }
];

/**
 * Generate highly realistic, deterministic month-by-month readings for 2025 and 2026.
 * Incorporates seasonal profiles based on category.
 * Added moderate variations to produce interesting "anomalies" such as excessive growth,
 * saving, and historical highs, satisfying the analytical and anomaly detection rules.
 */
export function generateDeterministicDatabase() {
  const readings: Reading[] = [];
  const lossReadings: LossReading[] = [];

  const years = [2025, 2026];
  const months = Array.from({ length: 12 }, (_, i) => i + 1);

  // Seed baseline consumption based on category style
  // (e.g. heating is massive in winter, lighting is high in winter, workshop is flat, etc.)
  function getCategoryFactor(category: string, month: number): number {
    switch (category) {
      case "Отопление":
      case "Обогрев стрелок":
        // Winter peak (Jan, Feb, Dec) and shoulder (Mar, Nov)
        if (month === 12 || month === 1 || month === 2) return 2.8;
        if (month === 3 || month === 11) return 1.5;
        if (month === 4 || month === 10) return 0.6;
        return 0.1; // summer standby or switched off

      case "Освещение горловин":
      case "Освещение парков":
      case "Освещение платформ":
      case "Освещение территории":
        // Daylight length factor (long dark nights in Dec-Jan, short nights in Jun-Jul)
        if (month === 12 || month === 1) return 1.6;
        if (month === 11 || month === 2) return 1.4;
        if (month === 10 || month === 3) return 1.1;
        if (month === 6 || month === 7) return 0.5;
        return 0.9;

      case "СЦБ и связь":
      case "Пост ЭЦ":
        // Constant critical load with 5% ambient fluctuation
        return 1.0 + Math.sin(month) * 0.04;

      case "Компрессорная":
        // Sorting shifts might fluctuate depending on cargo flow
        return 1.0 + Math.sin(month * 1.5) * 0.2;

      case "Вокзал":
      case "Бытовые нагрузки":
        // Peak during summer passenger travel and winter cold
        if (month === 6 || month === 7 || month === 8) return 1.2;
        if (month === 12 || month === 1) return 1.1;
        return 0.9;

      default:
        // Moderately flat
        return 1.0 + Math.cos(month) * 0.1;
    }
  }

  // Generate Readings
  INITIAL_SUPPLY_POINTS.forEach((sp) => {
    // Generate base rate and anomaly factors
    let baseRate = 2000; // default base size
    if (sp.category === "Пост ЭЦ") baseRate = 4500;
    if (sp.category === "Отопление") baseRate = 8000;
    if (sp.category === "Компрессорная") baseRate = 9500;
    if (sp.category === "Бытовые нагрузки") baseRate = 1200;
    if (sp.category === "Мастерская") baseRate = 3000;
    if (sp.category === "Гараж") baseRate = 800;

    years.forEach((year) => {
      months.forEach((month) => {
        // We only have readings up to June 2026. For dates in the future (July 2026 - Dec 2026),
        // we won't seed them by default, mimicking real active months.
        if (year === 2026 && month > 6) {
          return;
        }

        const sizeFactor = getCategoryFactor(sp.category, month);
        
        // Dynamic additions to simulate real-world events and trigger SPECIFIC anomaly conditions:
        // 1. Extreme overrun (>30% growth and >5000 kWh growth) on tp-bol-1 (Post EC chief Bologoe) in June 2026
        // 2. High progress/saving on tp-alex-3 (Bologoe Heating index) in 2026
        // 3. Steady prolonged growth on tp-vlad-1 (heating Vladimir) in April-June 2026
        // 4. Historical maximum trigger on tp-gry-1 (compressor complex)
        
        let growthAnomaly = 1.0;
        
        // We simulate overall background drift (let's say 2026 general usage is slightly higher, e.g. 2% more traffic)
        let drift = year === 2026 ? 1.02 : 1.0;

        if (sp.id === "tp-bol-1" && year === 2026 && month === 6) {
          // Major heating leak / equipment flaw !
          growthAnomaly = 1.45; // 45% spikes!
        }

        if (sp.id === "tp-alex-3" && year === 2026) {
          // Saving program implemented on switch heaters!
          growthAnomaly = 0.77; // 23% saving!
        }

        if (sp.id === "tp-vlad-1" && year === 2026) {
          // Prolonged temperature drop or failure in Spring-Summer 2026 (steady growth across Apr, May, June)
          if (month === 4) growthAnomaly = 1.12;
          if (month === 5) growthAnomaly = 1.25;
          if (month === 6) growthAnomaly = 1.38;
        }

        if (sp.id === "tp-gry-1" && year === 2026) {
          // Cargo volume surge, compressor running 24/7. Record-breaking June!
          if (month === 6) growthAnomaly = 1.55; 
        }

        // Deterministic noise based on ID string code and month for realism
        const sumCharCodes = sp.id.split("").reduce((acc, c) => acc + c.charCodeAt(0), 0);
        const noise = 1.0 + (((sumCharCodes * month * year) % 100) - 50) / 1000; // +/- 5% noise

        const kwhValue = Math.round(baseRate * sizeFactor * drift * growthAnomaly * noise);
        readings.push({
          supplyPointId: sp.id,
          year,
          month,
          value: kwhValue
        });
      });
    });
  });

  // Generate losses readings
  INITIAL_LOSS_OBJECTS.forEach((lo) => {
    let baseLoss = 800; // kWh baseline
    if (lo.id === "loss-trans-1") baseLoss = 1500;
    if (lo.id === "loss-line-2") baseLoss = 2200;
    if (lo.id === "loss-cable-3") baseLoss = 600;
    if (lo.id === "loss-reactive-4") baseLoss = 1800;

    years.forEach((year) => {
      months.forEach((month) => {
        if (year === 2026 && month > 6) return;

        // Losses general behavior tracks load (peaks in winter due to line resistance, heating currents)
        const sizeFactor = 0.8 + 0.4 * (month === 12 || month === 1 || month === 2 ? 1.5 : (month === 6 || month === 7 ? 0.7 : 1.0));
        const drift = year === 2026 ? 1.04 : 1.0; // losses grew in 2026 due to aging grid

        // Create anomaly on loss-line-2 in Jun 2026 (broken insulator leak)
        let lossAnomaly = 1.0;
        if (lo.id === "loss-line-2" && year === 2026 && month === 6) {
          lossAnomaly = 1.35; // 35% growth!
        }

        const sumCodes = lo.id.split("").reduce((acc, c) => acc + c.charCodeAt(0), 0);
        const noise = 1.0 + (((sumCodes * month) % 80) - 40) / 1000; // +/- 4% noise

        const kwhValue = Math.round(baseLoss * sizeFactor * drift * lossAnomaly * noise);
        lossReadings.push({
          lossObjectId: lo.id,
          year,
          month,
          value: kwhValue
        });
      });
    });
  });

  return { readings, lossReadings };
}
