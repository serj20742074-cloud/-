/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { Reading, LossReading, SupplyPoint, Station, LossObject } from './initialData';

export interface Anomaly {
  id: string; // unique anomaly identifier
  targetId: string;
  targetName: string;
  targetType: 'station' | 'supply_point' | 'loss';
  type: 'growth_percent' | 'growth_absolute' | 'prolonged_growth' | 'historical_max';
  severity: 'low' | 'medium' | 'high';
  description: string;
  metric: string;
}

/**
 * Detects anomalies for a single supply point in the selected period (year, month).
 */
export function detectSupplyPointAnomalies(
  sp: SupplyPoint,
  year: number,
  month: number,
  allReadings: Reading[]
): Anomaly[] {
  const anomalies: Anomaly[] = [];
  const currentReading = allReadings.find(r => r.supplyPointId === sp.id && r.year === year && r.month === month);
  if (!currentReading) return [];

  const valCurrent = currentReading.value;

  // 1. Percentage & Absolute growth compared to SAME MONTH PREVIOUS YEAR
  const prevYearReading = allReadings.find(r => r.supplyPointId === sp.id && r.year === year - 1 && r.month === month);
  if (prevYearReading) {
    const valPrev = prevYearReading.value;
    if (valPrev > 0) {
      const diffAbs = valCurrent - valPrev;
      const diffPercent = (diffAbs / valPrev) * 100;

      // Check Percentage Growth: 10% (low), 20% (medium), 30% (high)
      if (diffPercent >= 10) {
        let severity: 'low' | 'medium' | 'high' = 'low';
        if (diffPercent >= 30) severity = 'high';
        else if (diffPercent >= 20) severity = 'medium';

        anomalies.push({
          id: `sp-growth-pct-${sp.id}-${year}-${month}`,
          targetId: sp.id,
          targetName: sp.name,
          targetType: 'supply_point',
          type: 'growth_percent',
          severity,
          description: `Превышение относительно прошлого года на ${diffPercent.toFixed(1)}% (+${diffAbs.toLocaleString()} кВт·ч)`,
          metric: `+${diffPercent.toFixed(1)}%`
        });
      }

      // Check Absolute Growth: 500 (low), 1000 (medium), 5000 (high) kWh
      if (diffAbs >= 500) {
        let severity: 'low' | 'medium' | 'high' = 'low';
        if (diffAbs >= 5000) severity = 'high';
        else if (diffAbs >= 1000) severity = 'medium';

        anomalies.push({
          id: `sp-growth-abs-${sp.id}-${year}-${month}`,
          targetId: sp.id,
          targetName: sp.name,
          targetType: 'supply_point',
          type: 'growth_absolute',
          severity,
          description: `Абсолютный прирост на ${diffAbs.toLocaleString()} кВт·ч относительно аналогичного месяца прошлого года`,
          metric: `${diffAbs.toLocaleString()} кВт·ч`
        });
      }
    }
  }

  // 2. Prolonged Growth: 3 or 6 months consecutively month-on-month
  // E.g. for month 6: r(6) > r(5) > r(4) > r(3) etc.
  const getConsecutiveGrowthMonthsCount = (): number => {
    let count = 0;
    let checkMonth = month;
    let checkYear = year;
    let prevVal = valCurrent;

    while (true) {
      // Step back one month
      checkMonth--;
      if (checkMonth === 0) {
        checkMonth = 12;
        checkYear--;
      }

      const rd = allReadings.find(r => r.supplyPointId === sp.id && r.year === checkYear && r.month === checkMonth);
      if (!rd) break; // no reading, stop checking historical pipeline

      if (prevVal > rd.value) {
        count++;
        prevVal = rd.value;
      } else {
        break;
      }
    }
    return count;
  };

  const growthMonths = getConsecutiveGrowthMonthsCount();
  if (growthMonths >= 6) {
    anomalies.push({
      id: `sp-growth-dur6-${sp.id}-${year}-${month}`,
      targetId: sp.id,
      targetName: sp.name,
      targetType: 'supply_point',
      type: 'prolonged_growth',
      severity: 'high',
      description: `Непрерывный затяжной рост расхода на протяжении 6+ месяцев подряд`,
      metric: `${growthMonths} мес.`
    });
  } else if (growthMonths >= 3) {
    anomalies.push({
      id: `sp-growth-dur3-${sp.id}-${year}-${month}`,
      targetId: sp.id,
      targetName: sp.name,
      targetType: 'supply_point',
      type: 'prolonged_growth',
      severity: 'medium',
      description: `Продолжительное увеличение расхода в течение ${growthMonths} месяцев подряд`,
      metric: `${growthMonths} мес.`
    });
  }

  // 3. Historical Maximum: current exceeds all values prior to current month
  const priorReadings = allReadings.filter(r => 
    r.supplyPointId === sp.id && 
    (r.year < year || (r.year === year && r.month < month))
  );
  if (priorReadings.length > 0) {
    const maxPrior = Math.max(...priorReadings.map(r => r.value));
    if (valCurrent > maxPrior) {
      anomalies.push({
        id: `sp-hist-max-${sp.id}-${year}-${month}`,
        targetId: sp.id,
        targetName: sp.name,
        targetType: 'supply_point',
        type: 'historical_max',
        severity: 'high',
        description: `Превышен исторический максимум за все время наблюдений! Предыдущий пик: ${maxPrior.toLocaleString()} кВт·ч`,
        metric: `Пик ${valCurrent.toLocaleString()}`
      });
    }
  }

  return anomalies;
}

/**
 * Detects anomalies for a loss object in the selected period (year, month).
 */
export function detectLossObjectAnomalies(
  lo: LossObject,
  year: number,
  month: number,
  allLossReadings: LossReading[]
): Anomaly[] {
  const anomalies: Anomaly[] = [];
  const currentReading = allLossReadings.find(r => r.lossObjectId === lo.id && r.year === year && r.month === month);
  if (!currentReading) return [];

  const valCurrent = currentReading.value;

  const prevYearReading = allLossReadings.find(r => r.lossObjectId === lo.id && r.year === year - 1 && r.month === month);
  if (prevYearReading) {
    const valPrev = prevYearReading.value;
    if (valPrev > 0) {
      const diffAbs = valCurrent - valPrev;
      const diffPercent = (diffAbs / valPrev) * 100;

      if (diffPercent >= 10) {
        let severity: 'low' | 'medium' | 'high' = 'low';
        if (diffPercent >= 30) severity = 'high';
        else if (diffPercent >= 20) severity = 'medium';

        anomalies.push({
          id: `lo-growth-pct-${lo.id}-${year}-${month}`,
          targetId: lo.id,
          targetName: lo.name,
          targetType: 'loss',
          type: 'growth_percent',
          severity,
          description: `Увеличение потерь на ${diffPercent.toFixed(1)}% (+${diffAbs.toLocaleString()} кВт·ч) к прошлому году`,
          metric: `+${diffPercent.toFixed(1)}%`
        });
      }

      if (diffAbs >= 500) {
        let severity: 'low' | 'medium' | 'high' = 'low';
        if (diffAbs >= 5000) severity = 'high';
        else if (diffAbs >= 1000) severity = 'medium';

        anomalies.push({
          id: `lo-growth-abs-${lo.id}-${year}-${month}`,
          targetId: lo.id,
          targetName: lo.name,
          targetType: 'loss',
          type: 'growth_absolute',
          severity,
          description: `Технологические потери выросли на ${diffAbs.toLocaleString()} кВт·ч по сравнению с прошлым годом`,
          metric: `${diffAbs.toLocaleString()} кВт·ч`
        });
      }
    }
  }

  // Historical Maximum
  const priorReadings = allLossReadings.filter(r => 
    r.lossObjectId === lo.id && 
    (r.year < year || (r.year === year && r.month < month))
  );
  if (priorReadings.length > 0) {
    const maxPrior = Math.max(...priorReadings.map(r => r.value));
    if (valCurrent > maxPrior) {
      anomalies.push({
        id: `lo-hist-max-${lo.id}-${year}-${month}`,
        targetId: lo.id,
        targetName: lo.name,
        targetType: 'loss',
        type: 'historical_max',
        severity: 'high',
        description: `Выявлен критический пик потерь! Превзойден исторический предел в ${maxPrior.toLocaleString()} кВт·ч`,
        metric: `Пик ${valCurrent.toLocaleString()}`
      });
    }
  }

  return anomalies;
}

/**
 * Detects aggregated anomalies on a full station level (summing all active supply points).
 */
export function detectStationAnomalies(
  station: Station,
  year: number,
  month: number,
  allPoints: SupplyPoint[],
  allReadings: Reading[]
): Anomaly[] {
  const anomalies: Anomaly[] = [];
  const stationPoints = allPoints.filter(p => p.stationId === station.id && p.isActive);
  if (stationPoints.length === 0) return [];

  const getStationSum = (y: number, m: number): number | null => {
    let hasReading = false;
    let sum = 0;
    stationPoints.forEach(p => {
      const rd = allReadings.find(r => r.supplyPointId === p.id && r.year === y && r.month === m);
      if (rd) {
        hasReading = true;
        sum += rd.value;
      }
    });
    return hasReading ? sum : null;
  };

  const valCurrent = getStationSum(year, month);
  if (valCurrent === null) return [];

  const valPrev = getStationSum(year - 1, month);
  if (valPrev !== null && valPrev > 0) {
    const diffAbs = valCurrent - valPrev;
    const diffPercent = (diffAbs / valPrev) * 100;

    if (diffPercent >= 10) {
      let severity: 'low' | 'medium' | 'high' = 'low';
      if (diffPercent >= 30) severity = 'high';
      else if (diffPercent >= 20) severity = 'medium';

      anomalies.push({
        id: `st-growth-pct-${station.id}-${year}-${month}`,
        targetId: station.id,
        targetName: station.name,
        targetType: 'station',
        type: 'growth_percent',
        severity,
        description: `Суммарное потребление станции выросло на ${diffPercent.toFixed(1)}% (+${diffAbs.toLocaleString()} кВт·ч)`,
        metric: `+${diffPercent.toFixed(1)}%`
      });
    }

    if (diffAbs >= 500) {
      let severity: 'low' | 'medium' | 'high' = 'low';
      if (diffAbs >= 5000) severity = 'high';
      else if (diffAbs >= 1000) severity = 'medium';

      anomalies.push({
        id: `st-growth-abs-${station.id}-${year}-${month}`,
        targetId: station.id,
        targetName: station.name,
        targetType: 'station',
        type: 'growth_absolute',
        severity,
        description: `Суммарный перерасход станции: +${diffAbs.toLocaleString()} кВт·ч в сравнении с прошлым годом`,
        metric: `${diffAbs.toLocaleString()} кВт·ч`
      });
    }
  }

  // Prolonged Growth consecutively month-on-month
  const getConsecutiveGrowthMonthsCount = (): number => {
    let count = 0;
    let checkMonth = month;
    let checkYear = year;
    let prevVal = valCurrent;

    while (true) {
      checkMonth--;
      if (checkMonth === 0) {
        checkMonth = 12;
        checkYear--;
      }

      const sum = getStationSum(checkYear, checkMonth);
      if (sum === null) break;

      if (prevVal > sum) {
        count++;
        prevVal = sum;
      } else {
        break;
      }
    }
    return count;
  };

  const growthMonths = getConsecutiveGrowthMonthsCount();
  if (growthMonths >= 3) {
    anomalies.push({
      id: `st-growth-dur-${station.id}-${year}-${month}`,
      targetId: station.id,
      targetName: station.name,
      targetType: 'station',
      type: 'prolonged_growth',
      severity: growthMonths >= 6 ? 'high' : 'medium',
      description: `Устойчивый рост потребления всей станции в течение ${growthMonths} месяцев подряд`,
      metric: `${growthMonths} мес.`
    });
  }

  // Historical Maximum
  let maxPrior = 0;
  let hasPrior = false;
  // Let's sweep prior months
  for (let y = 2024; y <= year; y++) {
    for (let m = 1; m <= 12; m++) {
      if (y === year && m >= month) break;
      const sum = getStationSum(y, m);
      if (sum !== null) {
        hasPrior = true;
        if (sum > maxPrior) maxPrior = sum;
      }
    }
  }

  if (hasPrior && valCurrent > maxPrior) {
    anomalies.push({
      id: `st-hist-max-${station.id}-${year}-${month}`,
      targetId: station.id,
      targetName: station.name,
      targetType: 'station',
      type: 'historical_max',
      severity: 'high',
      description: `Вся станция зафиксировала пиковое потребление в этом месяце. Предыдущий максимум: ${maxPrior.toLocaleString()}`,
      metric: `Пик ${valCurrent.toLocaleString()}`
    });
  }

  return anomalies;
}

/**
 * Aggregates all anomalies in the system for the chosen year/month.
 */
export function getAllAnomalies(
  stations: Station[],
  supplyPoints: SupplyPoint[],
  lossObjects: LossObject[],
  readings: Reading[],
  lossReadings: LossReading[],
  year: number,
  month: number
): Anomaly[] {
  const list: Anomaly[] = [];

  // General check
  stations.forEach(st => {
    list.push(...detectStationAnomalies(st, year, month, supplyPoints, readings));
  });

  supplyPoints.forEach(sp => {
    if (sp.isActive) {
      list.push(...detectSupplyPointAnomalies(sp, year, month, readings));
    }
  });

  lossObjects.forEach(lo => {
    list.push(...detectLossObjectAnomalies(lo, year, month, lossReadings));
  });

  return list;
}
