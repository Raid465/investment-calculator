import { useState, useMemo, useCallback } from 'react';
import {
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  ReferenceLine,
} from 'recharts';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

type Currency = 'USD' | 'JOD';

interface InvestmentInputs {
  initialAmount: number;
  monthlyDeposit: number;
  annualRate: number;
  years: number;
  inflationRate: number;
  adjustForInflation: boolean;
}

interface Scenario {
  id: string;
  name: string;
  color: string;
  annualRate: number;
}

interface YearData {
  year: number;
  invested: number;
  portfolioValue: number;
  earnings: number;
  growthPercent: number;
  realValue?: number;
}

interface GoalSeekInputs {
  targetAmount: number;
  targetYear: number;
  initialAmount: number;
  annualRate: number;
}

const CURRENCY_CONFIG: Record<Currency, { symbol: string; locale: string; code: string }> = {
  USD: { symbol: '$', locale: 'en-US', code: 'USD' },
  JOD: { symbol: 'د.أ', locale: 'ar-JO', code: 'JOD' },
};

const DEFAULT_SCENARIOS: Scenario[] = [
  { id: 'conservative', name: 'متحفظ', color: '#60a5fa', annualRate: 10 },
  { id: 'moderate', name: 'متوسط', color: '#34d399', annualRate: 18 },
  { id: 'optimistic', name: 'متفائل', color: '#fbbf24', annualRate: 35 },
];

function formatCurrency(amount: number, currency: Currency): string {
  const config = CURRENCY_CONFIG[currency];
  return new Intl.NumberFormat(config.locale, {
    style: 'currency',
    currency: config.code,
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(amount);
}

function formatCompact(value: number): string {
  if (value >= 1_000_000) return `${(value / 1_000_000).toFixed(1)}M`;
  if (value >= 1_000) return `${(value / 1_000).toFixed(0)}K`;
  return value.toFixed(0);
}

function calculateYearlyData(
  inputs: InvestmentInputs,
  overrideRate?: number,
  overrideInitial?: number,
  overrideMonthly?: number
): YearData[] {
  const { annualRate, years, inflationRate, adjustForInflation } = inputs;
  const initialAmount = overrideInitial ?? inputs.initialAmount;
  const monthlyDeposit = overrideMonthly ?? inputs.monthlyDeposit;
  const rate = overrideRate ?? annualRate;
  const monthlyRate = rate / 100 / 12;
  const data: YearData[] = [];
  let balance = initialAmount;

  for (let year = 1; year <= years; year++) {
    for (let month = 0; month < 12; month++) {
      balance = balance * (1 + monthlyRate) + monthlyDeposit;
    }
    const totalInvested = initialAmount + monthlyDeposit * 12 * year;
    const earnings = balance - totalInvested;
    const growthPercent = totalInvested > 0 ? (earnings / totalInvested) * 100 : 0;
    const realValue = adjustForInflation
      ? balance / Math.pow(1 + inflationRate / 100, year)
      : undefined;

    data.push({
      year,
      invested: Math.round(totalInvested),
      portfolioValue: Math.round(balance),
      earnings: Math.round(earnings),
      growthPercent: Math.round(growthPercent * 100) / 100,
      realValue: realValue ? Math.round(realValue) : undefined,
    });
  }

  return data;
}

function calculateGoalSeek(inputs: GoalSeekInputs): {
  monthlyPayment: number;
  isAlreadyReached: boolean;
  futureValueOfInitial: number;
} {
  const { targetAmount, targetYear, initialAmount, annualRate } = inputs;
  const monthlyRate = annualRate / 100 / 12;
  const months = targetYear * 12;
  const futureValueOfInitial = initialAmount * Math.pow(1 + monthlyRate, months);

  if (futureValueOfInitial >= targetAmount) {
    return { monthlyPayment: 0, isAlreadyReached: true, futureValueOfInitial };
  }

  if (monthlyRate === 0) {
    const payment = (targetAmount - initialAmount) / months;
    return { monthlyPayment: Math.max(0, payment), isAlreadyReached: false, futureValueOfInitial };
  }

  const remaining = targetAmount - futureValueOfInitial;
  const annuityFactor = (Math.pow(1 + monthlyRate, months) - 1) / monthlyRate;
  const payment = remaining / annuityFactor;

  return { monthlyPayment: Math.max(0, payment), isAlreadyReached: false, futureValueOfInitial };
}

function calculateGoalProjection(inputs: GoalSeekInputs, monthlyPayment: number): YearData[] {
  const { targetYear, initialAmount, annualRate } = inputs;
  const monthlyRate = annualRate / 100 / 12;
  const data: YearData[] = [];
  let balance = initialAmount;

  for (let year = 1; year <= targetYear; year++) {
    for (let month = 0; month < 12; month++) {
      balance = balance * (1 + monthlyRate) + monthlyPayment;
    }
    const totalInvested = initialAmount + monthlyPayment * 12 * year;
    const earnings = balance - totalInvested;
    const growthPercent = totalInvested > 0 ? (earnings / totalInvested) * 100 : 0;

    data.push({
      year,
      invested: Math.round(totalInvested),
      portfolioValue: Math.round(balance),
      earnings: Math.round(earnings),
      growthPercent: Math.round(growthPercent * 100) / 100,
    });
  }

  return data;
}

async function exportToExcel(
  data: YearData[],
  currency: Currency,
  inputs: InvestmentInputs,
  scenarios?: Scenario[]
): Promise<void> {
  const config = CURRENCY_CONFIG[currency];
  const sym = config.symbol;
  const wb = new ExcelJS.Workbook();
  wb.creator = 'حاسبة الاستثمار';
  wb.created = new Date();

  // ─── Colors ───
  const DARK_BG = '0F172A';
  const CARD_BG = '1E293B';
  const HEADER_BG = '10B981';
  const WHITE = 'FFFFFF';
  const GRAY = '94A3B8';
  const BLUE = '3B82F6';
  const GREEN = '10B981';
  const AMBER = 'F59E0B';
  const PURPLE = '8B5CF6';

  const addSheet = (name: string, title: string, rows: YearData[]) => {
    const ws = wb.addWorksheet(name, {
      properties: { tabColor: { argb: GREEN } },
    });

    // Column widths
    ws.columns = [
      { width: 10 },
      { width: 22 },
      { width: 22 },
      { width: 22 },
      { width: 15 },
      { width: 22 },
    ];

    // Title row
    ws.mergeCells('A1:F1');
    const titleCell = ws.getCell('A1');
    titleCell.value = title;
    titleCell.font = { name: 'Arial', size: 18, bold: true, color: { argb: WHITE } };
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_BG } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(1).height = 45;

    // Info row
    ws.mergeCells('A2:F2');
    const infoCell = ws.getCell('A2');
    infoCell.value = `رأس المال: ${sym}${inputs.initialAmount.toLocaleString()}  |  الإيداع الشهري: ${sym}${inputs.monthlyDeposit.toLocaleString()}  |  العائد: ${inputs.annualRate}%  |  المدة: ${inputs.years} سنة`;
    infoCell.font = { name: 'Arial', size: 11, color: { argb: GRAY } };
    infoCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARK_BG } };
    infoCell.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(2).height = 30;

    // Empty row
    ws.getRow(3).height = 8;

    // Header row
    const headers = ['السنة', `المستثمر (${sym})`, `قيمة المحفظة (${sym})`, `الأرباح (${sym})`, 'النمو %', inputs.adjustForInflation ? `القيمة الحقيقية (${sym})` : ''];
    const headerRow = ws.addRow(headers);
    headerRow.height = 32;
    headerRow.eachCell((cell) => {
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: WHITE } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BG } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        bottom: { style: 'medium', color: { argb: '059669' } },
      };
    });

    // Data rows
    rows.forEach((d, i) => {
      const rowData = [
        d.year,
        d.invested,
        d.portfolioValue,
        d.earnings,
        d.growthPercent / 100,
        d.realValue || '',
      ];
      const row = ws.addRow(rowData);
      row.height = 26;
      const isEven = i % 2 === 0;
      const bgColor = isEven ? CARD_BG : DARK_BG;

      row.eachCell((cell, colNumber) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { name: 'Arial', size: 10, color: { argb: WHITE } };
        cell.border = {
          bottom: { style: 'thin', color: { argb: '334155' } },
        };

        if (colNumber === 1) {
          cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: WHITE } };
        } else if (colNumber === 2) {
          cell.font = { name: 'Arial', size: 10, color: { argb: BLUE } };
          cell.numFmt = `#,##0 "${sym}"`;
        } else if (colNumber === 3) {
          cell.font = { name: 'Arial', size: 10, color: { argb: GREEN } };
          cell.numFmt = `#,##0 "${sym}"`;
        } else if (colNumber === 4) {
          cell.font = { name: 'Arial', size: 10, color: { argb: AMBER } };
          cell.numFmt = `#,##0 "${sym}"`;
        } else if (colNumber === 5) {
          cell.font = { name: 'Arial', size: 10, color: { argb: PURPLE } };
          cell.numFmt = '0.00%';
        } else if (colNumber === 6) {
          cell.font = { name: 'Arial', size: 10, color: { argb: 'FB7185' } };
          cell.numFmt = `#,##0 "${sym}"`;
        }
      });
    });

    // Summary row
    const last = rows[rows.length - 1];
    if (last) {
      ws.addRow([]);
      const summaryRow = ws.addRow(['الإجمالي', last.invested, last.portfolioValue, last.earnings, last.growthPercent / 100, '']);
      summaryRow.height = 32;
      summaryRow.eachCell((cell, colNumber) => {
        cell.font = { name: 'Arial', size: 12, bold: true, color: { argb: WHITE } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0D4F4F' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
          top: { style: 'medium', color: { argb: GREEN } },
          bottom: { style: 'medium', color: { argb: GREEN } },
        };
        if (colNumber >= 2 && colNumber <= 4) {
          cell.numFmt = `#,##0 "${sym}"`;
        } else if (colNumber === 5) {
          cell.numFmt = '0.00%';
        }
      });
    }

    // Add chart
    const chartStartRow = rows.length + 7;
    const imageId = wb.addImage({
      buffer: createChartImage(rows),
      extension: 'png',
    });
    ws.addImage(imageId, {
      tl: { col: 0, row: chartStartRow },
      ext: { width: 850, height: 330 },
    });

    // Freeze header
    ws.views = [{ state: 'frozen', ySplit: 4 }];
  };

  addSheet('الحاسبة', `تقرير الاستثمار — ${currency}`, data);

  if (scenarios) {
    scenarios.forEach((s) => {
      const scenarioData = calculateYearlyData(inputs, s.annualRate);
      addSheet(s.name, `${s.name} — عائد ${s.annualRate}%`, scenarioData);
    });
  }

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  saveAs(blob, `investment-report-${currency}.xlsx`);
}

function createChartImage(data: YearData[]): ArrayBuffer {
  const canvas = document.createElement('canvas');
  canvas.width = 900;
  canvas.height = 350;
  const ctx = canvas.getContext('2d')!;

  ctx.fillStyle = '#0F172A';
  ctx.fillRect(0, 0, 900, 350);

  const padding = { top: 40, right: 30, bottom: 50, left: 80 };
  const chartW = 900 - padding.left - padding.right;
  const chartH = 350 - padding.top - padding.bottom;

  const maxVal = Math.max(...data.map((d) => d.portfolioValue));
  const yMax = Math.ceil(maxVal / 10000) * 10000 || 10000;

  const xScale = (i: number) => padding.left + (i / Math.max(data.length - 1, 1)) * chartW;
  const yScale = (v: number) => padding.top + chartH - (v / yMax) * chartH;

  // Grid lines
  ctx.strokeStyle = '#1E293B';
  ctx.lineWidth = 1;
  for (let i = 0; i <= 5; i++) {
    const y = padding.top + (i / 5) * chartH;
    ctx.beginPath();
    ctx.moveTo(padding.left, y);
    ctx.lineTo(900 - padding.right, y);
    ctx.stroke();

    const val = yMax - (i / 5) * yMax;
    ctx.fillStyle = '#64748B';
    ctx.font = '11px Arial';
    ctx.textAlign = 'right';
    ctx.fillText(`${(val / 1000).toFixed(0)}K`, padding.left - 10, y + 4);
  }

  // X-axis labels
  ctx.fillStyle = '#64748B';
  ctx.font = '11px Arial';
  ctx.textAlign = 'center';
  data.forEach((d, i) => {
    if (i % Math.max(1, Math.floor(data.length / 10)) === 0 || i === data.length - 1) {
      ctx.fillText(`Y${d.year}`, xScale(i), 350 - padding.bottom + 20);
    }
  });

  // Stacked area - Invested
  ctx.beginPath();
  ctx.moveTo(xScale(0), yScale(0));
  data.forEach((d, i) => ctx.lineTo(xScale(i), yScale(d.invested)));
  ctx.lineTo(xScale(data.length - 1), yScale(0));
  ctx.closePath();
  ctx.fillStyle = 'rgba(59, 130, 246, 0.25)';
  ctx.fill();

  // Line - Invested
  ctx.beginPath();
  data.forEach((d, i) => {
    if (i === 0) ctx.moveTo(xScale(i), yScale(d.invested));
    else ctx.lineTo(xScale(i), yScale(d.invested));
  });
  ctx.strokeStyle = '#3B82F6';
  ctx.lineWidth = 2;
  ctx.stroke();

  // Stacked area - Earnings
  ctx.beginPath();
  ctx.moveTo(xScale(0), yScale(data[0].invested));
  data.forEach((d, i) => ctx.lineTo(xScale(i), yScale(d.portfolioValue)));
  for (let i = data.length - 1; i >= 0; i--) {
    ctx.lineTo(xScale(i), yScale(data[i].invested));
  }
  ctx.closePath();
  ctx.fillStyle = 'rgba(16, 185, 129, 0.25)';
  ctx.fill();

  // Line - Total
  ctx.beginPath();
  data.forEach((d, i) => {
    if (i === 0) ctx.moveTo(xScale(i), yScale(d.portfolioValue));
    else ctx.lineTo(xScale(i), yScale(d.portfolioValue));
  });
  ctx.strokeStyle = '#10B981';
  ctx.lineWidth = 2.5;
  ctx.stroke();

  // Legend
  const legendY = 20;
  ctx.fillStyle = '#3B82F6';
  ctx.fillRect(padding.left, legendY, 14, 10);
  ctx.fillStyle = '#CBD5E1';
  ctx.font = '12px Arial';
  ctx.textAlign = 'left';
  ctx.fillText('المستثمر', padding.left + 20, legendY + 9);

  ctx.fillStyle = '#10B981';
  ctx.fillRect(padding.left + 100, legendY, 14, 10);
  ctx.fillStyle = '#CBD5E1';
  ctx.fillText('الإجمالي', padding.left + 120, legendY + 9);

  const base64 = canvas.toDataURL('image/png').split(';base64,')[1];
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function InputField({
  label,
  value,
  onChange,
  min,
  max,
  step,
  suffix,
}: {
  label: string;
  value: number;
  onChange: (v: number) => void;
  min: number;
  max: number;
  step: number;
  suffix?: string;
}) {
  const [raw, setRaw] = useState<string>(String(value));
  const [focused, setFocused] = useState(false);

  if (!focused && raw !== String(value)) {
    setRaw(String(value));
  }

  return (
    <div className="space-y-1.5">
      <label className="block text-sm font-medium text-slate-300">
        {label}
        {suffix && <span className="text-slate-500 mr-1">({suffix})</span>}
      </label>
      <div className="flex items-center gap-3">
        <input
          type="range"
          min={min}
          max={max}
          step={step}
          value={value}
          onChange={(e) => {
            const v = Number(e.target.value);
            onChange(v);
            if (!focused) setRaw(String(v));
          }}
          className="flex-1 h-2 rounded-full appearance-none cursor-pointer bg-slate-700 accent-emerald-500"
        />
        <input
          type="number"
          value={raw}
          onFocus={() => {
            setFocused(true);
            setRaw(String(value));
          }}
          onChange={(e) => {
            const text = e.target.value;
            setRaw(text);
            const v = Number(text);
            if (!isNaN(v) && text !== '' && text !== '-') {
              onChange(Math.min(max, Math.max(min, v)));
            }
          }}
          onBlur={() => {
            setFocused(false);
            let v = Number(raw);
            if (isNaN(v) || raw.trim() === '') v = min;
            v = Math.min(max, Math.max(min, v));
            onChange(v);
            setRaw(String(v));
          }}
          min={min}
          max={max}
          step={step}
          className="w-24 bg-slate-800 border border-slate-600 rounded-lg px-3 py-1.5 text-sm text-right text-white focus:outline-none focus:border-emerald-500 transition"
        />
      </div>
    </div>
  );
}

function MetricCard({
  label,
  value,
  subValue,
  color = 'text-white',
  icon,
}: {
  label: string;
  value: string;
  subValue?: string;
  color?: string;
  icon: React.ReactNode;
}) {
  return (
    <div className="bg-slate-800/60 backdrop-blur border border-slate-700 rounded-xl p-4 space-y-2 hover:border-slate-600 transition">
      <div className="flex items-center justify-between">
        <p className="text-xs text-slate-400 uppercase tracking-wider">{label}</p>
        <div className="text-slate-600">{icon}</div>
      </div>
      <p className={`text-2xl font-bold ${color}`}>{value}</p>
      {subValue && <p className="text-xs text-slate-500">{subValue}</p>}
    </div>
  );
}

function ChartTooltip({ active, payload, label, currency }: any) {
  if (!active || !payload?.length) return null;
  const total = payload.reduce((sum: number, p: any) => sum + (p.value || 0), 0);
  return (
    <div className="bg-slate-900/95 backdrop-blur border border-slate-600 rounded-xl p-4 shadow-2xl text-sm min-w-[200px]">
      <p className="text-slate-400 mb-3 font-medium border-b border-slate-700 pb-2">
        📅 السنة {label}
      </p>
      {payload.map((entry: any, i: number) => {
        const pct = total > 0 ? ((entry.value / total) * 100).toFixed(1) : '0';
        return (
          <p key={i} className="flex justify-between gap-6 py-1">
            <span style={{ color: entry.color }} className="flex items-center gap-2">
              <span className="w-2 h-2 rounded-full" style={{ background: entry.color }} />
              {entry.name}
            </span>
            <span className="font-mono font-bold text-white">
              {formatCurrency(entry.value, currency)}
              <span className="text-xs text-slate-500 ml-1">({pct}%)</span>
            </span>
          </p>
        );
      })}
      <div className="border-t border-slate-700 mt-2 pt-2 flex justify-between">
        <span className="text-slate-400">الإجمالي</span>
        <span className="font-mono font-bold text-emerald-400">{formatCurrency(total, currency)}</span>
      </div>
    </div>
  );
}

function GoalTooltip({ active, payload, label, currency }: any) {
  if (!active || !payload?.length) return null;
  return (
    <div className="bg-slate-900/95 backdrop-blur border border-slate-600 rounded-xl p-4 shadow-2xl text-sm min-w-[180px]">
      <p className="text-slate-400 mb-2 font-medium">السنة {label}</p>
      {payload.map((entry: any, i: number) => (
        <p key={i} className="flex justify-between gap-4 py-0.5">
          <span style={{ color: entry.color }}>{entry.name}</span>
          <span className="font-mono font-bold text-white">{formatCurrency(entry.value, currency)}</span>
        </p>
      ))}
    </div>
  );
}

export default function App() {
  const [currency, setCurrency] = useState<Currency>('USD');
  const [activeTab, setActiveTab] = useState<'calculator' | 'scenarios' | 'goal'>('calculator');
  const [inputs, setInputs] = useState<InvestmentInputs>({
    initialAmount: 10000,
    monthlyDeposit: 500,
    annualRate: 10,
    years: 20,
    inflationRate: 3,
    adjustForInflation: false,
  });
  const [scenarios, setScenarios] = useState<Scenario[]>(DEFAULT_SCENARIOS);
  const [goalInputs, setGoalInputs] = useState<GoalSeekInputs>({
    targetAmount: 500000,
    targetYear: 15,
    initialAmount: 10000,
    annualRate: 10,
  });

  const updateInput = useCallback(
    <K extends keyof InvestmentInputs>(key: K, value: InvestmentInputs[K]) =>
      setInputs((prev) => ({ ...prev, [key]: value })),
    []
  );

  const yearData = useMemo(() => calculateYearlyData(inputs), [inputs]);

  const scenarioData = useMemo(
    () => scenarios.map((s) => ({ ...s, data: calculateYearlyData(inputs, s.annualRate) })),
    [inputs, scenarios]
  );

  const chartData = useMemo(
    () =>
      yearData.map((d) => ({
        year: d.year,
        'المبلغ المستثمر': d.invested,
        الأرباح: d.earnings,
        'المحفظة الكلية': d.portfolioValue,
      })),
    [yearData]
  );

  const scenarioChartData = useMemo(() => {
    return Array.from({ length: inputs.years }, (_, i) => {
      const point: Record<string, number> = { year: i + 1 };
      scenarioData.forEach((s) => {
        const y = s.data[i];
        if (y) point[s.name] = y.portfolioValue;
      });
      return point;
    });
  }, [scenarioData, inputs.years]);

  const goalResult = useMemo(() => calculateGoalSeek(goalInputs), [goalInputs]);
  const goalProjection = useMemo(
    () => calculateGoalProjection(goalInputs, goalResult.monthlyPayment),
    [goalInputs, goalResult.monthlyPayment]
  );

  const goalChartData = useMemo(
    () =>
      goalProjection.map((d) => ({
        year: d.year,
        'المبلغ المستثمر': d.invested,
        الأرباح: d.earnings,
        'المحفظة الكلية': d.portfolioValue,
      })),
    [goalProjection]
  );

  const lastYear = yearData[yearData.length - 1];

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-950 via-slate-900 to-slate-950">
      {/* Header */}
      <header className="border-b border-slate-800 bg-slate-950/50 backdrop-blur sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-emerald-500/20 rounded-xl flex items-center justify-center">
              <svg className="w-6 h-6 text-emerald-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6" />
              </svg>
            </div>
            <div>
              <h1 className="text-lg font-bold text-white">حاسبة الاستثمار</h1>
              <p className="text-xs text-slate-500">Compound Interest Calculator</p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            {(['USD', 'JOD'] as const).map((c) => (
              <button
                key={c}
                onClick={() => setCurrency(c)}
                className={`px-3 py-1.5 rounded-lg text-sm font-medium transition ${
                  currency === c
                    ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30'
                    : 'text-slate-400 hover:text-white'
                }`}
              >
                {c === 'USD' ? '$ دولار' : 'د.أ دينار'}
              </button>
            ))}
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 space-y-8">
        {/* Tabs */}
        <div className="flex gap-1 bg-slate-800/50 p-1 rounded-xl w-fit mx-auto">
          {([
            { key: 'calculator', label: 'الحاسبة', icon: '🧮' },
            { key: 'scenarios', label: 'مقارنة السيناريوهات', icon: '📊' },
            { key: 'goal', label: 'الهدف العكسي', icon: '🎯' },
          ] as const).map((tab) => (
            <button
              key={tab.key}
              onClick={() => setActiveTab(tab.key)}
              className={`px-5 py-2.5 rounded-lg text-sm font-medium transition flex items-center gap-2 ${
                activeTab === tab.key
                  ? 'bg-emerald-500 text-white shadow-lg shadow-emerald-500/25'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <span className="hidden sm:inline">{tab.icon}</span>
              {tab.label}
            </button>
          ))}
        </div>

        {/* ═══ Calculator Tab ═══ */}
        {activeTab === 'calculator' && (
          <div className="grid lg:grid-cols-3 gap-6">
            {/* Inputs */}
            <div className="bg-slate-800/40 backdrop-blur border border-slate-700 rounded-2xl p-6 space-y-5">
              <h2 className="text-white font-bold text-lg flex items-center gap-2">
                <span className="w-2 h-2 bg-emerald-400 rounded-full" />
                معطيات الاستثمار
              </h2>
              <InputField label="رأس المال الابتدائي" value={inputs.initialAmount} onChange={(v) => updateInput('initialAmount', v)} min={0} max={1000000} step={1000} suffix={currency === 'USD' ? '$' : 'د.أ'} />
              <InputField label="الإيداع الشهري" value={inputs.monthlyDeposit} onChange={(v) => updateInput('monthlyDeposit', v)} min={0} max={50000} step={50} suffix={currency === 'USD' ? '$/شهر' : 'د.أ/شهر'} />
              <InputField label="معدل العائد السنوي" value={inputs.annualRate} onChange={(v) => updateInput('annualRate', v)} min={0} max={50} step={0.5} suffix="%" />
              <InputField label="عدد السنوات" value={inputs.years} onChange={(v) => updateInput('years', v)} min={1} max={50} step={1} suffix="سنة" />

              <div className="border-t border-slate-700 pt-4 space-y-3">
                <label className="flex items-center gap-3 cursor-pointer group">
                  <input type="checkbox" checked={inputs.adjustForInflation} onChange={(e) => updateInput('adjustForInflation', e.target.checked)} className="w-4 h-4 rounded accent-emerald-500" />
                  <span className="text-sm text-slate-300 group-hover:text-white transition">تعديل حسب التضخم</span>
                </label>
                {inputs.adjustForInflation && (
                  <InputField label="معدل التضخم السنوي" value={inputs.inflationRate} onChange={(v) => updateInput('inflationRate', v)} min={0} max={30} step={0.5} suffix="%" />
                )}
              </div>

              <button onClick={() => exportToExcel(yearData, currency, inputs, scenarios)} className="w-full py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-sm font-bold transition flex items-center justify-center gap-2">
                <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
                تصدير Excel
              </button>
            </div>

            {/* Results */}
            <div className="lg:col-span-2 space-y-6">
              {lastYear && (
                <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                  <MetricCard
                    label="القيمة النهائية"
                    value={formatCurrency(lastYear.portfolioValue, currency)}
                    color="text-emerald-400"
                    icon={<svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1" /></svg>}
                  />
                  <MetricCard
                    label="إجمالي المستثمر"
                    value={formatCurrency(lastYear.invested, currency)}
                    color="text-blue-400"
                    icon={<svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 9V7a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2m2 4h10a2 2 0 002-2v-6a2 2 0 00-2-2H9a2 2 0 00-2 2v6a2 2 0 002 2zm7-5a2 2 0 11-4 0 2 2 0 014 0z" /></svg>}
                  />
                  <MetricCard
                    label="إجمالي الأرباح"
                    value={formatCurrency(lastYear.earnings, currency)}
                    color="text-amber-400"
                    icon={<svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6" /></svg>}
                  />
                  <MetricCard
                    label="نسبة النمو"
                    value={`${lastYear.growthPercent}%`}
                    subValue={lastYear.realValue ? `القيمة الحقيقية: ${formatCurrency(lastYear.realValue, currency)}` : undefined}
                    color="text-purple-400"
                    icon={<svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" /></svg>}
                  />
                </div>
              )}

              {/* Chart */}
              <div className="bg-slate-800/40 backdrop-blur border border-slate-700 rounded-2xl p-6">
                <div className="flex items-center justify-between mb-4">
                  <h3 className="text-white font-bold">نمو المحفظة</h3>
                  <div className="flex items-center gap-4 text-xs">
                    <span className="flex items-center gap-1.5"><span className="w-3 h-1.5 rounded-full bg-blue-500" /> مستثمر</span>
                    <span className="flex items-center gap-1.5"><span className="w-3 h-1.5 rounded-full bg-emerald-500" /> أرباح</span>
                    <span className="flex items-center gap-1.5"><span className="w-3 h-1.5 rounded-full bg-white/30" /> الإجمالي</span>
                  </div>
                </div>
                <ResponsiveContainer width="100%" height={350}>
                  <AreaChart data={chartData}>
                    <defs>
                      <linearGradient id="gInvested" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#3b82f6" stopOpacity={0.4} />
                        <stop offset="100%" stopColor="#3b82f6" stopOpacity={0.05} />
                      </linearGradient>
                      <linearGradient id="gEarnings" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#10b981" stopOpacity={0.4} />
                        <stop offset="100%" stopColor="#10b981" stopOpacity={0.05} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                    <XAxis dataKey="year" stroke="#475569" tick={{ fill: '#94a3b8', fontSize: 12 }} tickLine={false} />
                    <YAxis stroke="#475569" tick={{ fill: '#94a3b8', fontSize: 12 }} tickFormatter={formatCompact} tickLine={false} width={60} />
                    <Tooltip content={<ChartTooltip currency={currency} />} />
                    <Area type="monotone" dataKey="المبلغ المستثمر" stackId="1" stroke="#3b82f6" fill="url(#gInvested)" strokeWidth={2} animationDuration={800} />
                    <Area type="monotone" dataKey="الأرباح" stackId="1" stroke="#10b981" fill="url(#gEarnings)" strokeWidth={2} animationDuration={800} />
                    <Area type="monotone" dataKey="المحفظة الكلية" fill="none" stroke="rgba(255,255,255,0.15)" strokeWidth={1} strokeDasharray="6 4" dot={false} animationDuration={800} />
                  </AreaChart>
                </ResponsiveContainer>
              </div>

              {/* Table */}
              <div className="bg-slate-800/40 backdrop-blur border border-slate-700 rounded-2xl p-6 overflow-x-auto">
                <h3 className="text-white font-bold mb-4">جدول تفصيلي سنة بسنة</h3>
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b border-slate-700">
                      {['السنة', 'المستثمر التراكمي', 'قيمة المحفظة', 'الأرباح', 'النمو', ...(inputs.adjustForInflation ? ['القيمة الحقيقية'] : [])].map((h) => (
                        <th key={h} className="py-3 px-2 text-slate-400 font-medium text-right">{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {yearData.map((d) => (
                      <tr key={d.year} className="border-b border-slate-700/50 hover:bg-slate-700/20 transition">
                        <td className="py-2.5 px-2 text-white font-mono">{d.year}</td>
                        <td className="py-2.5 px-2 text-blue-400 font-mono text-right">{formatCurrency(d.invested, currency)}</td>
                        <td className="py-2.5 px-2 text-emerald-400 font-mono text-right">{formatCurrency(d.portfolioValue, currency)}</td>
                        <td className="py-2.5 px-2 text-amber-400 font-mono text-right">{formatCurrency(d.earnings, currency)}</td>
                        <td className="py-2.5 px-2 text-purple-400 font-mono text-right">{d.growthPercent}%</td>
                        {inputs.adjustForInflation && (
                          <td className="py-2.5 px-2 text-rose-400 font-mono text-right">{d.realValue ? formatCurrency(d.realValue, currency) : '—'}</td>
                        )}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ═══ Scenarios Tab ═══ */}
        {activeTab === 'scenarios' && (
          <div className="space-y-6">
            {/* Shared Inputs */}
            <div className="bg-slate-800/40 backdrop-blur border border-slate-700 rounded-2xl p-5">
              <h3 className="text-white font-bold mb-4 flex items-center gap-2">
                <span className="w-2 h-2 bg-blue-400 rounded-full" />
                المعطيات المشتركة لجميع السيناريوهات
              </h3>
              <div className="grid md:grid-cols-3 gap-4">
                <InputField label="رأس المال الابتدائي" value={inputs.initialAmount} onChange={(v) => updateInput('initialAmount', v)} min={0} max={1000000} step={1000} suffix={currency === 'USD' ? '$' : 'د.أ'} />
                <InputField label="الإيداع الشهري" value={inputs.monthlyDeposit} onChange={(v) => updateInput('monthlyDeposit', v)} min={0} max={50000} step={50} suffix={currency === 'USD' ? '$/شهر' : 'د.أ/شهر'} />
                <InputField label="عدد السنوات" value={inputs.years} onChange={(v) => updateInput('years', v)} min={1} max={50} step={1} suffix="سنة" />
              </div>
            </div>

            <div className="grid md:grid-cols-3 gap-4">
              {scenarios.map((scenario) => {
                const data = calculateYearlyData(inputs, scenario.annualRate);
                const last = data[data.length - 1];
                return (
                  <div key={scenario.id} className="bg-slate-800/40 backdrop-blur border border-slate-700 rounded-2xl p-5 space-y-3 hover:border-slate-600 transition">
                    <div className="flex items-center gap-2">
                      <div className="w-3 h-3 rounded-full" style={{ backgroundColor: scenario.color }} />
                      <input
                        type="text"
                        value={scenario.name}
                        onChange={(e) => setScenarios((prev) => prev.map((s) => s.id === scenario.id ? { ...s, name: e.target.value } : s))}
                        className="bg-transparent text-white font-bold text-lg border-none outline-none w-full"
                      />
                    </div>
                    <InputField label="معدل العائد السنوي" value={scenario.annualRate} onChange={(v) => setScenarios((prev) => prev.map((s) => s.id === scenario.id ? { ...s, annualRate: v } : s))} min={0} max={50} step={0.5} suffix="%" />
                    {last && (
                      <div className="pt-2 border-t border-slate-700 space-y-1">
                        <p className="text-sm text-slate-400">القيمة النهائية: <span className="text-white font-bold">{formatCurrency(last.portfolioValue, currency)}</span></p>
                        <p className="text-sm text-slate-400">المستثمر: <span className="text-blue-400 font-bold">{formatCurrency(last.invested, currency)}</span></p>
                        <p className="text-sm text-slate-400">الأرباح: <span className="text-amber-400 font-bold">{formatCurrency(last.earnings, currency)}</span></p>
                        <p className="text-sm text-slate-400">النمو: <span className="text-purple-400 font-bold">{last.growthPercent}%</span></p>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Comparison Chart */}
            <div className="bg-slate-800/40 backdrop-blur border border-slate-700 rounded-2xl p-6">
              <h3 className="text-white font-bold mb-4">مقارنة السيناريوهات — القيمة الإجمالية</h3>
              <ResponsiveContainer width="100%" height={400}>
                <AreaChart data={scenarioChartData}>
                  <defs>
                    {scenarios.map((s) => (
                      <linearGradient key={s.id} id={`sg-${s.id}`} x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor={s.color} stopOpacity={0.25} />
                        <stop offset="100%" stopColor={s.color} stopOpacity={0.02} />
                      </linearGradient>
                    ))}
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                  <XAxis dataKey="year" stroke="#475569" tick={{ fill: '#94a3b8', fontSize: 12 }} tickLine={false} />
                  <YAxis stroke="#475569" tick={{ fill: '#94a3b8', fontSize: 12 }} tickFormatter={formatCompact} tickLine={false} width={60} />
                  <Tooltip content={<ChartTooltip currency={currency} />} />
                  <Legend wrapperStyle={{ paddingTop: '10px' }} formatter={(v) => <span style={{ color: '#cbd5e1', fontSize: 13 }}>{v}</span>} />
                  {scenarios.map((s) => (
                    <Area key={s.id} type="monotone" dataKey={s.name} stroke={s.color} fill={`url(#sg-${s.id})`} strokeWidth={2.5} animationDuration={800} />
                  ))}
                </AreaChart>
              </ResponsiveContainer>
            </div>

            {/* Comparison Table */}
            <div className="bg-slate-800/40 backdrop-blur border border-slate-700 rounded-2xl p-6 overflow-x-auto">
              <h3 className="text-white font-bold mb-4">جدول المقارنة</h3>
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-slate-700">
                    {['السيناريو', 'العائد', 'المستثمر', 'القيمة النهائية', 'الأرباح', 'النمو'].map((h) => (
                      <th key={h} className="py-3 px-3 text-slate-400 font-medium text-right">{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {scenarioData.map((s) => {
                    const last = s.data[s.data.length - 1];
                    if (!last) return null;
                    return (
                      <tr key={s.id} className="border-b border-slate-700/50 hover:bg-slate-700/20 transition">
                        <td className="py-3 px-3 font-medium flex items-center gap-2">
                          <span className="w-2.5 h-2.5 rounded-full" style={{ backgroundColor: s.color }} />
                          <span className="text-white">{s.name}</span>
                        </td>
                        <td className="py-3 px-3 text-slate-300 font-mono text-right">{s.annualRate}%</td>
                        <td className="py-3 px-3 text-blue-400 font-mono text-right">{formatCurrency(last.invested, currency)}</td>
                        <td className="py-3 px-3 text-emerald-400 font-mono text-right font-bold">{formatCurrency(last.portfolioValue, currency)}</td>
                        <td className="py-3 px-3 text-amber-400 font-mono text-right">{formatCurrency(last.earnings, currency)}</td>
                        <td className="py-3 px-3 text-purple-400 font-mono text-right">{last.growthPercent}%</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* ═══ Goal Seek Tab ═══ */}
        {activeTab === 'goal' && (
          <div className="grid lg:grid-cols-2 gap-6">
            {/* Inputs */}
            <div className="bg-slate-800/40 backdrop-blur border border-slate-700 rounded-2xl p-6 space-y-5">
              <h2 className="text-white font-bold text-lg flex items-center gap-2">
                <span className="w-2 h-2 bg-amber-400 rounded-full" />
                حاسبة الهدف العكسي
              </h2>
              <p className="text-sm text-slate-400">حدد هدفك المالي والأداة تحسب كم تحتاج تستثمر شهرياً للوصول إليه</p>
              <InputField label="المبلغ المستهدف" value={goalInputs.targetAmount} onChange={(v) => setGoalInputs((p) => ({ ...p, targetAmount: v }))} min={0} max={10000000} step={1000} suffix={currency === 'USD' ? '$' : 'د.أ'} />
              <InputField label="السنة المستهدفة" value={goalInputs.targetYear} onChange={(v) => setGoalInputs((p) => ({ ...p, targetYear: v }))} min={1} max={50} step={1} suffix="سنة" />
              <InputField label="رأس المال الابتدائي" value={goalInputs.initialAmount} onChange={(v) => setGoalInputs((p) => ({ ...p, initialAmount: v }))} min={0} max={1000000} step={1000} suffix={currency === 'USD' ? '$' : 'د.أ'} />
              <InputField label="معدل العائد السنوي المتوقع" value={goalInputs.annualRate} onChange={(v) => setGoalInputs((p) => ({ ...p, annualRate: v }))} min={0} max={50} step={0.5} suffix="%" />
            </div>

            {/* Results */}
            <div className="space-y-6">
              {/* Main Result */}
              <div className={`border rounded-2xl p-8 text-center space-y-4 ${goalResult.isAlreadyReached ? 'bg-gradient-to-br from-amber-500/10 to-emerald-500/10 border-amber-500/20' : 'bg-gradient-to-br from-emerald-500/10 to-blue-500/10 border-emerald-500/20'}`}>
                {goalResult.isAlreadyReached ? (
                  <>
                    <p className="text-amber-400 text-sm uppercase tracking-wider">الهدف محقق بالفعل!</p>
                    <p className="text-4xl font-bold text-emerald-400">
                      {formatCurrency(Math.round(goalResult.futureValueOfInitial), currency)}
                    </p>
                    <p className="text-slate-400 text-sm">
                      رأس المال الابتدائي{' '}
                      {formatCurrency(goalInputs.initialAmount, currency)} كافٍ للوصول إلى{' '}
                      {formatCurrency(goalInputs.targetAmount, currency)} خلال {goalInputs.targetYear} سنة
                    </p>
                    <p className="text-emerald-400 text-xs">لا حاجة لإيداعات شهرية إضافية</p>
                  </>
                ) : (
                  <>
                    <p className="text-slate-400 text-sm uppercase tracking-wider">المبلغ الشهري المطلوب</p>
                    <p className="text-5xl font-bold text-emerald-400">
                      {formatCurrency(Math.ceil(goalResult.monthlyPayment), currency)}
                    </p>
                    <p className="text-slate-500 text-sm">
                      شهرياً لمدة {goalInputs.targetYear} سنة للوصول إلى {formatCurrency(goalInputs.targetAmount, currency)}
                    </p>
                  </>
                )}
              </div>

              {/* Details */}
              <div className="bg-slate-800/40 backdrop-blur border border-slate-700 rounded-2xl p-6 space-y-3">
                <h3 className="text-white font-bold">تفاصيل الحساب</h3>
                <div className="space-y-2 text-sm">
                  <div className="flex justify-between py-2 border-b border-slate-700/50">
                    <span className="text-slate-400">المبلغ المستهدف</span>
                    <span className="text-white font-mono">{formatCurrency(goalInputs.targetAmount, currency)}</span>
                  </div>
                  <div className="flex justify-between py-2 border-b border-slate-700/50">
                    <span className="text-slate-400">رأس المال الابتدائي</span>
                    <span className="text-blue-400 font-mono">{formatCurrency(goalInputs.initialAmount, currency)}</span>
                  </div>
                  <div className="flex justify-between py-2 border-b border-slate-700/50">
                    <span className="text-slate-400">قيمة رأس المال بعد {goalInputs.targetYear} سنة</span>
                    <span className="text-cyan-400 font-mono">{formatCurrency(Math.round(goalResult.futureValueOfInitial), currency)}</span>
                  </div>
                  <div className="flex justify-between py-2 border-b border-slate-700/50">
                    <span className="text-slate-400">الفجوة المطلوب تغطيتها</span>
                    <span className="text-amber-400 font-mono">
                      {formatCurrency(Math.max(0, goalInputs.targetAmount - Math.round(goalResult.futureValueOfInitial)), currency)}
                    </span>
                  </div>
                  {!goalResult.isAlreadyReached && (
                    <>
                      <div className="flex justify-between py-2 border-b border-slate-700/50">
                        <span className="text-slate-400">إجمالي الاشتراكات الشهرية</span>
                        <span className="text-purple-400 font-mono">
                          {formatCurrency(Math.ceil(goalResult.monthlyPayment) * goalInputs.targetYear * 12, currency)}
                        </span>
                      </div>
                      <div className="flex justify-between py-2">
                        <span className="text-slate-400">إجمالي الأرباح المتوقعة</span>
                        <span className="text-emerald-400 font-mono">
                          {formatCurrency(
                            Math.max(0, goalInputs.targetAmount - goalInputs.initialAmount - Math.ceil(goalResult.monthlyPayment) * goalInputs.targetYear * 12),
                            currency
                          )}
                        </span>
                      </div>
                    </>
                  )}
                </div>
              </div>

              {/* Goal Projection Chart */}
              {!goalResult.isAlreadyReached && goalProjection.length > 0 && (
                <div className="bg-slate-800/40 backdrop-blur border border-slate-700 rounded-2xl p-6">
                  <h3 className="text-white font-bold mb-4">مسورة النمو نحو الهدف</h3>
                  <ResponsiveContainer width="100%" height={280}>
                    <AreaChart data={goalChartData}>
                      <defs>
                        <linearGradient id="gGoalInvested" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#3b82f6" stopOpacity={0.3} />
                          <stop offset="100%" stopColor="#3b82f6" stopOpacity={0.02} />
                        </linearGradient>
                        <linearGradient id="gGoalEarnings" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#10b981" stopOpacity={0.3} />
                          <stop offset="100%" stopColor="#10b981" stopOpacity={0.02} />
                        </linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" />
                      <XAxis dataKey="year" stroke="#475569" tick={{ fill: '#94a3b8', fontSize: 12 }} tickLine={false} />
                      <YAxis stroke="#475569" tick={{ fill: '#94a3b8', fontSize: 12 }} tickFormatter={formatCompact} tickLine={false} width={60} />
                      <Tooltip content={<GoalTooltip currency={currency} targetAmount={goalInputs.targetAmount} />} />
                      <ReferenceLine y={goalInputs.targetAmount} stroke="#fbbf24" strokeDasharray="8 4" strokeWidth={2} label={{ value: 'الهدف', fill: '#fbbf24', fontSize: 12, position: 'right' }} />
                      <Area type="monotone" dataKey="المبلغ المستثمر" stackId="1" stroke="#3b82f6" fill="url(#gGoalInvested)" strokeWidth={2} animationDuration={800} />
                      <Area type="monotone" dataKey="الأرباح" stackId="1" stroke="#10b981" fill="url(#gGoalEarnings)" strokeWidth={2} animationDuration={800} />
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
              )}
            </div>
          </div>
        )}
      </main>

      <footer className="border-t border-slate-800 mt-12 py-6 text-center text-sm text-slate-600">
        حاسبة الاستثمار التفاعلية — جميع الحسابات تقريبية ولا تُعتبر نصيحة مالية
      </footer>
    </div>
  );
}
