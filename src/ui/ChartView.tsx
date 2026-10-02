/**
 * Chart rendering.
 *
 * Plain SVG so charts stay crisp at any zoom, cost nothing to redraw and can be
 * exported by simply serialising the node.
 */

import { useMemo } from 'react';
import type { ChartData, ChartSpec } from '../core/chartdata';
import { SERIES_COLORS } from '../core/chartdata';

interface ChartViewProps {
  chart: ChartSpec;
  data: ChartData;
  width?: number;
  height?: number;
  onTypeChange?: (type: ChartSpec['type']) => void;
  onToggle?: (key: 'showLegend' | 'showTitle' | 'showGridlines' | 'showDataLabels' | 'stacked') => void;
  onDelete?: () => void;
}

const PAD = { top: 24, right: 16, bottom: 34, left: 48 };

export default function ChartView({
  chart,
  data,
  width = 520,
  height = 320,
  onTypeChange,
  onToggle,
  onDelete,
}: ChartViewProps) {
  const titleH = chart.showTitle ? 22 : 6;
  const legendH = chart.showLegend ? 20 : 0;
  const plotW = width - PAD.left - PAD.right;
  const plotH = height - PAD.top - PAD.bottom - titleH - legendH;
  const isPie = chart.type === 'pie' || chart.type === 'doughnut';

  const ticks = useMemo(() => niceTicks(data.min, data.max, 5), [data.min, data.max]);
  const yMin = ticks[0];
  const yMax = ticks[ticks.length - 1];
  const span = yMax - yMin || 1;

  const yOf = (v: number): number => PAD.top + titleH + plotH - ((v - yMin) / span) * plotH;

  const xFor = (i: number, count: number): number => {
    if (count <= 1) return PAD.left + plotW / 2;
    return PAD.left + (plotW / count) * (i + 0.5);
  };

  return (
    <div className="chart-card">
      <header>
        <span>{chart.name}</span>
        <span className="spacer" />
        {onTypeChange && (
          <select value={chart.type} onChange={(e) => onTypeChange(e.target.value as ChartSpec['type'])}>
            <option value="column">Column</option>
            <option value="bar">Bar</option>
            <option value="line">Line</option>
            <option value="area">Area</option>
            <option value="pie">Pie</option>
            <option value="doughnut">Doughnut</option>
            <option value="scatter">Scatter</option>
            <option value="radar">Radar</option>
          </select>
        )}
        {onToggle && (
          <>
            <button className="tool icon" title="Legend" onClick={() => onToggle('showLegend')}>◧</button>
            <button className="tool icon" title="Gridlines" onClick={() => onToggle('showGridlines')}>⊞</button>
            <button className="tool icon" title="Data labels" onClick={() => onToggle('showDataLabels')}>123</button>
          </>
        )}
        {onDelete && (
          <button className="tool icon" title="Delete chart" onClick={onDelete}>✕</button>
        )}
      </header>

      <div className="chart-body">
        <svg className="chart-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label={chart.title}>
          {chart.showTitle && (
            <text x={width / 2} y={14} textAnchor="middle" fontSize={12} fontWeight={600} fill="#1f2733">
              {chart.title}
            </text>
          )}

          {isPie ? (
            <Pie data={data} cx={PAD.left + plotW / 2} cy={PAD.top + titleH + plotH / 2} r={Math.min(plotW, plotH) / 2} chart={chart} />
          ) : chart.type === 'bar' ? (
            <Bars data={data} plotW={plotW} plotH={plotH} yOf={yOf} xFor={xFor} top={PAD.top + titleH} chart={chart} width={width} />
          ) : chart.type === 'line' || chart.type === 'area' || chart.type === 'scatter' ? (
            <Lines data={data} plotW={plotW} plotH={plotH} yOf={yOf} xFor={xFor} top={PAD.top + titleH} chart={chart} width={width} />
          ) : (
            <Columns data={data} plotW={plotW} plotH={plotH} yOf={yOf} xFor={xFor} top={PAD.top + titleH} chart={chart} width={width} />
          )}

          {/* value axis */}
          {!isPie && chart.showGridlines && (
            <g>
              {ticks.map((t) => (
                <g key={t}>
                  <line
                    x1={PAD.left}
                    x2={PAD.left + plotW}
                    y1={yOf(t)}
                    y2={yOf(t)}
                    stroke="#eceff3"
                  />
                  <text x={PAD.left - 6} y={yOf(t) + 3} textAnchor="end" fontSize={10} fill="#6b7684">
                    {formatTick(t)}
                  </text>
                </g>
              ))}
              <line x1={PAD.left} x2={PAD.left + plotW} y1={yOf(yMin)} y2={yOf(yMin)} stroke="#b9c0c8" />
            </g>
          )}

          {/* category axis */}
          {!isPie && (
            <g>
              {data.categories.map((label, i) =>
                data.categories.length <= 24 ? (
                  <text
                    key={`${label}-${i}`}
                    x={xFor(i, data.categories.length)}
                    y={PAD.top + titleH + plotH + 14}
                    textAnchor="middle"
                    fontSize={10}
                    fill="#6b7684"
                  >
                    {truncate(label, 10)}
                  </text>
                ) : null,
              )}
              <line
                x1={PAD.left}
                x2={PAD.left + plotW}
                y1={PAD.top + titleH + plotH}
                y2={PAD.top + titleH + plotH}
                stroke="#b9c0c8"
              />
            </g>
          )}
        </svg>

        {chart.showLegend && (
          <div className="legend">
            {(isPie
              ? (data.series[0]?.points ?? []).map((p, i) => ({ name: p.label, color: SERIES_COLORS[i % SERIES_COLORS.length] }))
              : data.series.map((s, i) => ({ name: s.name, color: SERIES_COLORS[i % SERIES_COLORS.length] }))
            ).map((item, i) => (
              <span className="item" key={`${item.name}-${i}`}>
                <span className="swatch" style={{ background: item.color }} />
                {truncate(item.name, 20)}
              </span>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------- marks */

interface MarkProps {
  data: ChartData;
  plotW: number;
  plotH: number;
  yOf: (v: number) => number;
  xFor: (i: number, count: number) => number;
  top: number;
  chart: ChartSpec;
  width: number;
}

function Columns({ data, plotW, plotH, yOf, xFor, top, chart, width }: MarkProps) {
  const count = data.categories.length || 1;
  const n = data.series.length || 1;
  const slot = plotW / count;
  const barW = Math.max(2, (slot * 0.7) / n);
  const zero = yOf(0);

  return (
    <g>
      {data.series.map((s, si) => (
        <g key={s.name}>
          {s.points.map((p, i) => {
            const x = xFor(i, count) - (n * barW) / 2 + si * barW;
            const y = yOf(p.value);
            const h = Math.abs(zero - y);
            const fill = chart.series[si]?.color ?? SERIES_COLORS[si % SERIES_COLORS.length];
            const stackedBase = chart.stacked ? stackedOffset(data, si) : 0;
            return (
              <rect
                key={`${s.name}-${i}`}
                x={x}
                y={Math.min(y + stackedBase, zero)}
                width={barW - 1}
                height={Math.max(1, h)}
                fill={fill}
                rx={1}
              >
                <title>{`${s.name}: ${p.label} = ${p.value}`}</title>
              </rect>
            );
          })}
        </g>
      ))}
      {chart.showDataLabels &&
        data.series.map((s, si) =>
          s.points.map((p, i) => (
            <text
              key={`l-${s.name}-${i}`}
              x={xFor(i, count) - (n * barW) / 2 + si * barW + barW / 2}
              y={yOf(p.value) - 4}
              textAnchor="middle"
              fontSize={9}
              fill="#5b6572"
            >
              {formatTick(p.value)}
            </text>
          )),
        )}
    </g>
  );
}

function Bars({ data, plotW, plotH, yOf, xFor, top, chart, width }: MarkProps) {
  const count = data.categories.length || 1;
  const n = data.series.length || 1;
  const slot = plotH / count;
  const barH = Math.max(2, (slot * 0.7) / n);

  return (
    <g>
      {data.series.map((s, si) => (
        <g key={s.name}>
          {s.points.map((p, i) => {
            const y = top + (plotH / count) * (i + 0.5) - (n * barH) / 2 + si * barH;
            const w = Math.abs(xFor(p.value, 1) - PAD.left);
            const fill = chart.series[si]?.color ?? SERIES_COLORS[si % SERIES_COLORS.length];
            return (
              <rect
                key={`${s.name}-${i}`}
                x={PAD.left}
                y={y}
                width={Math.max(1, w)}
                height={barH - 1}
                fill={fill}
                rx={1}
              >
                <title>{`${s.name}: ${p.label} = ${p.value}`}</title>
              </rect>
            );
          })}
        </g>
      ))}
    </g>
  );
}

function Lines({ data, plotW, plotH, yOf, xFor, top, chart, width }: MarkProps) {
  const count = data.categories.length || 1;

  return (
    <g>
      {chart.type === 'area' &&
        data.series.map((s, si) => {
          const fill = chart.series[si]?.color ?? SERIES_COLORS[si % SERIES_COLORS.length];
          const d =
            `M ${xFor(0, count)} ${yOf(0)} ` +
            s.points.map((p, i) => `L ${xFor(i, count)} ${yOf(p.value)}`).join(' ') +
            ` L ${xFor(s.points.length - 1, count)} ${yOf(0)} Z`;
          return <path key={`a-${s.name}`} d={d} fill={fill} opacity={0.28} />;
        })}

      {data.series.map((s, si) => {
        const color = chart.series[si]?.color ?? SERIES_COLORS[si % SERIES_COLORS.length];
        if (chart.type === 'scatter') {
          return (
            <g key={`s-${s.name}`}>
              {s.points.map((p, i) => (
                <circle key={i} cx={xFor(i, count)} cy={yOf(p.value)} r={3} fill={color}>
                  <title>{`${s.name}: ${p.label} = ${p.value}`}</title>
                </circle>
              ))}
            </g>
          );
        }
        const d = s.points
          .map((p, i) => `${i === 0 ? 'M' : 'L'} ${xFor(i, count)} ${yOf(p.value)}`)
          .join(' ');
        return (
          <g key={`l-${s.name}`}>
            <path d={d} fill="none" stroke={color} strokeWidth={2} strokeLinejoin="round" strokeLinecap="round" />
            {!chart.smooth &&
              s.points.map((p, i) => (
                <circle key={i} cx={xFor(i, count)} cy={yOf(p.value)} r={2.6} fill={color}>
                  <title>{`${s.name}: ${p.label} = ${p.value}`}</title>
                </circle>
              ))}
          </g>
        );
      })}
    </g>
  );
}

function Pie({ data, cx, cy, r, chart }: { data: ChartData; cx: number; cy: number; r: number; chart: ChartSpec }) {
  const points = data.series[0]?.points ?? [];
  const total = points.reduce((s, p) => s + Math.abs(p.value), 0);
  if (total === 0) return <circle cx={cx} cy={cy} r={r} fill="#eceff3" />;

  const isDoughnut = chart.type === 'doughnut';
  const inner = isDoughnut ? r * (chart.holeSize ?? 0.55) : 0;
  let angle = -Math.PI / 2;

  return (
    <g>
      {points.map((p, i) => {
        const sweep = (Math.abs(p.value) / total) * Math.PI * 2;
        const end = angle + sweep;
        const color = SERIES_COLORS[i % SERIES_COLORS.length];
        const path = isDoughnut ? donutArc(cx, cy, r, inner, angle, end) : pieArc(cx, cy, r, angle, end);
        angle = end;
        return (
          <path key={`${p.label}-${i}`} d={path} fill={color} stroke="#fff" strokeWidth={1}>
            <title>{`${p.label}: ${p.value}`}</title>
          </path>
        );
      })}
      {isDoughnut && <text x={cx} y={cy + 4} textAnchor="middle" fontSize={11} fill="#5b6572">{formatTick(total)}</text>}
    </g>
  );
}

function pieArc(cx: number, cy: number, r: number, a0: number, a1: number): string {
  const x0 = cx + r * Math.cos(a0);
  const y0 = cy + r * Math.sin(a0);
  const x1 = cx + r * Math.cos(a1);
  const y1 = cy + r * Math.sin(a1);
  const large = a1 - a0 > Math.PI ? 1 : 0;
  return `M ${cx} ${cy} L ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} Z`;
}

function donutArc(cx: number, cy: number, rOuter: number, rInner: number, a0: number, a1: number): string {
  const large = a1 - a0 > Math.PI ? 1 : 0;
  const x0o = cx + rOuter * Math.cos(a0);
  const y0o = cy + rOuter * Math.sin(a0);
  const x1o = cx + rOuter * Math.cos(a1);
  const y1o = cy + rOuter * Math.sin(a1);
  const x1i = cx + rInner * Math.cos(a1);
  const y1i = cy + rInner * Math.sin(a1);
  const x0i = cx + rInner * Math.cos(a0);
  const y0i = cy + rInner * Math.sin(a0);
  return (
    `M ${x0o} ${y0o} A ${rOuter} ${rOuter} 0 ${large} 1 ${x1o} ${y1o} ` +
    `L ${x1i} ${y1i} A ${rInner} ${rInner} 0 ${large} 0 ${x0i} ${y0i} Z`
  );
}

function stackedOffset(data: ChartData, upto: number): number {
  let sum = 0;
  for (let i = 0; i < upto; i++) sum += data.series[i]?.points[0]?.value ?? 0;
  return -sum;
}

/* ------------------------------------------------------------------ scales */

/** Axis ticks on a round 1/2/5 step. */
export function niceTicks(min: number, max: number, count: number): number[] {
  if (min === max) {
    min = Math.min(0, min);
    max = max || 1;
  }
  const raw = (max - min) / Math.max(1, count - 1);
  const mag = 10 ** Math.floor(Math.log10(Math.abs(raw) || 1));
  const norm = raw / mag;
  const step = (norm >= 5 ? 10 : norm >= 2 ? 5 : norm >= 1 ? 2 : 1) * mag;
  const start = Math.floor(min / step) * step;
  const end = Math.ceil(max / step) * step;
  const out: number[] = [];
  for (let v = start; v <= end + step / 2; v += step) {
    out.push(Math.abs(v) < step / 1e6 ? 0 : v);
  }
  return out;
}

function formatTick(v: number): string {
  const abs = Math.abs(v);
  if (abs >= 1e9) return `${(v / 1e9).toFixed(1)}B`;
  if (abs >= 1e6) return `${(v / 1e6).toFixed(1)}M`;
  if (abs >= 1e4) return `${(v / 1e3).toFixed(1)}k`;
  if (Number.isInteger(v)) return String(v);
  return String(Number(v.toFixed(2)));
}

function truncate(s: string, n: number): string {
  return s.length > n ? `${s.slice(0, n - 1)}…` : s;
}
