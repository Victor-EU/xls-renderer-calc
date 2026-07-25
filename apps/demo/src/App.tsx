import { useCallback, useEffect, useMemo, useState } from 'react';
import { loadXlsx, type PreviewDocument } from '@xlscalc/xlsx-preview';
import { ExcelView, plainText } from '@xlscalc/xlsx-preview/view';

/**
 * The harness around the renderer.
 *
 * The banner is the part that changed meaning rather than shape. It used to say
 * "this file needs a recalc — 75 cells render blank", which was a diagnosis the
 * user could do nothing about. It now says what was computed, how long it took,
 * and — the part that matters — what could *not* be computed.
 */

const SAMPLES = [
  { label: 'Financial model (recalculated)', file: 'financial-model.xlsx' },
  { label: 'Same model, no cached values', file: 'financial-model-nocache.xlsx' },
  { label: 'Hardcoded total (audit demo)', file: 'hardcoded-total.xlsx' },
  { label: 'Formula tour', file: 'formula-tour.xlsx' },
];

export default function App() {
  const [doc, setDoc] = useState<PreviewDocument | null>(null);
  const [name, setName] = useState('');
  const [active, setActive] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const [gridlines, setGridlines] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [showProvenance, setShowProvenance] = useState(false);
  const [diffMode, setDiffMode] = useState(false);
  const [selected, setSelected] = useState<{ row: number; col: number } | null>(null);

  const open = useCallback(async (buf: ArrayBuffer, label: string) => {
    setBusy(true);
    setError(null);
    try {
      const next = await loadXlsx(buf);
      setDoc(next);
      setName(label);
      setActive(0);
      setSelected(null);
    } catch (e) {
      setDoc(null);
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }, []);

  const loadSample = useCallback(
    async (file: string, label: string) => {
      const res = await fetch(`/${file}`);
      if (!res.ok) {
        setError(`sample ${file} not found`);
        return;
      }
      await open(await res.arrayBuffer(), label);
    },
    [open],
  );

  useEffect(() => {
    void loadSample(SAMPLES[1]!.file, SAMPLES[1]!.label);
  }, [loadSample]);

  const onDrop = useCallback(
    async (e: React.DragEvent) => {
      e.preventDefault();
      const file = e.dataTransfer.files[0];
      if (file) await open(await file.arrayBuffer(), file.name);
    },
    [open],
  );

  const highlight = useMemo(() => {
    if (!doc || !selected) return undefined;
    const cells = doc.model.engine.precedents(active, selected.row, selected.col);
    return new Set(cells.filter((c) => c.sheet === active).map((c) => `${c.row}:${c.col}`));
  }, [doc, selected, active]);

  const flagged = useMemo(() => {
    if (!doc) return undefined;
    return new Set(doc.model.hardcoded.filter((h) => h.sheet === active).map((h) => `${h.row}:${h.col}`));
  }, [doc, active]);

  const sheet = doc?.sheets[active];
  const worksheet = doc && sheet ? (doc.styled.getWorksheet(sheet.name) ?? doc.styled.worksheets[active]) : undefined;
  const stats = doc?.model.report.stats;

  return (
    <div className="app" onDragOver={(e) => e.preventDefault()} onDrop={onDrop}>
      <header className="top">
        <div className="title">
          <strong>xlsx preview</strong>
          <span className="sub">formulas computed in the browser · nothing uploaded</span>
        </div>
        <div className="samples">
          {SAMPLES.map((s) => (
            <button key={s.file} onClick={() => void loadSample(s.file, s.label)}>
              {s.label}
            </button>
          ))}
          <label className="filebtn">
            open a file…
            <input
              type="file"
              accept=".xlsx"
              onChange={async (e) => {
                const file = e.target.files?.[0];
                if (file) await open(await file.arrayBuffer(), file.name);
              }}
            />
          </label>
        </div>
      </header>

      {error && <div className="banner bad">could not read this file — {error}</div>}
      {busy && <div className="banner">reading…</div>}

      {doc && stats && (
        <>
          <div className={`banner ${stats.unsupported + stats.circular > 0 ? 'warn' : 'good'}`}>
            {stats.unsupported + stats.circular === 0 ? (
              <>
                <strong>Computed live in your browser.</strong> {stats.computed + stats.volatile} of{' '}
                {stats.formulas} formulas · 0 unsupported · {doc.evalMs} ms
              </>
            ) : (
              <>
                <strong>
                  {stats.computed + stats.volatile}/{stats.formulas} computed.
                </strong>{' '}
                {stats.unsupported + stats.circular} cells could not be computed and are shown as ⚠ —
                never as a guessed number.
              </>
            )}
          </div>

          <div className="chips">
            <span className="chip file">{name}</span>
            <Chip label="computed" value={stats.computed} tone="good" />
            {stats.volatile > 0 && <Chip label="volatile" value={stats.volatile} />}
            <Chip label="from file" value={stats.cached} />
            <Chip label="unsupported" value={stats.unsupported} tone={stats.unsupported ? 'bad' : undefined} />
            {stats.circular > 0 && <Chip label="circular" value={stats.circular} tone="bad" />}
            <Chip
              label="disagrees with file"
              value={stats.mismatched + doc.model.hardcoded.length}
              tone={stats.mismatched + doc.model.hardcoded.length ? 'bad' : undefined}
            />
            <span className="chip">parse {doc.parseMs} ms</span>
            <span className="chip">eval {doc.evalMs} ms</span>
            {doc.model.facts.uncached > 0 && (
              <span className="chip warn">{doc.model.facts.uncached} had no value in the file</span>
            )}
          </div>

          <div className="controls">
            <div className="tabs">
              {doc.sheets.map((s, i) => (
                <button
                  key={s.name}
                  className={i === active ? 'tab on' : 'tab'}
                  onClick={() => {
                    setActive(i);
                    setSelected(null);
                  }}
                >
                  {s.name}
                  {s.formulas > 0 && <em>{s.formulas}f</em>}
                </button>
              ))}
            </div>
            <div className="toggles">
              <label>
                <input type="checkbox" checked={gridlines} onChange={(e) => setGridlines(e.target.checked)} />
                gridlines
              </label>
              <label>
                <input
                  type="checkbox"
                  checked={showProvenance}
                  onChange={(e) => setShowProvenance(e.target.checked)}
                />
                mark computed
              </label>
              <label>
                <input type="checkbox" checked={diffMode} onChange={(e) => setDiffMode(e.target.checked)} />
                diff vs file
              </label>
              <span className="zoom">
                <button onClick={() => setZoom((z) => Math.max(0.5, +(z - 0.1).toFixed(2)))}>−</button>
                {Math.round(zoom * 100)}%
                <button onClick={() => setZoom((z) => Math.min(2, +(z + 0.1).toFixed(2)))}>+</button>
              </span>
            </div>
          </div>

          {doc.model.report.gaps.length > 0 && (
            <div className="gaps">
              <strong>Not computed</strong>
              <ul>
                {doc.model.report.gaps.map((g) => (
                  <li key={`${g.code}:${g.subject}`}>
                    <code>{g.subject}</code> × {g.count} — {g.reason} <span className="at">{g.sample}</span>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {diffMode && (doc.model.report.mismatches.length > 0 || doc.model.hardcoded.length > 0) && (
            <div className="gaps bad">
              <strong>Numbers this model states that its own formulas do not support</strong>
              <ul>
                {doc.model.hardcoded.map((h) => (
                  <li key={`hc-${h.sheetName}-${h.address}`}>
                    <code>
                      {h.sheetName}!{h.address}
                    </code>{' '}
                    is typed in as <b>{plainText(h.stated)}</b>, but the rest of the {h.axis} (
                    {h.pattern.join(', ')}) uses <code>={h.formula}</code>, which is{' '}
                    <b>{plainText(h.expected)}</b>
                  </li>
                ))}
                {doc.model.report.mismatches.slice(0, 40).map((m) => (
                  <li key={m.address}>
                    <code>{m.address}</code> — the file stores <b>{plainText(m.cached)}</b>, the formula
                    computes <b>{plainText(m.computed)}</b>
                  </li>
                ))}
              </ul>
            </div>
          )}

          <div className="sheet">
            {worksheet && (
              <ExcelView
                worksheet={worksheet}
                model={doc.model}
                sheetIndex={active}
                gridlines={gridlines}
                zoom={zoom}
                showProvenance={showProvenance}
                diffMode={diffMode}
                flagged={flagged}
                highlight={highlight}
                onSelect={setSelected}
              />
            )}
          </div>

          {selected && <Inspector doc={doc} sheet={active} selected={selected} />}
        </>
      )}

      {!doc && !busy && <div className="drop">drop an .xlsx here</div>}
    </div>
  );
}

function Chip({ label, value, tone }: { label: string; value: number; tone?: 'good' | 'bad' }) {
  return (
    <span className={`chip ${tone ?? ''}`}>
      {label} <b>{value}</b>
    </span>
  );
}

function Inspector({
  doc,
  sheet,
  selected,
}: {
  doc: PreviewDocument;
  sheet: number;
  selected: { row: number; col: number };
}) {
  const cell = doc.model.cell(sheet, selected.row, selected.col);
  const precedents = doc.model.engine.precedents(sheet, selected.row, selected.col, 24);
  const address = `${colName(selected.col)}${selected.row}`;
  if (!cell) return <div className="inspector">{address} — empty</div>;

  return (
    <div className="inspector">
      <div className="row">
        <b>{address}</b>
        <span className={`prov ${cell.provenance}`}>{cell.provenance}</span>
        {cell.formula && <code>={cell.formula.replace(/^=/, '')}</code>}
      </div>
      <div className="row">
        <span>
          value <b>{cell.provenance === 'unsupported' ? '⚠ not computed' : plainText(cell.value)}</b>
        </span>
        {cell.cached !== undefined && plainText(cell.cached) !== plainText(cell.value) && (
          <span className="warn">file said {plainText(cell.cached)}</span>
        )}
        {cell.reason && <span className="warn">{cell.reason}</span>}
      </div>
      {precedents.length > 0 && (
        <div className="row">
          reads{' '}
          {precedents.map((p) => `${p.sheet === sheet ? '' : `${doc.sheets[p.sheet]?.name}!`}${colName(p.col)}${p.row}`).join(', ')}
        </div>
      )}
    </div>
  );
}

function colName(col: number): string {
  let n = col;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
