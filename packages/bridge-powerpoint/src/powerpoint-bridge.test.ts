import { afterEach, describe, it, expect, vi } from 'vitest';
import {
  asChangeId,
  DocStateSnapshotSchema,
  ResolvedContextSchema,
  type ActuationRequest,
  type ContextRef,
} from '@ge/contracts';
import type { HostEvent } from '@ge/triggers';
import { MAX_READ_SLIDES, PowerPointBridge } from './powerpoint-bridge.js';

/**
 * Behavioural tests for the {@link PowerPointBridge} host wiring — the file that touches the
 * `PowerPoint.run` / `Office` globals directly. There is no host-port indirection here (unlike the
 * Word bridge), so we install a self-contained in-memory PowerPoint + Office simulator onto
 * `globalThis` for the duration of each test. The simulator re-implements ONLY the slice of the
 * Office.js object model the bridge actually drives (slides → shapes → textRange text, the
 * compose/insert WRITE paths, the requirement gate, and the Office event bus). The REAL bridge runs
 * unchanged against it, so these assert real semantics: capability gating, bounded reads, slide
 * addressing, the two insert-slide write paths, provenance/location stamping, and event wiring.
 */

/* ─────────────────────────── in-memory host model ─────────────────────────── */

interface ShapeSeed {
  text: string;
  /** Simulated `Shape.type`; defaults to 'TextBox'. */
  type?: string;
  /** Simulated `placeholderFormat.containedType` for a 'Placeholder' shape (null = text/empty). */
  containedType?: string | null;
  /** Simulated `placeholderFormat.type` for a 'Placeholder' shape (e.g. 'Title', 'Body'). */
  placeholderType?: string;
  /** The host refuses this shape's text frame even though its type suggests one. */
  refuseText?: boolean;
  /** Simulated ShapeFill.foregroundColor (undefined until first set). */
  fillColor?: string;
  /** Simulated ShapeLineFormat.color (undefined until first set). */
  lineColor?: string;
  /** Simulated ShapeFont fields on the shape's text range. */
  font?: {
    bold?: boolean | null;
    italic?: boolean | null;
    underline?: string | null;
    color?: string | null;
    size?: number | null;
    name?: string | null;
  };
  /** Simulated native table grid backing `Shape.getTable()` (table shapes only). */
  tableGrid?: string[][];
  /** Recorded setZOrder positions applied to this shape. */
  zOrderCalls?: string[];
}
interface SlideSeed {
  id: string;
  shapes: ShapeSeed[];
  /**
   * Set on a just-added slide: its `id` is PROVISIONAL until the current `PowerPoint.run` ends, when
   * it becomes this settled id. Like PowerPoint for the web, writes addressed through the
   * provisional id (`getItem(id)` or the `items` list) fail the next sync; `getItemAt` works.
   */
  settledId?: string;
}

/** A no-op write check, for fakes not reached through a slide. */
const NO_GUARD = (): void => undefined;
interface DeckSeed {
  slides: SlideSeed[];
  /** Zero-based indices of selected slides (the bridge reads items[0]). */
  selectedIndices: number[];
  selectedShapeIds: string[];
  insertedDecks: string[];
  insertedDeckOptions: PowerPoint.InsertSlideOptions[];
  /** Shapes a `slides.add()` slide gets from the default layout (default: title + body placeholders). */
  newSlideShapes?: ShapeSeed[];
  addedShapes: Array<{
    slideId: string;
    kind: 'textBox' | 'geometric' | 'line' | 'table';
    detail: string;
    options: PowerPoint.ShapeAddOptions | PowerPoint.TableAddOptions | undefined;
  }>;
}

/** A cell whose text reads/writes one slot of the backing {@link ShapeSeed.tableGrid}. */
class FakeTableCell {
  constructor(
    private readonly grid: string[][],
    private readonly rowIndex: number,
    private readonly columnIndex: number,
  ) {}
  get text(): string {
    return this.grid[this.rowIndex]?.[this.columnIndex] ?? '';
  }
  set text(value: string) {
    const row = this.grid[this.rowIndex];
    if (row && this.columnIndex < row.length) row[this.columnIndex] = value;
  }
}

/** A simulated native table whose storage is the owning shape's `tableGrid` seed. */
class FakeTable {
  constructor(private readonly shape: ShapeSeed) {}
  load(_p?: string): this {
    return this;
  }
  get values(): string[][] {
    return (this.shape.tableGrid ?? []).map((row) => [...row]);
  }
  getCellOrNullObject(rowIndex: number, columnIndex: number): FakeTableCell {
    const grid = (this.shape.tableGrid ??= []);
    while (grid.length <= rowIndex) grid.push([]);
    const row = grid[rowIndex];
    if (row) while (row.length <= columnIndex) row.push('');
    return new FakeTableCell(grid, rowIndex, columnIndex);
  }
}

class FakeFont {
  constructor(
    private readonly font: NonNullable<ShapeSeed['font']>,
    private readonly guard: () => void = NO_GUARD,
  ) {}
  load(_p?: string): this {
    return this;
  }
  get bold(): boolean | null {
    return this.font.bold ?? null;
  }
  set bold(v: boolean) {
    this.guard();
    this.font.bold = v;
  }
  get italic(): boolean | null {
    return this.font.italic ?? null;
  }
  set italic(v: boolean) {
    this.guard();
    this.font.italic = v;
  }
  get underline(): string | null {
    return this.font.underline ?? null;
  }
  set underline(v: 'None' | 'Single') {
    this.guard();
    this.font.underline = v;
  }
  get color(): string | null {
    return this.font.color ?? null;
  }
  set color(v: string) {
    this.guard();
    this.font.color = v;
  }
  get size(): number | null {
    return this.font.size ?? null;
  }
  set size(v: number) {
    this.guard();
    this.font.size = v;
  }
  get name(): string | null {
    return this.font.name ?? null;
  }
  set name(v: string) {
    this.guard();
    this.font.name = v;
  }
}

class FakeTextRange {
  readonly font: FakeFont;
  constructor(
    private readonly shape: ShapeSeed,
    private readonly guard: () => void = NO_GUARD,
  ) {
    this.shape.font ??= {};
    this.font = new FakeFont(this.shape.font, guard);
  }
  get text(): string {
    return this.shape.text;
  }
  set text(v: string) {
    this.guard();
    this.shape.text = v;
  }
  load(_p?: string): this {
    return this;
  }
}

class FakeShapeFill {
  constructor(private readonly shape: ShapeSeed) {}
  load(_p?: string): this {
    return this;
  }
  get foregroundColor(): string {
    return this.shape.fillColor ?? '';
  }
  set foregroundColor(v: string) {
    this.shape.fillColor = v;
  }
}

class FakeShapeLineFormat {
  constructor(private readonly shape: ShapeSeed) {}
  load(_p?: string): this {
    return this;
  }
  get color(): string {
    return this.shape.lineColor ?? '';
  }
  set color(v: string) {
    this.shape.lineColor = v;
  }
}

/** Shape types the simulated host gives a text frame (mirrors the host's InvalidArgument throw). */
const FAKE_TEXT_SHAPE_TYPES = new Set([
  'GeometricShape',
  'TextBox',
  'Placeholder',
  'Callout',
  'Freeform',
]);

/**
 * Like the real host, touching a text frame / placeholder format the shape doesn't have queues an
 * operation that fails at the NEXT `ctx.sync()` — failing the whole batch, not just that shape.
 */
let pendingHostError: string | undefined;

class FakeShape {
  private readonly frame: { textRange: FakeTextRange };
  readonly fill: FakeShapeFill;
  readonly lineFormat: FakeShapeLineFormat;
  constructor(
    readonly shape: ShapeSeed,
    readonly id: string,
    readonly isNullObject = false,
    guard: () => void = NO_GUARD,
  ) {
    this.frame = { textRange: new FakeTextRange(shape, guard) };
    this.fill = new FakeShapeFill(shape);
    this.lineFormat = new FakeShapeLineFormat(shape);
  }
  get type(): string {
    return this.shape.type ?? 'TextBox';
  }
  get textFrame(): { textRange: FakeTextRange } {
    const contained = this.shape.containedType;
    const refuses =
      this.shape.refuseText === true ||
      !FAKE_TEXT_SHAPE_TYPES.has(this.type) ||
      (this.type === 'Placeholder' && contained != null && !FAKE_TEXT_SHAPE_TYPES.has(contained));
    if (refuses) pendingHostError = `InvalidArgument: ${this.type} ${this.id} has no TextFrame`;
    return this.frame;
  }
  get placeholderFormat(): {
    containedType: string | null;
    type: string;
    load(_p?: string): unknown;
  } {
    if (this.type !== 'Placeholder') {
      pendingHostError = `GeneralException: ${this.id} is not a placeholder`;
    }
    const format = {
      containedType: this.shape.containedType ?? null,
      type: this.shape.placeholderType ?? 'Unsupported',
      load: () => format,
    };
    return format;
  }
  load(_p?: string): this {
    return this;
  }
  setZOrder(position: string): void {
    (this.shape.zOrderCalls ??= []).push(position);
  }
  getTable(): FakeTable {
    if (this.type !== 'Table') pendingHostError = `InvalidArgument: ${this.id} is not a table`;
    return new FakeTable(this.shape);
  }
}

class FakeShapeCollection {
  items: FakeShape[];
  constructor(
    private readonly slide: SlideSeed,
    private readonly deck: DeckSeed,
    private readonly slideId: string,
    private readonly guard: () => void = NO_GUARD,
  ) {
    this.items = slide.shapes.map(
      (s, i) => new FakeShape(s, `${slide.id}-shape-${i}`, false, guard),
    );
  }
  load(_p?: string): this {
    return this;
  }
  getItemAt(index: number): FakeShape {
    const shape = this.items[index];
    if (!shape) throw new Error(`fake-powerpoint: no shape at ${index}`);
    return shape;
  }
  getItemOrNullObject(id: string): FakeShape {
    const index = this.items.findIndex((shape) => shape.id === id);
    const existing = this.items[index];
    if (existing) return existing;
    return new FakeShape(
      { text: '', zOrderCalls: [] },
      id || `${this.slideId}-missing-shape`,
      true,
    );
  }

  /** Append a simulated shape to the host model and the live collection, then return it. */
  private append(seed: ShapeSeed): FakeShape {
    const id = `${this.slideId}-shape-${this.items.length}`;
    this.slide.shapes.push(seed);
    this.guard();
    const shape = new FakeShape(seed, id, false, this.guard);
    this.items.push(shape);
    return shape;
  }

  addTextBox(text: string, options?: PowerPoint.ShapeAddOptions): FakeShape {
    this.deck.addedShapes.push({
      slideId: this.slideId,
      kind: 'textBox',
      detail: text,
      options,
    });
    return this.append({ text, zOrderCalls: [] });
  }

  addGeometricShape(geometryType: string, options?: PowerPoint.ShapeAddOptions): FakeShape {
    this.deck.addedShapes.push({
      slideId: this.slideId,
      kind: 'geometric',
      detail: geometryType,
      options,
    });
    return this.append({ text: '', zOrderCalls: [] });
  }

  addLine(connectorType?: string, options?: PowerPoint.ShapeAddOptions): FakeShape {
    this.deck.addedShapes.push({
      slideId: this.slideId,
      kind: 'line',
      detail: connectorType ?? 'Straight',
      options,
    });
    return this.append({ text: '', zOrderCalls: [] });
  }

  addTable(rowCount: number, columnCount: number, options?: PowerPoint.TableAddOptions): FakeShape {
    this.deck.addedShapes.push({
      slideId: this.slideId,
      kind: 'table',
      detail: `${rowCount}x${columnCount}`,
      options,
    });
    const grid: string[][] = Array.from({ length: rowCount }, () =>
      Array.from({ length: columnCount }, () => ''),
    );
    return this.append({ text: '', type: 'Table', tableGrid: grid, zOrderCalls: [] });
  }
}

class FakeSlide {
  constructor(
    private readonly slide: SlideSeed,
    readonly index: number,
    private readonly seed: DeckSeed,
    /**
     * Reached via `getItemAt` (by position) rather than by id or the `items` list. Like Office.js,
     * LOADING the proxy re-addresses it by id (`slides.getItem(id)`), so `load()` clears this.
     */
    private byPosition = false,
  ) {}
  /** A write through this slide: fails the next sync if it addresses a provisional id. */
  private readonly guard = (): void => {
    if (this.slide.settledId !== undefined && !this.byPosition) {
      pendingHostError = 'InvalidParam passed to GetItem(id)';
    }
  };
  get id(): string {
    return this.slide.id;
  }
  load(_p?: string): this {
    this.byPosition = false;
    return this;
  }
  get shapes(): FakeShapeCollection {
    return new FakeShapeCollection(this.slide, this.seed, this.slide.id, this.guard);
  }
  setSelectedShapes(shapeIds: string[]): void {
    this.seed.selectedIndices = [this.index];
    this.seed.selectedShapeIds = [...shapeIds];
  }
}

class FakeSlideCollection {
  items: FakeSlide[];
  /** Recorded calls so tests can assert what the bridge drove. */
  addCalls = 0;
  constructor(private readonly seed: DeckSeed) {
    this.items = seed.slides.map((s, i) => new FakeSlide(s, i, seed));
  }
  load(_p?: string): this {
    return this;
  }
  getCount(): { value: number } {
    return { value: this.seed.slides.length };
  }
  add(): void {
    this.addCalls += 1;
    const n = this.seed.slides.length + 1;
    this.seed.slides.push({
      id: `provisional-${n}#0`,
      settledId: `sim-slide-${n}`,
      shapes: (
        this.seed.newSlideShapes ?? [
          { text: '', type: 'Placeholder', placeholderType: 'Title', zOrderCalls: [] },
          { text: '', type: 'Placeholder', placeholderType: 'Body', zOrderCalls: [] },
        ]
      ).map((shape) => ({ ...shape, zOrderCalls: [] })),
    });
    this.items = this.seed.slides.map((s, i) => new FakeSlide(s, i, this.seed));
  }
  getItemAt(index: number): FakeSlide {
    const slide = this.seed.slides[index];
    if (!slide) throw new Error(`fake-powerpoint: no slide at ${index}`);
    return new FakeSlide(slide, index, this.seed, true);
  }
  getItem(id: string): FakeSlide {
    const index = this.seed.slides.findIndex((s) => s.id === id);
    const slide = this.seed.slides[index];
    if (!slide) throw new Error(`fake-powerpoint: no slide ${id}`);
    return new FakeSlide(slide, index, this.seed);
  }
}

class FakePresentation {
  insertCalls: string[] = [];
  /** Simulated `PageSetup` (PowerPointApi 1.10): a 10" × 5.625" deck, in points. */
  readonly pageSetup = {
    slideWidth: 720,
    slideHeight: 405,
    load(_p?: string): unknown {
      return this;
    },
  };
  constructor(private readonly seed: DeckSeed) {}
  get slides(): FakeSlideCollection {
    return new FakeSlideCollection(this.seed);
  }
  getSelectedSlides(): FakeSlideCollection {
    const all = new FakeSlideCollection(this.seed);
    const sel = new FakeSlideCollection(this.seed);
    sel.items = this.seed.selectedIndices
      .map((i) => all.items[i])
      .filter((s): s is FakeSlide => s !== undefined);
    return sel;
  }
  setSelectedSlides(slideIds: string[]): void {
    this.seed.selectedIndices = slideIds
      .map((id) => this.seed.slides.findIndex((s) => s.id === id))
      .filter((i) => i >= 0);
  }
  insertSlidesFromBase64(b64: string, options?: PowerPoint.InsertSlideOptions): void {
    this.insertCalls.push(b64);
    this.seed.insertedDecks.push(b64);
    this.seed.insertedDeckOptions.push(options ?? {});
    this.seed.slides.push({ id: `sim-inserted-${this.seed.insertedDecks.length}`, shapes: [] });
  }
}

class FakeContext {
  readonly presentation: FakePresentation;
  syncs = 0;
  constructor(seed: DeckSeed) {
    this.presentation = new FakePresentation(seed);
  }
  sync(): Promise<void> {
    this.syncs += 1;
    const error = pendingHostError;
    pendingHostError = undefined;
    return error ? Promise.reject(new Error(error)) : Promise.resolve();
  }
}

type OfficeHandler = () => void;
interface OfficeRegistry {
  added: Array<{ type: unknown; handler: OfficeHandler }>;
  removed: Array<{ type: unknown; handler: OfficeHandler }>;
}

interface InstallOpts {
  /** Requirement-set support: name → max supported version (numeric compare on the version). */
  requirements?: Record<string, number>;
  /** When true, every Office addHandlerAsync throws (host without the event bus). */
  officeThrowsOnAdd?: boolean;
  /** When true, removeHandlerAsync throws (teardown best-effort path). */
  officeThrowsOnRemove?: boolean;
}

interface Installed {
  ctxRef: { ctx?: FakeContext };
  office: OfficeRegistry;
  restore(): void;
}

/**
 * Install fake `PowerPoint` + `Office` globals modelling `seed`, returning hooks to inspect the
 * last context and the Office handler registry. `requirements` drives `isSetSupported(name, ver)`.
 */
function install(seed: DeckSeed, opts: InstallOpts = {}): Installed {
  const reqs = opts.requirements ?? { PowerPointApi: 8 };
  const ctxRef: { ctx?: FakeContext } = {};
  const office: OfficeRegistry = { added: [], removed: [] };

  pendingHostError = undefined;
  const prevPP = (globalThis as Record<string, unknown>).PowerPoint;
  const prevOffice = (globalThis as Record<string, unknown>).Office;

  (globalThis as Record<string, unknown>).PowerPoint = {
    run: async <T>(cb: (ctx: FakeContext) => Promise<T>): Promise<T> => {
      const ctx = new FakeContext(seed);
      ctxRef.ctx = ctx;
      try {
        return await cb(ctx);
      } finally {
        // A request completed: just-added slides get their real ids (see SlideSeed.settledId).
        for (const slide of seed.slides) {
          if (slide.settledId !== undefined) {
            slide.id = slide.settledId;
            delete slide.settledId;
          }
        }
      }
    },
  };

  (globalThis as Record<string, unknown>).Office = {
    EventType: {
      DocumentSelectionChanged: 'DocumentSelectionChanged',
      ActiveViewChanged: 'ActiveViewChanged',
    },
    context: {
      requirements: {
        isSetSupported: (name: string, version?: string): boolean => {
          const max = reqs[name];
          if (max === undefined) return false;
          // `max` 1.x = highest supported minor version; >= 2 = every 1.x set. Compare minors as
          // integers so '1.10' is newer than '1.5' (a float compare would read it as 1.1).
          if (max >= 2) return true;
          const minor = Number((version ?? '1.0').split('.')[1] ?? '0');
          return minor <= Math.round((max - 1) * 10);
        },
      },
      document: {
        addHandlerAsync: (type: unknown, handler: OfficeHandler): void => {
          if (opts.officeThrowsOnAdd) throw new Error('no event bus');
          office.added.push({ type, handler });
        },
        removeHandlerAsync: (type: unknown, arg: { handler: OfficeHandler }): void => {
          if (opts.officeThrowsOnRemove) throw new Error('cannot remove');
          office.removed.push({ type, handler: arg.handler });
        },
      },
    },
  };

  return {
    ctxRef,
    office,
    restore: () => {
      (globalThis as Record<string, unknown>).PowerPoint = prevPP;
      (globalThis as Record<string, unknown>).Office = prevOffice;
    },
  };
}

let installed: Installed | undefined;
afterEach(() => {
  installed?.restore();
  installed = undefined;
});

function deck(slides: SlideSeed[], selectedIndices: number[] = []): DeckSeed {
  // Deep-copy so shared fixtures (SAMPLE_SLIDES) aren't mutated by write paths across tests.
  const copy = slides.map((s) => ({
    id: s.id,
    shapes: s.shapes.map((sh) => ({
      ...sh,
      font: sh.font ? { ...sh.font } : undefined,
      tableGrid: sh.tableGrid?.map((row) => [...row]),
      zOrderCalls: [...(sh.zOrderCalls ?? [])],
    })),
  }));
  return {
    slides: copy,
    selectedIndices,
    selectedShapeIds: [],
    insertedDecks: [],
    insertedDeckOptions: [],
    addedShapes: [],
  };
}

function insertSlide(params: ActuationRequest['params'], id = 'c1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'insert-slide', surface: 'powerpoint', params };
}

function setShapeText(
  slideId: string,
  shapeId: string,
  text: string,
  id = 'shape-1',
): ActuationRequest {
  return {
    changeId: asChangeId(id),
    kind: 'set-shape-text',
    surface: 'powerpoint',
    params: { target: { slideId, shapeId }, text },
  };
}

const SAMPLE_SLIDES: SlideSeed[] = [
  {
    id: 's1',
    shapes: [
      { text: 'SLA Terms', zOrderCalls: [] },
      { text: '99.5% contracted\nMonthly window', zOrderCalls: [] },
    ],
  },
  {
    id: 's2',
    shapes: [
      { text: 'Roster', zOrderCalls: [] },
      { text: 'Pat, Sam', zOrderCalls: [] },
    ],
  },
  {
    id: 's3',
    shapes: [
      { text: 'Risk Summary', zOrderCalls: [] },
      { text: 'SLA gap flagged', zOrderCalls: [] },
    ],
  },
];

const WHOLE_DECK: ContextRef = {
  id: 'pp:deck',
  kind: 'document',
  surface: 'powerpoint',
  title: 'Whole deck',
};

function contextText(ctx: Array<{ value: { as: string; text?: string } }>): string {
  return ctx.map((c) => (c.value.as === 'text' ? (c.value.text ?? '') : '')).join('\n');
}

/* ───────────────────────────── getCapabilities ───────────────────────────── */

describe('PowerPointBridge surface + capabilities', () => {
  it('reports the powerpoint surface and a capability manifest advertising insert-slide', () => {
    const bridge = new PowerPointBridge();
    expect(bridge.surface).toBe('powerpoint');
    const caps = bridge.getCapabilities();
    expect(caps.surface).toBe('powerpoint');
    expect(caps.actuations.map((a) => a.kind)).toContain('insert-slide');
  });
});

/* ───────────────────────────── listContext ───────────────────────────── */

describe('PowerPointBridge.listContext', () => {
  it('degrades to a deck-only ref when getSelectedSlides (1.5) is unsupported', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.4 } });
    const refs = await new PowerPointBridge().listContext();
    expect(refs).toEqual([
      { id: 'pp:deck', kind: 'document', surface: 'powerpoint', title: 'Whole deck' },
    ]);
  });

  it('lists the selected slide (live, 1-based title) plus the whole deck when 1.5 is supported', async () => {
    installed = install(deck(SAMPLE_SLIDES, [1]));
    const refs = await new PowerPointBridge().listContext();
    expect(refs).toHaveLength(4);
    expect(refs[0]).toEqual({
      id: 'pp:slide:s2',
      kind: 'slide',
      surface: 'powerpoint',
      title: 'Slide 2: Roster', // index 1 → human "Slide 2"
      preview: 'Roster',
      live: true,
      anchor: { matchText: 'Roster', locator: 'slide:s2' },
      hostRef: { type: 'powerpoint.slide', slideId: 's2' },
    });
    expect(refs[1]).toMatchObject({
      id: 'pp:shape:s2:s2-shape-0',
      kind: 'shape',
      title: 'Shape 1 on slide 2',
      preview: 'Roster',
    });
    expect(refs[2]).toMatchObject({
      id: 'pp:shape:s2:s2-shape-1',
      kind: 'shape',
      title: 'Shape 2 on slide 2',
      preview: 'Pat, Sam',
    });
    expect(refs[3]).toMatchObject({ id: 'pp:deck', kind: 'document' });
  });

  it('lists only the deck when nothing is selected', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    const refs = await new PowerPointBridge().listContext();
    expect(refs).toEqual([
      { id: 'pp:deck', kind: 'document', surface: 'powerpoint', title: 'Whole deck' },
    ]);
  });
});

/* ───────────────────────────── resolveContext ───────────────────────────── */

describe('PowerPointBridge.resolveContext', () => {
  const slideRef: ContextRef = {
    id: 'pp:slide:s2',
    kind: 'slide',
    surface: 'powerpoint',
    title: 'Slide 2',
  };

  it('resolves a slide ref by host slide id independent of the current selection', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const ctx = await new PowerPointBridge().resolveContext(slideRef);
    expect(ctx.length).toBeGreaterThan(0);
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s2')).toBe(true);
  });

  it('resolves an exact slide ref when no slide is selected', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    const ctx = await new PowerPointBridge().resolveContext(slideRef);
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s2')).toBe(true);
  });

  it('returns [] for a stale slide ref', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    expect(
      await new PowerPointBridge().resolveContext({
        id: 'pp:slide:missing',
        kind: 'slide',
        surface: 'powerpoint',
        title: 'Missing slide',
      }),
    ).toEqual([]);
  });

  it('resolves a shape ref to the exact shape text', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    const ctx = await new PowerPointBridge().resolveContext({
      id: 'pp:shape:s2:s2-shape-1',
      kind: 'shape',
      surface: 'powerpoint',
      title: 'Roster body',
    });
    expect(ctx).toHaveLength(1);
    expect(ctx[0]?.ref.id).toBe('pp:shape:s2:s2-shape-1');
    expect(ctx[0]?.value.as).toBe('text');
    if (ctx[0]?.value.as === 'text') expect(ctx[0].value.text).toBe('Pat, Sam');
  });

  it('treats selection and shape refs the same way as slide refs', async () => {
    installed = install(deck(SAMPLE_SLIDES, [2]));
    const sel = await new PowerPointBridge().resolveContext({
      id: 'x',
      kind: 'selection',
      surface: 'powerpoint',
      title: 'Selection',
    });
    expect(sel.some((c) => c.ref.anchor?.locator === 'slide:s3')).toBe(true);
  });

  it('resolves a document ref by reading every slide in the deck, anchored per slide', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const ctx = await new PowerPointBridge().resolveContext({
      id: 'pp:deck',
      kind: 'document',
      surface: 'powerpoint',
      title: 'Whole deck',
    });
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const locators = ctx.map((c) => c.ref.anchor?.locator);
    expect(locators).toContain('slide:s1');
    expect(locators).toContain('slide:s2');
    expect(locators).toContain('slide:s3');
  });

  it('reads the whole deck in a constant number of syncs, whatever the slide count', async () => {
    const wholeDeck: ContextRef = {
      id: 'pp:deck',
      kind: 'document',
      surface: 'powerpoint',
      title: 'Whole deck',
    };
    const slides = (count: number): SlideSeed[] =>
      Array.from({ length: count }, (_, i) => ({
        id: `s${i}`,
        shapes: [
          { text: `Title ${i}`, zOrderCalls: [] },
          { text: `Body ${i}`, zOrderCalls: [] },
        ],
      }));

    installed = install(deck(slides(3)));
    await new PowerPointBridge().resolveContext(wholeDeck);
    const small = installed.ctxRef.ctx?.syncs;
    installed.restore();

    installed = install(deck(slides(30)));
    const ctx = await new PowerPointBridge().resolveContext(wholeDeck);
    expect(installed.ctxRef.ctx?.syncs).toBe(small);
    expect(small).toBeLessThanOrEqual(3);
    expect(ctx.map((c) => c.value.as === 'text' && c.value.text).join('\n')).toContain('Body 29');
  });

  it('skips shapes without a text frame instead of failing the whole-deck read', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Deck title', zOrderCalls: [] },
            { text: '', type: 'Image', zOrderCalls: [] },
            { text: '', type: 'Table', zOrderCalls: [] },
          ],
        },
        { id: 's2', shapes: [{ text: 'Second slide', zOrderCalls: [] }] },
      ]),
    );
    const ctx = await new PowerPointBridge().resolveContext({
      id: 'pp:deck',
      kind: 'document',
      surface: 'powerpoint',
      title: 'Whole deck',
    });
    const text = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(text).toContain('Deck title');
    expect(text).toContain('Second slide');
  });

  it('skips a placeholder holding a picture without falling back to per-shape reads', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Offsite', type: 'Placeholder', containedType: null, zOrderCalls: [] },
            { text: '', type: 'Placeholder', containedType: 'Image', zOrderCalls: [] },
            { text: 'Lisbon, 8-10 October', type: 'Placeholder', zOrderCalls: [] },
          ],
        },
        { id: 's2', shapes: [{ text: 'Next slide', zOrderCalls: [] }] },
      ]),
    );
    const ctx = await new PowerPointBridge().resolveContext(WHOLE_DECK);
    const text = contextText(ctx);
    expect(text).toContain('Offsite');
    expect(text).toContain('Lisbon, 8-10 October');
    expect(text).toContain('Next slide');
    // slides + shapes + placeholder formats + text: the batch path, no per-shape retries.
    expect(installed.ctxRef.ctx?.syncs).toBe(4);
  });

  it('falls back to per-shape reads when the host refuses an unexpected shape', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Kept title', zOrderCalls: [] },
            { text: 'secret', type: 'GeometricShape', refuseText: true, zOrderCalls: [] },
            { text: 'Kept body', zOrderCalls: [] },
          ],
        },
      ]),
    );
    const text = contextText(await new PowerPointBridge().resolveContext(WHOLE_DECK));
    expect(text).toContain('Kept title');
    expect(text).toContain('Kept body');
    expect(text).not.toContain('secret');
  });

  it('still reads the deck on a host without placeholderFormat (PowerPointApi < 1.8)', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Old host title', type: 'Placeholder', zOrderCalls: [] },
            { text: '', type: 'Placeholder', containedType: 'Image', zOrderCalls: [] },
          ],
        },
      ]),
      { requirements: { PowerPointApi: 1.5 } },
    );
    const text = contextText(await new PowerPointBridge().resolveContext(WHOLE_DECK));
    expect(text).toContain('Old host title');
  });

  it('reads native table cells into the whole deck, in the same round-trip as text', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Wave 3 site budgets', zOrderCalls: [] },
            {
              text: '',
              type: 'Table',
              tableGrid: [
                ['Site', 'Budget'],
                ['Madrid', '$310K'],
                ['', ''],
              ],
              zOrderCalls: [],
            },
          ],
        },
      ]),
    );
    const text = contextText(await new PowerPointBridge().resolveContext(WHOLE_DECK));
    expect(text).toContain('Site | Budget');
    expect(text).toContain('Madrid | $310K');
    // slides + shapes + content (text and tables together): no extra sync for tables.
    expect(installed.ctxRef.ctx?.syncs).toBe(3);
  });

  it('skips tables on a host without Table.values (PowerPointApi < 1.8)', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: 'Budgets', zOrderCalls: [] },
            { text: '', type: 'Table', tableGrid: [['Madrid', '$310K']], zOrderCalls: [] },
          ],
        },
      ]),
      { requirements: { PowerPointApi: 1.5 } },
    );
    const text = contextText(await new PowerPointBridge().resolveContext(WHOLE_DECK));
    expect(text).toContain('Budgets');
    expect(text).not.toContain('Madrid');
  });

  it('resolves a table shape chip to its cells and a picture shape chip to nothing', async () => {
    installed = install(
      deck([
        {
          id: 's1',
          shapes: [
            { text: '', type: 'Image', zOrderCalls: [] },
            {
              text: '',
              type: 'Table',
              tableGrid: [
                ['Aisle', 'Picks / hr'],
                ['C-14', '312'],
              ],
              zOrderCalls: [],
            },
          ],
        },
      ]),
    );
    const shapeRef = (shapeId: string): ContextRef => ({
      id: `pp:shape:s1:${shapeId}`,
      kind: 'shape',
      surface: 'powerpoint',
      title: 'Shape',
      hostRef: { type: 'powerpoint.shape', slideId: 's1', shapeId },
    });
    const bridge = new PowerPointBridge();
    expect(await bridge.resolveContext(shapeRef('s1-shape-0'))).toEqual([]);
    expect(contextText(await bridge.resolveContext(shapeRef('s1-shape-1')))).toContain(
      'C-14 | 312',
    );
  });

  it('bounds the whole-deck read to MAX_READ_SLIDES slides', async () => {
    installed = install(
      deck(
        Array.from({ length: MAX_READ_SLIDES + 5 }, (_, i) => ({
          id: `s${i}`,
          shapes: [{ text: `Title ${i + 1}`, zOrderCalls: [] }],
        })),
      ),
    );
    const ctx = await new PowerPointBridge().resolveContext({
      id: 'pp:deck',
      kind: 'document',
      surface: 'powerpoint',
      title: 'Whole deck',
    });
    const text = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(text).toContain(`Title ${MAX_READ_SLIDES}`);
    expect(text).not.toContain(`Title ${MAX_READ_SLIDES + 1}`);
  });
});

/* ───────────────────────────── captureDocState ───────────────────────────── */

describe('PowerPointBridge.captureDocState', () => {
  it('returns undefined when TextRange.text (1.4) is unsupported', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.3 } });
    expect(await new PowerPointBridge().captureDocState()).toBeUndefined();
  });

  it('returns undefined for an empty deck', async () => {
    installed = install(deck([], []));
    expect(await new PowerPointBridge().captureDocState()).toBeUndefined();
  });

  it('builds a valid snapshot whose inventory comes from the slides and bumps version each capture', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const bridge = new PowerPointBridge();

    const first = await bridge.captureDocState();
    expect(first).toBeDefined();
    if (!first) return;
    expect(() => DocStateSnapshotSchema.parse(first)).not.toThrow();
    expect(first.surface).toBe('powerpoint');
    expect(first.version).toBe(1);
    expect(first.outline.some((o) => o.text.includes('SLA Terms'))).toBe(true);

    const second = await bridge.captureDocState();
    expect(second?.version).toBe(2);
  });

  it('reads at most MAX_READ_SLIDES slides (bounded per-turn read, ADR-0006)', async () => {
    // Give every slide a unique title; the snapshot must not list beyond the cap.
    const many: SlideSeed[] = Array.from({ length: MAX_READ_SLIDES + 10 }, (_, i) => ({
      id: `m${i}`,
      shapes: [{ text: `Title ${i}` }],
    }));
    installed = install(deck(many, [0]));
    const snap = await new PowerPointBridge().captureDocState();
    expect(snap).toBeDefined();
    if (!snap) return;
    // The slide past the cap must not appear in the inventory.
    expect(snap.outline.some((o) => o.text.includes(`Title ${MAX_READ_SLIDES + 5}`))).toBe(false);
    expect(snap.outline.some((o) => o.text.includes('Title 0'))).toBe(true);
  });
});

/* ───────────────────────────── searchDocument ───────────────────────────── */

describe('PowerPointBridge.searchDocument', () => {
  it('returns [] for an empty/whitespace query without touching the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    expect(await new PowerPointBridge().searchDocument('   ')).toEqual([]);
    // No PowerPoint.run was invoked for an empty query.
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('returns [] when the host predates TextRange.text (1.4)', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.3 } });
    expect(await new PowerPointBridge().searchDocument('SLA')).toEqual([]);
  });

  it('returns the slides whose title/body match the query (case-insensitive)', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const ctx = await new PowerPointBridge().searchDocument('sla');
    for (const c of ctx) expect(() => ResolvedContextSchema.parse(c)).not.toThrow();
    const locators = ctx.map((c) => c.ref.anchor?.locator);
    expect(locators).toContain('slide:s1'); // "SLA Terms"
    expect(locators).toContain('slide:s3'); // "SLA gap flagged"
    expect(locators).not.toContain('slide:s2'); // Roster — no SLA
  });

  it('returns [] when nothing matches', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    expect(await new PowerPointBridge().searchDocument('nonexistent-token')).toEqual([]);
  });
});

/* ───────────────────────────── readRange ───────────────────────────── */

describe('PowerPointBridge.readRange', () => {
  it('resolves an addressable slide selector (1-based) to that single slide', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    const ctx = await new PowerPointBridge().readRange('slide:1');
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s1')).toBe(true);
    // Only one slide resolved, not the whole deck.
    expect(new Set(ctx.map((c) => c.ref.anchor?.locator))).toEqual(new Set(['slide:s1']));
  });

  it('accepts a bare 1-based index and maps to the zero-based slide', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    const ctx = await new PowerPointBridge().readRange('3');
    expect(ctx.some((c) => c.ref.anchor?.locator === 'slide:s3')).toBe(true);
  });

  it('returns [] for an unaddressable selector without invoking the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    expect(await new PowerPointBridge().readRange('Agenda')).toEqual([]);
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('returns [] for an out-of-range slide index (degrade, not crash)', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    expect(await new PowerPointBridge().readRange('slide:99')).toEqual([]);
  });

  it('accepts a host slide id and the pp:slide: ref form', async () => {
    installed = install(deck(SAMPLE_SLIDES, []));
    for (const selector of ['s2', 'slide:s2', 'pp:slide:s2', 'pp:slide:2']) {
      const ctx = await new PowerPointBridge().readRange(selector);
      expect(new Set(ctx.map((c) => c.ref.anchor?.locator))).toEqual(new Set(['slide:s2']));
    }
    const last = await new PowerPointBridge().readRange('last');
    expect(last.some((c) => c.ref.anchor?.locator === 'slide:s3')).toBe(true);
  });

  it('lists the slide shapes with ids, types and the title, so shape commands can address them', async () => {
    installed = install(
      deck(
        [
          {
            id: 's1',
            shapes: [
              { text: 'FY26 Plan', type: 'Placeholder', placeholderType: 'Title', zOrderCalls: [] },
              { text: 'Grow ARR', type: 'TextBox', zOrderCalls: [] },
              { text: '', type: 'Image', zOrderCalls: [] },
            ],
          },
        ],
        [],
      ),
    );
    const ctx = await new PowerPointBridge().readRange('slide:1');
    const text = ctx.map((c) => (c.value.as === 'text' ? c.value.text : '')).join('\n');
    expect(text).toContain(
      'for shape commands use slide=1 and shape=<id>; shape=title is id s1-shape-0',
    );
    expect(text).toContain('- id s1-shape-0 · Placeholder · title · "FY26 Plan"');
    expect(text).toContain('- id s1-shape-1 · TextBox · "Grow ARR"');
    expect(text).toContain('- id s1-shape-2 · Image');
  });

  it('returns [] when the host predates TextRange.text (1.4)', async () => {
    installed = install(deck(SAMPLE_SLIDES, []), { requirements: { PowerPointApi: 1.3 } });
    expect(await new PowerPointBridge().readRange('slide:1')).toEqual([]);
  });
});

/* ───────────────────────────── revealContext ───────────────────────────── */

describe('PowerPointBridge.revealContext', () => {
  it('selects an addressable slide ref by host slide id', async () => {
    const seed = deck(SAMPLE_SLIDES, [0]);
    installed = install(seed);
    const bridge = new PowerPointBridge();
    const ref: ContextRef = {
      id: 'pp:slide:s3',
      kind: 'slide',
      surface: 'powerpoint',
      title: 'Slide 3',
    };

    expect(bridge.canRevealContext(ref)).toBe(true);
    await bridge.revealContext(ref);

    expect(seed.selectedIndices).toEqual([2]);
  });

  it('selects a shape when the ref carries slide and shape ids', async () => {
    const seed = deck(SAMPLE_SLIDES, [0]);
    installed = install(seed);
    await new PowerPointBridge().revealContext({
      id: 'pp:shape:s2:s2-shape-1',
      kind: 'shape',
      surface: 'powerpoint',
      title: 'Body text',
    });

    expect(seed.selectedIndices).toEqual([1]);
    expect(seed.selectedShapeIds).toEqual(['s2-shape-1']);
  });

  it('does not advertise deck-wide refs as revealable', () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    expect(
      new PowerPointBridge().canRevealContext({
        id: 'pp:deck',
        kind: 'document',
        surface: 'powerpoint',
        title: 'Whole deck',
      }),
    ).toBe(false);
  });
});

/* ───────────────────────────── actuate: dispatch + insert-slide ───────────────────────────── */

describe('PowerPointBridge.actuate dispatch', () => {
  it('returns an unsupported error for a kind PowerPoint cannot handle, echoing the changeId', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('chg-x'),
      kind: 'write-cells',
      surface: 'powerpoint',
      params: { cells: [['1']] },
    });
    expect(res).toMatchObject({
      ok: false,
      changeId: asChangeId('chg-x'),
      kind: 'write-cells',
      error: { code: 'unsupported' },
    });
    // No host write happened for an unsupported kind.
    expect(installed.ctxRef.ctx).toBeUndefined();
  });
});

describe('PowerPointBridge.actuate insert-slide (empty)', () => {
  it('rejects an insert with no base64, no title, and no bullets before touching the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: '', bullets: [] } }),
    );
    expect(res).toMatchObject({
      ok: false,
      kind: 'insert-slide',
      error: { code: 'empty_slide' },
    });
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('rejects an insert whose only content is whitespace notes (notes do not satisfy the gate)', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: '', bullets: [], notes: '   ' } }),
    );
    expect(res.ok).toBe(false);
    expect(res.error?.code).toBe('empty_slide');
  });
});

describe('PowerPointBridge.actuate insert-slide (base64 prebuilt deck)', () => {
  it('merges a prebuilt base64 deck via insertSlidesFromBase64 and reports inserted-deck', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ ooxml: 'UEsDBBQ=' }, 'chg-deck'),
    );
    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('chg-deck'),
      kind: 'insert-slide',
      location: 'inserted-deck',
    });
    // The base64 payload reached the host merge path exactly once.
    expect(d.insertedDecks).toEqual(['UEsDBBQ=']);
    expect(d.insertedDeckOptions).toEqual([{ targetSlideId: 's3' }]);
  });

  it('imports an explicit generated deck artifact in one call with formatting and slide count', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide(
        {
          deck: {
            format: 'pptx',
            base64: 'compiled-deck',
            slideCount: 2,
            formatting: 'UseDestinationTheme',
            targetSlideId: 's1',
            specFingerprint: '9f34a012',
          },
        },
        'chg-generated-deck',
      ),
    );
    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('chg-generated-deck'),
      kind: 'insert-slide',
      location: 'inserted-deck:2',
    });
    expect(d.insertedDecks).toEqual(['compiled-deck']);
    expect(d.insertedDeckOptions).toEqual([
      { formatting: 'UseDestinationTheme', targetSlideId: 's1' },
    ]);
  });

  it('prefers the base64 path over native compose when both are present (no slides.add)', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const before = d.slides.length;
    const res = await new PowerPointBridge().actuate(
      insertSlide({ ooxml: 'UEs=', slide: { title: 'Ignored', bullets: ['x'] } }),
    );
    expect(res.location).toBe('inserted-deck');
    expect(d.insertedDecks).toEqual(['UEs=']);
    // Native compose would have seeded a 2-shape placeholder slide; base64 seeds a 0-shape marker.
    const appended = d.slides[d.slides.length - 1];
    expect(d.slides.length).toBe(before + 1);
    expect(appended?.shapes).toHaveLength(0);
  });
});

describe('PowerPointBridge.actuate insert-slide (native compose)', () => {
  it('appends a slide and writes the title into shape[0] and bullets into shape[1]', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const before = d.slides.length;
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'New Topic', bullets: ['Point A', 'Point B'] } }, 'chg-native'),
    );
    expect(res.ok).toBe(true);
    expect(res.changeId).toBe(asChangeId('chg-native'));
    expect(res.location).toBe('slide:sim-slide-4'); // appended at index `before`
    // Exactly one slide was appended.
    expect(d.slides.length).toBe(before + 1);
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes[0]?.text).toBe('New Topic');
    expect(appended?.shapes[1]?.text).toBe('Point A\nPoint B');
  });

  it('writes a title-only slide without touching the body placeholder', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    installed = install(d);
    await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Just A Title', bullets: [] } }),
    );
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes[0]?.text).toBe('Just A Title');
    // Body placeholder stays empty when there are no bullets.
    expect(appended?.shapes[1]?.text).toBe('');
  });

  it('writes bullets even when the title is empty (body-only compose)', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: '', bullets: ['Only a bullet'] } }),
    );
    expect(res.ok).toBe(true);
    const appended = d.slides[d.slides.length - 1];
    // Title placeholder untouched (empty); bullets written to the body placeholder.
    expect(appended?.shapes[0]?.text).toBe('');
    expect(appended?.shapes[1]?.text).toBe('Only a bullet');
  });
});

describe('PowerPointBridge.actuate insert-slide (layout without placeholders)', () => {
  it('matches placeholders by type, not position', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [
      { text: '', type: 'Placeholder', placeholderType: 'Body' },
      { text: '', type: 'Placeholder', placeholderType: 'Title' },
    ];
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Typed', bullets: ['One', 'Two'] } }),
    );
    expect(res.ok).toBe(true);
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes[1]?.text).toBe('Typed');
    expect(appended?.shapes[0]?.text).toBe('One\nTwo');
    expect(d.addedShapes).toEqual([]);
  });

  it('adds sized text boxes when the default layout has no placeholders (Blank layout)', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [];
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'this is a test slide', bullets: ['First', 'Second'] } }),
    );
    expect(res.ok).toBe(true);
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes.map((shape) => shape.text)).toEqual([
      'this is a test slide',
      '• First\n• Second',
    ]);
    // Both boxes sit inside the 720 × 405 pt slide the host reported.
    for (const added of d.addedShapes) {
      const box = added.options as PowerPoint.ShapeAddOptions;
      expect(added.kind).toBe('textBox');
      expect((box.left ?? 0) + (box.width ?? 0)).toBeLessThanOrEqual(720);
      expect((box.top ?? 0) + (box.height ?? 0)).toBeLessThanOrEqual(405);
    }
    expect(appended?.shapes[0]?.font?.bold).toBe(true);
  });

  it('adds only a body text box when the layout has a title placeholder but no body', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [{ text: '', type: 'Placeholder', placeholderType: 'Title' }];
    installed = install(d);
    await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Heading', bullets: ['Point'] } }),
    );
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes.map((shape) => shape.text)).toEqual(['Heading', '• Point']);
    expect(d.addedShapes.map((added) => added.kind)).toEqual(['textBox']);
  });

  it('falls back to the first two placeholders in order below PowerPointApi 1.8', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [
      { text: '', type: 'Placeholder' },
      { text: '', type: 'Placeholder' },
    ];
    installed = install(d, { requirements: { PowerPointApi: 1.5 } });
    await new PowerPointBridge().actuate(insertSlide({ slide: { title: 'T', bullets: ['B'] } }));
    const appended = d.slides[d.slides.length - 1];
    expect(appended?.shapes.map((shape) => shape.text)).toEqual(['T', 'B']);
  });

  it('writes to the just-added slide by position and reports its settled id (PowerPoint web)', async () => {
    // The fake gives a new slide a provisional id until the request ends, and fails writes addressed
    // through that id — as PowerPoint for the web does (5010 "InvalidParam passed to GetItem(id)").
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [];
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'Q4 plan', bullets: ['hire 5 engineers'] } }),
    );
    expect(res).toMatchObject({ ok: true, location: 'slide:sim-slide-2' });
    expect(d.slides[1]?.shapes.map((shape) => shape.text)).toEqual([
      'Q4 plan',
      '• hire 5 engineers',
    ]);
  });

  it('includes the PowerPoint error code when a write is uncertain', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [
      { text: '', type: 'Placeholder', placeholderType: 'Title', refuseText: true },
    ];
    installed = install(d);
    const sync = vi.spyOn(FakeContext.prototype, 'sync').mockImplementation(function (
      this: FakeContext,
    ) {
      this.syncs += 1;
      const error = pendingHostError;
      pendingHostError = undefined;
      return error
        ? Promise.reject(Object.assign(new Error(error), { code: 'GeneralException' }))
        : Promise.resolve();
    });
    try {
      const res = await new PowerPointBridge().actuate(
        insertSlide({ slide: { title: 'T', bullets: [] } }),
      );
      expect(res.error?.code).toBe('outcome_unknown');
      expect(res.error?.message).toContain('(PowerPoint: GeneralException)');
    } finally {
      sync.mockRestore();
    }
  });

  it('never writes onto another slide when the new one cannot be identified (co-author race)', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    installed = install(d);
    // A co-author's slide lands in the same moment: two new ids appear, so neither is trusted.
    const add = vi.spyOn(FakeSlideCollection.prototype, 'add').mockImplementation(() => {
      d.slides.push({ id: 'coauthor-slide', shapes: [{ text: 'Theirs' }] });
      d.slides.push({ id: 'ours', shapes: [] });
    });
    try {
      const res = await new PowerPointBridge().actuate(
        insertSlide({ slide: { title: 'T', bullets: ['B'] } }),
      );
      expect(res).toMatchObject({ ok: false, error: { code: 'outcome_unknown' } });
      expect(d.slides.find((slide) => slide.id === 'coauthor-slide')?.shapes[0]?.text).toBe(
        'Theirs',
      );
    } finally {
      add.mockRestore();
    }
  });

  it('reports an unconfirmed write as outcome-unknown, never as success', async () => {
    const d = deck([{ id: 's1', shapes: [{ text: 'Existing' }] }], [0]);
    d.newSlideShapes = [
      { text: '', type: 'Placeholder', placeholderType: 'Title', refuseText: true },
    ];
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      insertSlide({ slide: { title: 'T', bullets: [] } }),
    );
    expect(res.ok).toBe(false);
    expect(res.error?.code).toBe('outcome_unknown');
  });
});

describe('PowerPointBridge shape and slide references (tests 48–51)', () => {
  it('changes only the title when a shape command names slide 1 / shape title', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(setShapeText('1', 'title', 'FY26 Plan'));
    expect(res).toMatchObject({ ok: true, location: 'shape:s1:s1-shape-0' });
    expect(d.slides[0]?.shapes.map((shape) => shape.text)).toEqual([
      'FY26 Plan',
      '99.5% contracted\nMonthly window',
    ]);
  });

  it('formats the title shape resolved by slide number and "title"', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: '1', shapeId: 'title' },
        shapeFormat: { fill: '#0000FF', font: { color: '#FFFFFF' } },
      }),
    );
    expect(res.ok).toBe(true);
    expect(d.slides[0]?.shapes[0]).toMatchObject({
      fillColor: '#0000FF',
      font: { color: '#FFFFFF' },
    });
    expect(d.slides[0]?.shapes[1]?.fillColor).toBeUndefined();
  });

  it('names the valid shapes when a shape reference does not resolve, and writes nothing', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(setShapeText('1', 'nope', 'x'));
    expect(res).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(res.error?.message).toContain('Use shape=title');
    expect(res.error?.message).toContain('id s1-shape-0 (TextBox)');
    expect(d.slides[0]?.shapes[0]?.text).toBe('SLA Terms');
  });

  it('rejects a format-shape with no formatting instead of reporting a no-op as applied', async () => {
    // Live: `/format-shape shape:257#0:6 text="…"` came back "applied" having changed nothing.
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      formatShape({ target: { slideId: '1', shapeId: 'title' }, shapeFormat: {} }),
    );
    expect(res).toMatchObject({ ok: false, error: { code: 'no_format' } });
    expect(res.error?.message).toContain('shape pp:shape:<slide>:<shape> "text"');
  });

  it('does not treat a shape number as a position (ids are small numbers too)', async () => {
    // Live: `shape pp:shape:1:1 "FY26 Plan"` resolved "1" as the 1st shape — the logo image.
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(setShapeText('1', '1', 'FY26 Plan'));
    expect(res).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(res.error?.message).toContain('Use shape=title');
    expect(d.slides[0]?.shapes.map((shape) => shape.text)).toEqual([
      'SLA Terms',
      '99.5% contracted\nMonthly window',
    ]);
  });

  it('refuses to write text into a picture, with a message naming shape=title', async () => {
    const d = deck(
      [{ id: 's1', shapes: [{ text: '', type: 'Image' }, { text: 'Deck title' }] }],
      [0],
    );
    installed = install(d);
    const res = await new PowerPointBridge().actuate(setShapeText('1', 's1-shape-0', 'x'));
    expect(res).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(res.error?.message).toContain('is a Image and has no text');
    expect(d.slides[0]?.shapes[1]?.text).toBe('Deck title');
  });

  it('rejects title= on an existing slide instead of silently ignoring it', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('table-title-existing'),
      kind: 'add-table-slide',
      surface: 'powerpoint',
      params: {
        target: { slideId: '2' },
        slide: { title: 'Ignored?', bullets: [] },
        tableGrid: { hasHeaders: false, rows: [['a']] },
      },
    });
    expect(res).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(res.error?.message).toContain('only used with slide=new');
    expect(d.addedShapes).toEqual([]);
  });

  it('rejects a text box with no text, and a shape with no position, before writing', async () => {
    // Live: `/add-shape slide:2 {"text": …, "position": "bottom"}` → an empty box at the top-left.
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const bridge = new PowerPointBridge();
    const empty = await bridge.actuate(
      addShape({
        target: { slideId: '2' },
        shape: { shapeType: 'textBox', left: 72, top: 300, width: 300, height: 40 },
      }),
    );
    expect(empty).toMatchObject({ ok: false, error: { code: 'no_text' } });
    const unplaced = await bridge.actuate(
      addShape({ target: { slideId: '2' }, shape: { shapeType: 'textBox', text: 'Draft' } }),
    );
    expect(unplaced).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(unplaced.error?.message).toContain('missing left, top, width, height');
    expect(unplaced.error?.message).toContain('The slide is 720 × 405 pt');
    expect(d.addedShapes).toEqual([]);
  });

  it('accepts pp:slide:<id> and rejects an off-slide position with the slide size', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const bridge = new PowerPointBridge();
    const ok = await bridge.actuate(
      addShape({
        target: { slideId: 'pp:slide:s2' },
        shape: { shapeType: 'textBox', text: 'Draft', left: 72, top: 330, width: 400, height: 40 },
      }),
    );
    expect(ok.ok).toBe(true);
    expect(d.addedShapes[0]?.slideId).toBe('s2');
    const off = await bridge.actuate(
      addShape({
        target: { slideId: '2' },
        shape: { shapeType: 'textBox', text: 'Draft', left: 100, top: 500, width: 300, height: 50 },
      }),
    );
    expect(off).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(off.error?.message).toContain('outside the slide, which is 720 × 405 pt');
    expect(d.addedShapes).toHaveLength(1);
  });

  it('creates a titled slide for /add-table-slide slide=new and puts the table on it', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('table-new'),
      kind: 'add-table-slide',
      surface: 'powerpoint',
      params: {
        target: { slideId: 'new' },
        slide: { title: 'Key risks', bullets: [] },
        tableGrid: {
          hasHeaders: true,
          rows: [
            ['Risk', 'Owner'],
            ['Vendor delay', 'Pat'],
          ],
        },
      },
    });
    expect(res).toMatchObject({ ok: true });
    expect(res.location).toMatch(/^shape:sim-slide-4:/);
    // With no position given, the table goes below the title, inside the 720 × 405 slide.
    const tableAdd = d.addedShapes.find((added) => added.kind === 'table');
    expect(tableAdd?.options).toMatchObject({ left: 48, top: 109, width: 624 });
    // Undo removes the slide this command created (and the table with it).
    expect(res.inverse).toEqual({ op: 'delete-object', objectType: 'slide', name: 'sim-slide-4' });
    const created = d.slides[3];
    expect(created?.shapes[0]?.text).toBe('Key risks');
    expect(created?.shapes.find((shape) => shape.tableGrid)?.tableGrid).toEqual([
      ['Risk', 'Owner'],
      ['Vendor delay', 'Pat'],
    ]);
    // The existing last slide was not touched.
    expect(d.slides[2]?.shapes.map((shape) => shape.text)).toEqual([
      'Risk Summary',
      'SLA gap flagged',
    ]);
  });

  it('adds a title text box for slide=new on a layout without placeholders', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    d.newSlideShapes = [];
    installed = install(d);
    const res = await new PowerPointBridge().actuate({
      changeId: asChangeId('table-new-blank'),
      kind: 'add-table-slide',
      surface: 'powerpoint',
      params: {
        target: { slideId: 'new' },
        slide: { title: 'Key risks', bullets: [] },
        tableGrid: { hasHeaders: false, rows: [['Vendor delay', 'Pat']] },
      },
    });
    expect(res.ok).toBe(true);
    expect(d.slides[3]?.shapes.map((shape) => shape.type ?? 'TextBox')).toEqual([
      'TextBox',
      'Table',
    ]);
    expect(d.slides[3]?.shapes[0]?.text).toBe('Key risks');
  });
});

describe('PowerPointBridge.actuate set-shape-text', () => {
  // Syncs: 1 resolve slide, 2 resolve shape, 3 shape exists?, 4 prior text read, 5 the write.
  it.each([1, 2, 3, 4, 5])(
    'distinguishes a failed pre-read from a dispatched text write (sync %s)',
    async (failAt) => {
      const d = deck(SAMPLE_SLIDES, [0]);
      installed = install(d);
      const original = FakeContext.prototype.sync;
      let calls = 0;
      const sync = vi.spyOn(FakeContext.prototype, 'sync').mockImplementation(function (
        this: FakeContext,
      ) {
        if (++calls === failAt) return Promise.reject(new Error('Office disconnected'));
        return original.call(this);
      });
      try {
        const result = await new PowerPointBridge().actuate(
          setShapeText('s2', 's2-shape-1', 'New text'),
        );
        const beforeWrite = failAt < 5;
        expect(result).toMatchObject({
          ok: false,
          error: { code: beforeWrite ? 'target_conflict' : 'outcome_unknown' },
        });
        expect(result.recoveryPending).toBe(beforeWrite ? undefined : true);
        expect(d.slides[1]?.shapes[1]?.text).toBe(beforeWrite ? 'Pat, Sam' : 'New text');
      } finally {
        sync.mockRestore();
      }
    },
  );

  it('rewrites one addressed shape and records the inverse text', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      setShapeText('s2', 's2-shape-1', 'Pat, Sam, and Lee'),
    );

    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('shape-1'),
      kind: 'set-shape-text',
      location: 'shape:s2:s2-shape-1',
      inverse: {
        op: 'restore-text',
        anchor: 'pp:shape:s2:s2-shape-1',
        priorText: 'Pat, Sam',
      },
    });
    expect(d.slides[1]?.shapes[1]?.text).toBe('Pat, Sam, and Lee');
  });

  it('fails closed when the addressed shape no longer exists', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      setShapeText('s2', 's2-shape-99', 'Should not apply'),
    );

    expect(res).toMatchObject({
      ok: false,
      kind: 'set-shape-text',
      degraded: true,
      error: { code: 'target_conflict' },
    });
    expect(d.slides[1]?.shapes[1]?.text).toBe('Pat, Sam');
  });

  it('fails before touching the host when the required shape text API is unavailable', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.3 } });
    const res = await new PowerPointBridge().actuate(
      setShapeText('s2', 's2-shape-1', 'Unsupported'),
    );

    expect(res).toMatchObject({
      ok: false,
      kind: 'set-shape-text',
      error: { code: 'unsupported' },
    });
    expect(installed.ctxRef.ctx).toBeUndefined();
  });
});

/* ───────────────────────────── actuate: add-shape ───────────────────────────── */

function addShape(params: ActuationRequest['params'], id = 'add-shape-1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'add-shape', surface: 'powerpoint', params };
}

describe('PowerPointBridge.actuate add-shape', () => {
  it('reports unknown outcome when shape creation landed but its sync failed', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const sync = vi
      .spyOn(FakeContext.prototype, 'sync')
      // The read-only slide lookup succeeds; the sync that carries the write fails.
      .mockResolvedValueOnce(undefined)
      .mockRejectedValue(new Error('Office disconnected'));
    try {
      const result = await new PowerPointBridge().actuate(
        addShape({
          target: { slideId: 's1' },
          shape: {
            shapeType: 'textBox',
            text: 'Created',
            left: 72,
            top: 300,
            width: 300,
            height: 40,
          },
        }),
      );
      expect(result).toMatchObject({
        ok: false,
        recoveryPending: true,
        error: { code: 'outcome_unknown' },
      });
      expect(d.slides[0]?.shapes).toHaveLength(3);
    } finally {
      sync.mockRestore();
    }
  });

  it('adds a text box to the addressed slide, recording the minted shape id as delete-object inverse', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape(
        {
          target: { slideId: 's2' },
          shape: {
            shapeType: 'textBox',
            text: 'Key risk',
            left: 72,
            top: 96,
            width: 280,
            height: 80,
          },
        },
        'chg-add-textbox',
      ),
    );

    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('chg-add-textbox'),
      kind: 'add-shape',
      location: 'shape:s2:s2-shape-2',
      inverse: { op: 'delete-object', objectType: 'shape', name: 's2-shape-2' },
    });
    // The text box landed on slide s2 with its text and explicit geometry.
    expect(d.addedShapes).toEqual([
      {
        slideId: 's2',
        kind: 'textBox',
        detail: 'Key risk',
        options: { left: 72, top: 96, width: 280, height: 80 },
      },
    ]);
    expect(d.slides[1]?.shapes).toHaveLength(3);
    expect(d.slides[1]?.shapes[2]?.text).toBe('Key risk');
  });

  it('adds a filled geometric shape with a whitelisted geometry literal', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape({
        target: { slideId: 's1' },
        shape: {
          shapeType: 'geometric',
          geometryType: 'Rectangle',
          fill: '#0F6CBD',
          left: 72,
          top: 96,
          width: 120,
          height: 60,
        },
      }),
    );

    expect(res.ok).toBe(true);
    expect(res.location).toBe('shape:s1:s1-shape-2');
    expect(d.addedShapes[0]).toMatchObject({ kind: 'geometric', detail: 'Rectangle' });
    expect(d.slides[0]?.shapes[2]?.fillColor).toBe('#0F6CBD');
  });

  it('maps the contract connector name onto the host ConnectorType for lines', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape({
        target: { slideId: 's3' },
        shape: {
          shapeType: 'line',
          connectorType: 'elbow',
          left: 72,
          top: 300,
          width: 300,
          height: 40,
        },
      }),
    );

    expect(res.ok).toBe(true);
    expect(d.addedShapes).toHaveLength(1);
    expect(d.addedShapes[0]).toMatchObject({ kind: 'line', detail: 'Elbow' });
  });

  it('fails closed without touching the host when the target or shape params are missing', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const bridge = new PowerPointBridge();

    const noSlide = await bridge.actuate(addShape({ shape: { shapeType: 'textBox', text: 'x' } }));
    expect(noSlide).toMatchObject({ ok: false, error: { code: 'no_target' } });

    const noShape = await bridge.actuate(addShape({ target: { slideId: 's2' } }));
    expect(noShape).toMatchObject({ ok: false, error: { code: 'no_shape' } });

    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('rejects a non-whitelisted geometric geometry as unsupported before touching the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const res = await new PowerPointBridge().actuate(
      addShape({
        target: { slideId: 's2' },
        shape: { shapeType: 'geometric', geometryType: 'PortalGun' },
      }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it.each([
    ['last', 's3'],
    ['LAST', 's3'],
    ['2', 's2'],
    ['slide:s1', 's1'],
  ])('resolves slide=%s to the right slide before writing', async (ref, slideId) => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape({
        target: { slideId: ref },
        shape: { shapeType: 'textBox', text: 'x', left: 72, top: 300, width: 300, height: 40 },
      }),
    );
    expect(res.ok).toBe(true);
    expect(d.addedShapes[0]?.slideId).toBe(slideId);
    expect(res.location).toMatch(new RegExp(`^shape:${slideId}:`));
  });

  it('points a missing-slide add-shape / add-table-slide at the slide command', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const bridge = new PowerPointBridge();
    const table = await bridge.actuate(
      addTableSlide({ tableGrid: { hasHeaders: false, rows: [['hire 5 engineers']] } }),
    );
    const shape = await bridge.actuate(addShape({ shape: { shapeType: 'textBox', text: 'x' } }));
    for (const res of [table, shape]) {
      expect(res).toMatchObject({ ok: false, error: { code: 'no_target' } });
      expect(res.error?.message).toContain('slide "Title" "bullet"');
      expect(res.error?.message).toContain('a slide id from the outline, a slide number, or last');
      // The runtime escapes < > in results; keep the hint free of them so it reads cleanly.
      expect(res.error?.message).not.toMatch(/[<>]/);
    }
    expect(table.error?.message).toContain('slide=new title=');
  });

  it('fails with the valid slides, not "outcome uncertain", when the slide does not exist', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addShape({ target: { slideId: 'nope' }, shape: { shapeType: 'textBox', text: 'x' } }),
    );
    expect(res).toMatchObject({ ok: false, degraded: true, error: { code: 'target_conflict' } });
    expect(res.error?.message).toContain('Slide "nope" was not found');
    expect(res.error?.message).toContain('or "last"');
    expect(res.error?.message).toContain('s1, s2, s3');
    expect(d.addedShapes).toEqual([]);
  });

  it('degrades to unsupported on a pre-1.4 host without touching the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.3 } });
    const res = await new PowerPointBridge().actuate(
      addShape({ target: { slideId: 's2' }, shape: { shapeType: 'textBox', text: 'x' } }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(res.error?.message).toContain('PowerPointApi 1.4');
    expect(installed.ctxRef.ctx).toBeUndefined();
  });
});

/* ───────────────────────────── actuate: format-shape ───────────────────────────── */

function formatShape(params: ActuationRequest['params'], id = 'format-shape-1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'format-shape', surface: 'powerpoint', params };
}

function styledDeck(): DeckSeed {
  return deck([
    {
      id: 's2',
      shapes: [
        { text: 'Roster', zOrderCalls: [] },
        {
          text: 'Pat, Sam',
          zOrderCalls: [],
          fillColor: '#FFFFFF',
          lineColor: '#000000',
          font: {
            bold: false,
            italic: false,
            underline: 'None',
            color: '#111111',
            size: 18,
            name: 'Arial',
          },
        },
      ],
    },
  ]);
}

describe('PowerPointBridge.actuate format-shape', () => {
  it('retains uncertainty when one format facet landed before the next prior-state read failed', async () => {
    const d = styledDeck();
    installed = install(d);
    const original = FakeContext.prototype.sync;
    let calls = 0;
    const sync = vi.spyOn(FakeContext.prototype, 'sync').mockImplementation(function (
      this: FakeContext,
    ) {
      // Syncs 1–2 resolve the slide and shape; 3 exists?, 4 fill prior read, 5 line prior read
      // (which also carries the queued fill write).
      if (++calls === 5) return Promise.reject(new Error('Office disconnected'));
      return original.call(this);
    });
    try {
      const result = await new PowerPointBridge().actuate(
        formatShape({
          target: { slideId: 's2', shapeId: 's2-shape-1' },
          shapeFormat: { fill: '#123456', line: '#654321' },
        }),
      );
      expect(result).toMatchObject({
        ok: false,
        recoveryPending: true,
        error: { code: 'outcome_unknown' },
      });
      expect(d.slides[0]?.shapes[1]?.fillColor).toBe('#123456');
      expect(d.slides[0]?.shapes[1]?.lineColor).toBe('#000000');
    } finally {
      sync.mockRestore();
    }
  });

  it('applies fill/line/font deltas and records only the touched prior fields as inverse', async () => {
    const d = styledDeck();
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      formatShape(
        {
          target: { slideId: 's2', shapeId: 's2-shape-1' },
          shapeFormat: {
            fill: '#0F6CBD',
            line: '#FF0000',
            font: { bold: true, italic: true, color: '#222222', size: 24, name: 'Verdana' },
          },
        },
        'chg-format',
      ),
    );

    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('chg-format'),
      kind: 'format-shape',
      location: 'shape:s2:s2-shape-1',
      inverse: {
        op: 'restore-shape-format',
        shapeId: 's2-shape-1',
        prior: {
          fill: '#FFFFFF',
          line: '#000000',
          'font.bold': 'false',
          'font.italic': 'false',
          'font.color': '#111111',
          'font.size': '18',
          'font.name': 'Arial',
        },
      },
    });
    const shape = d.slides[0]?.shapes[1];
    expect(shape?.fillColor).toBe('#0F6CBD');
    expect(shape?.lineColor).toBe('#FF0000');
    expect(shape?.font).toMatchObject({
      bold: true,
      italic: true,
      color: '#222222',
      size: 24,
      name: 'Verdana',
      underline: 'None', // untouched
    });
  });

  it('maps the boolean underline onto the host Single/None style literals', async () => {
    const d = styledDeck();
    installed = install(d);
    await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: 's2', shapeId: 's2-shape-1' },
        shapeFormat: { font: { underline: true } },
      }),
    );
    expect(d.slides[0]?.shapes[1]?.font?.underline).toBe('Single');
  });

  it('moves the shape in z-order on a 1.8 host', async () => {
    const d = styledDeck();
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: 's2', shapeId: 's2-shape-1' },
        shapeFormat: { zOrder: 'front' },
      }),
    );

    expect(res.ok).toBe(true);
    expect(res.inverse).toMatchObject({ op: 'restore-shape-format', prior: {} });
    expect(d.slides[0]?.shapes[1]?.zOrderCalls).toEqual(['BringToFront']);
  });

  it('rejects z-order on a pre-1.8 host without touching the host', async () => {
    installed = install(styledDeck(), { requirements: { PowerPointApi: 1.7 } });
    const res = await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: 's2', shapeId: 's2-shape-1' },
        shapeFormat: { zOrder: 'back' },
      }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(res.error?.message).toContain('PowerPointApi 1.8');
    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('fails closed when the addressed shape or formatting params are missing', async () => {
    const d = styledDeck();
    installed = install(d);
    const bridge = new PowerPointBridge();

    const noTarget = await bridge.actuate(formatShape({ shapeFormat: { fill: '#000000' } }));
    expect(noTarget).toMatchObject({ ok: false, error: { code: 'no_target' } });

    const noFormat = await bridge.actuate(
      formatShape({ target: { slideId: 's2', shapeId: 's2-shape-1' } }),
    );
    expect(noFormat).toMatchObject({ ok: false, error: { code: 'no_format' } });

    // Both param failures rejected the request before any host object was touched.
    expect(installed.ctxRef.ctx).toBeUndefined();

    // A stale shape id degrades instead of formatting some other shape.
    const stale = await bridge.actuate(
      formatShape({
        target: { slideId: 's2', shapeId: 's2-shape-99' },
        shapeFormat: { fill: '#000000' },
      }),
    );
    expect(stale).toMatchObject({
      ok: false,
      degraded: true,
      error: { code: 'target_conflict' },
    });
    expect(d.slides[0]?.shapes[1]?.fillColor).toBe('#FFFFFF');
  });

  it('degrades to unsupported on a pre-1.4 host without touching the host', async () => {
    installed = install(styledDeck(), { requirements: { PowerPointApi: 1.3 } });
    const res = await new PowerPointBridge().actuate(
      formatShape({
        target: { slideId: 's2', shapeId: 's2-shape-1' },
        shapeFormat: { fill: '#000000' },
      }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(res.error?.message).toContain('PowerPointApi 1.4');
    expect(installed.ctxRef.ctx).toBeUndefined();
  });
});

/* ───────────────────────────── actuate: add-table-slide ───────────────────────────── */

function addTableSlide(params: ActuationRequest['params'], id = 'add-table-1'): ActuationRequest {
  return { changeId: asChangeId(id), kind: 'add-table-slide', surface: 'powerpoint', params };
}

describe('PowerPointBridge.actuate add-table-slide', () => {
  it('reports uncertainty when a table landed but its identifying receipt could not be read', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const sync = vi
      .spyOn(FakeContext.prototype, 'sync')
      // The read-only slide-size and slide lookups succeed; the sync that carries the write fails.
      .mockResolvedValueOnce(undefined)
      .mockResolvedValueOnce(undefined)
      .mockRejectedValue(new Error('Office disconnected'));
    try {
      const result = await new PowerPointBridge().actuate(
        addTableSlide({
          target: { slideId: 's1' },
          tableGrid: { hasHeaders: false, rows: [['Created']] },
        }),
      );
      expect(result).toMatchObject({
        ok: false,
        recoveryPending: true,
        error: { code: 'outcome_unknown' },
      });
      expect(d.slides[0]?.shapes[2]?.tableGrid).toEqual([['Created']]);
    } finally {
      sync.mockRestore();
    }
  });

  it('seeds a native table from the value grid and records the minted shape id as inverse', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addTableSlide(
        {
          target: { slideId: 's3' },
          tableGrid: {
            hasHeaders: true,
            rows: [
              ['Metric', 'Value'],
              ['ARR', '$12M'],
            ],
            left: 72,
            top: 120,
            width: 560,
            height: 260,
          },
        },
        'chg-table',
      ),
    );

    expect(res).toEqual({
      ok: true,
      provenanceMissing: true,
      changeId: asChangeId('chg-table'),
      kind: 'add-table-slide',
      location: 'shape:s3:s3-shape-2',
      inverse: { op: 'delete-object', objectType: 'shape', name: 's3-shape-2' },
    });
    expect(d.addedShapes).toEqual([
      {
        slideId: 's3',
        kind: 'table',
        detail: '2x2',
        options: { left: 72, top: 120, width: 560, height: 260 },
      },
    ]);
    expect(d.slides[2]?.shapes[2]?.tableGrid).toEqual([
      ['Metric', 'Value'],
      ['ARR', '$12M'],
    ]);
  });

  it('pads ragged rows out to the widest row length', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const res = await new PowerPointBridge().actuate(
      addTableSlide({
        target: { slideId: 's1' },
        tableGrid: { hasHeaders: false, rows: [['Only'], ['B', 'C']] },
      }),
    );

    expect(res.ok).toBe(true);
    expect(res.location).toBe('shape:s1:s1-shape-2');
    expect(d.addedShapes[0]).toMatchObject({ kind: 'table', detail: '2x2' });
    expect(d.slides[0]?.shapes[2]?.tableGrid).toEqual([
      ['Only', ''],
      ['B', 'C'],
    ]);
  });

  it('fails closed without touching the host when the target or grid is missing/empty', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const bridge = new PowerPointBridge();

    const noSlide = await bridge.actuate(
      addTableSlide({ tableGrid: { hasHeaders: false, rows: [['a']] } }),
    );
    expect(noSlide).toMatchObject({ ok: false, error: { code: 'no_target' } });

    const noGrid = await bridge.actuate(addTableSlide({ target: { slideId: 's1' } }));
    expect(noGrid).toMatchObject({ ok: false, error: { code: 'no_table' } });

    const emptyGrid = await bridge.actuate(
      addTableSlide({ target: { slideId: 's1' }, tableGrid: { hasHeaders: false, rows: [] } }),
    );
    expect(emptyGrid).toMatchObject({ ok: false, error: { code: 'no_table' } });

    const emptyRow = await bridge.actuate(
      addTableSlide({ target: { slideId: 's1' }, tableGrid: { hasHeaders: false, rows: [[]] } }),
    );
    expect(emptyRow).toMatchObject({ ok: false, error: { code: 'no_table' } });

    expect(installed.ctxRef.ctx).toBeUndefined();
  });

  it('resolves slide=last and rejects an unknown slide before writing', async () => {
    const d = deck(SAMPLE_SLIDES, [0]);
    installed = install(d);
    const ok = await new PowerPointBridge().actuate(
      addTableSlide({
        target: { slideId: 'last' },
        tableGrid: { hasHeaders: false, rows: [['a']] },
      }),
    );
    expect(ok.ok).toBe(true);
    expect(d.addedShapes[0]?.slideId).toBe('s3');

    const missing = await new PowerPointBridge().actuate(
      addTableSlide({ target: { slideId: 's9' }, tableGrid: { hasHeaders: false, rows: [['a']] } }),
    );
    expect(missing).toMatchObject({ ok: false, error: { code: 'target_conflict' } });
    expect(d.addedShapes).toHaveLength(1);
  });

  it('degrades to unsupported below PowerPointApi 1.8 without touching the host', async () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { requirements: { PowerPointApi: 1.4 } });
    const res = await new PowerPointBridge().actuate(
      addTableSlide({ target: { slideId: 's1' }, tableGrid: { hasHeaders: false, rows: [['a']] } }),
    );

    expect(res).toMatchObject({ ok: false, error: { code: 'unsupported' } });
    expect(res.error?.message).toContain('PowerPointApi 1.8');
    expect(installed.ctxRef.ctx).toBeUndefined();
  });
});

/* ───────────────────────────── watch ───────────────────────────── */

describe('PowerPointBridge.watch', () => {
  it('registers selection + view handlers and emits the mapped local HostEvents', () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const events: HostEvent[] = [];
    const unsub = new PowerPointBridge().watch((e) => events.push(e));

    // Two handlers registered: selection-changed and active-view-changed.
    expect(installed.office.added).toHaveLength(2);
    const [selReg, viewReg] = installed.office.added;
    selReg?.handler();
    viewReg?.handler();

    expect(events).toEqual([
      { type: 'selection-changed', surface: 'powerpoint', origin: 'local' },
      { type: 'document-changed', surface: 'powerpoint', origin: 'local' },
    ]);

    unsub();
    // Teardown removed both handlers.
    expect(installed.office.removed).toHaveLength(2);
  });

  it('is idempotent: a second unsubscribe does not remove handlers again', () => {
    installed = install(deck(SAMPLE_SLIDES, [0]));
    const unsub = new PowerPointBridge().watch(() => {});
    unsub();
    const removedAfterFirst = installed.office.removed.length;
    unsub();
    expect(installed.office.removed.length).toBe(removedAfterFirst);
  });

  it('never throws when the host has no event bus (addHandlerAsync throws) and unsub is a no-op', () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { officeThrowsOnAdd: true });
    const events: HostEvent[] = [];
    const bridge = new PowerPointBridge();
    let unsub: () => void = () => {};
    expect(() => {
      unsub = bridge.watch((e) => events.push(e));
    }).not.toThrow();
    expect(installed.office.added).toHaveLength(0);
    // Nothing was registered, so teardown removes nothing and does not throw.
    expect(() => unsub()).not.toThrow();
    expect(installed.office.removed).toHaveLength(0);
  });

  it('swallows a throwing removeHandlerAsync during teardown (best-effort)', () => {
    installed = install(deck(SAMPLE_SLIDES, [0]), { officeThrowsOnRemove: true });
    const unsub = new PowerPointBridge().watch(() => {});
    expect(() => unsub()).not.toThrow();
  });
});
