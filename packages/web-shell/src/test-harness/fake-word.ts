/**
 * In-memory **Word host simulator**. Models the slice of the Office.js object model the real
 * {@link "@ge/bridge-word"!OfficeWordHost} drives (the bridge's only un-faked seam), so the REAL
 * bridge runs unchanged against a seeded document — including the load-bearing `body.search`
 * re-resolution and the search→insert tracked-change batch ordering.
 *
 * Enumerated host calls modelled (the fidelity boundary for Word):
 *   - `Word.run(cb)`; `Word.InsertLocation.replace`; `Word.ChangeTrackingMode.trackAll`.
 *   - `ctx.document.getSelection()` → `.text`.
 *   - `ctx.document.body` → `.text`; `.paragraphs` → items `{ text, styleBuiltIn }`.
 *   - `ctx.document.body.search(q, { matchCase })` → `results.items[]` with `.text` and
 *     `.paragraphs.getFirstOrNullObject()` → `.text` / `.isNullObject` (context-hint read).
 *   - `range.insertText(text, location)` (the tracked-change WRITE) — re-resolved at apply-time
 *     against the live body, recorded into the seed and reflected in subsequent searches.
 *   - `range.insertComment(text)` (comments-as-citations WRITE).
 *   - `ctx.document.changeTrackingMode = trackAll` (recorded for assertions).
 *   - `ctx.document.body.getComments()` → items `{ id }`; `comment.reply(text)`; `comment.resolved=`.
 *   - `ctx.document.customXmlParts.add(xml)` (durable provenance WRITE, via `Office`).
 *
 * Fidelity notes / boundary:
 *   - `body.search` is a case-insensitive substring scan over each paragraph's text; a hit's text is
 *     the matched query (as Office returns the matched range), and its first-paragraph context hint
 *     is the containing paragraph. This is faithful enough for anchor choice + drift degradation.
 *   - `insertText(replace)` replaces the FIRST occurrence of the anchor text in the body, so a later
 *     search for the old anchor degrades to drift (the apply-time re-resolution the bridge relies on).
 */

import { installGlobal, composeRestores } from './globals.js';
import {
  makeFakeOffice,
  makeOfficeSeed,
  type OfficeSeed,
  type OfficeHandlerRegistry,
} from './fake-office.js';

/** A document paragraph: its text + the Word built-in style name (drives heading level). */
export interface WordParagraphSeed {
  text: string;
  styleBuiltIn: string;
}

/** A document comment (the `getComments()` shape). */
export interface WordCommentSeed {
  id: string;
  text: string;
  replies: string[];
  resolved: boolean;
}

/** The Word document seed: selection text + ordered paragraphs + comments. */
export interface WordSeed {
  selectionText: string;
  paragraphs: WordParagraphSeed[];
  comments: WordCommentSeed[];
  /** Recorded: the last `changeTrackingMode` the bridge set (asserts tracked changes were turned on). */
  changeTrackingMode?: string;
  /** Recorded tracked-change inserts: `{ anchor, text }` per applied replace. */
  inserts: Array<{ anchor: string; text: string }>;
  /** Recorded comments added via `insertComment`: `{ anchor, text }`. */
  addedComments: Array<{ anchor: string; text: string }>;
}

/** The whole body text — paragraphs joined by newlines (mirrors `body.text`). */
function bodyText(seed: WordSeed): string {
  return seed.paragraphs.map((p) => p.text).join('\n');
}

/* ─────────────────────────── fake object model ─────────────────────────── */

const INSERT_LOCATION = { replace: 'Replace', start: 'Start', end: 'End' } as const;
const CHANGE_TRACKING_MODE = { trackAll: 'TrackAll', off: 'Off' } as const;

class FakeParagraphProxy {
  isNullObject = false;
  text = '';
  constructor(
    text: string | undefined,
    private readonly seed?: WordSeed,
    private readonly index = -1,
  ) {
    if (text === undefined) this.isNullObject = true;
    else this.text = text;
  }
  load(_props?: string): this {
    return this;
  }
  /** WordApi 1.3: the following paragraph, or a null object after the last one. */
  getNextOrNullObject(): FakeParagraphProxy {
    const next = this.seed?.paragraphs[this.index + 1];
    return new FakeParagraphProxy(next?.text, this.seed, this.index + 1);
  }
}

class FakeSearchResult {
  text: string;
  constructor(
    private readonly seed: WordSeed,
    matchText: string,
  ) {
    this.text = matchText;
  }
  load(_props?: string): this {
    return this;
  }
  /** WordApi 1.6 `Range.getTrackedChanges` — the simulated body holds no tracked deletions. */
  getTrackedChanges(): { items: Array<{ type: string }>; load(): void } {
    return { items: [], load: () => undefined };
  }
  get paragraphs(): { getFirstOrNullObject(): FakeParagraphProxy } {
    const index = this.seed.paragraphs.findIndex((p) =>
      p.text.toLowerCase().includes(this.text.toLowerCase()),
    );
    const containing = this.seed.paragraphs[index];
    return {
      getFirstOrNullObject: () => new FakeParagraphProxy(containing?.text, this.seed, index),
    };
  }
  /** The tracked-change WRITE: replace the first occurrence of the anchor in the body. */
  insertText(text: string, _location: string): void {
    this.seed.inserts.push({ anchor: this.text, text });
    replaceFirst(this.seed, this.text, text);
  }
  /** Comments-as-citations WRITE. */
  insertComment(text: string): void {
    this.seed.addedComments.push({ anchor: this.text, text });
    this.seed.comments.push({
      id: `sim-word-comment-${this.seed.comments.length + 1}`,
      text,
      replies: [],
      resolved: false,
    });
  }
}

/** Incremented by every `ctx.sync()`; a proxy's loaded properties exist only after one. */
let syncGeneration = 0;

class FakeSearchResultCollection {
  private readonly hits: FakeSearchResult[];
  private readonly createdAt = syncGeneration;
  constructor(seed: WordSeed, query: string, matchCase: boolean) {
    const hay = (s: string): string => (matchCase ? s : s.toLowerCase());
    const needle = matchCase ? query : query.toLowerCase();
    // One hit per paragraph that contains the query (mirrors Office's per-occurrence ranges,
    // collapsed to paragraph granularity — enough for anchor choice + drift).
    this.hits = seed.paragraphs
      .filter((p) => hay(p.text).includes(needle))
      .map(() => new FakeSearchResult(seed, query));
  }
  /** Like Word, `items` is unavailable until a sync: reading it first throws PropertyNotLoaded. */
  get items(): FakeSearchResult[] {
    if (syncGeneration === this.createdAt)
      throw new Error(
        "fake-word: The property 'items' is not available. Call load and context.sync() first.",
      );
    return this.hits;
  }
  load(_props?: string): this {
    return this;
  }
}

class FakeCommentProxy {
  constructor(private readonly target: WordCommentSeed) {}
  get id(): string {
    return this.target.id;
  }
  reply(text: string): void {
    this.target.replies.push(text);
  }
  set resolved(value: boolean) {
    this.target.resolved = value;
  }
  get resolved(): boolean {
    return this.target.resolved;
  }
}

class FakeCommentCollection {
  items: FakeCommentProxy[];
  constructor(seed: WordSeed) {
    this.items = seed.comments.map((c) => new FakeCommentProxy(c));
  }
  load(_props?: string): this {
    return this;
  }
}

class FakeParagraphCollection {
  items: WordParagraphSeed[];
  constructor(seed: WordSeed) {
    this.items = seed.paragraphs.map((p) => ({ text: p.text, styleBuiltIn: p.styleBuiltIn }));
  }
  load(_props?: string): this {
    return this;
  }
}

class FakeBody {
  constructor(private readonly seed: WordSeed) {}
  get text(): string {
    return bodyText(this.seed);
  }
  load(_props?: string): this {
    return this;
  }
  get paragraphs(): FakeParagraphCollection {
    return new FakeParagraphCollection(this.seed);
  }
  search(query: string, opts?: { matchCase?: boolean }): FakeSearchResultCollection {
    return new FakeSearchResultCollection(this.seed, query, opts?.matchCase ?? false);
  }
  getComments(): FakeCommentCollection {
    return new FakeCommentCollection(this.seed);
  }
}

class FakeSelection {
  constructor(private readonly seed: WordSeed) {}
  get text(): string {
    return this.seed.selectionText;
  }
  load(_props?: string): this {
    return this;
  }
}

class FakeCustomXmlPartCollection {
  constructor(private readonly office: OfficeSeed) {}
  add(xml: string): { id: string } {
    this.office.customXmlParts.push(xml);
    return { id: `xml-${this.office.customXmlParts.length}` };
  }
}

class FakeWordDocument {
  private _mode = '';
  readonly body: FakeBody;
  readonly customXmlParts: FakeCustomXmlPartCollection;
  constructor(
    private readonly seed: WordSeed,
    office: OfficeSeed,
  ) {
    this.body = new FakeBody(seed);
    this.customXmlParts = new FakeCustomXmlPartCollection(office);
  }
  getSelection(): FakeSelection {
    return new FakeSelection(this.seed);
  }
  set changeTrackingMode(mode: string) {
    this._mode = mode;
    this.seed.changeTrackingMode = mode;
  }
  get changeTrackingMode(): string {
    return this._mode;
  }
}

class FakeWordContext {
  readonly document: FakeWordDocument;
  constructor(seed: WordSeed, office: OfficeSeed) {
    this.document = new FakeWordDocument(seed, office);
  }
  sync(): Promise<void> {
    syncGeneration += 1;
    return Promise.resolve();
  }
}

/** The `Word` namespace object installed onto `globalThis.Word`. */
interface FakeWordNamespace {
  run<T>(callback: (ctx: FakeWordContext) => Promise<T>): Promise<T>;
  InsertLocation: typeof INSERT_LOCATION;
  ChangeTrackingMode: typeof CHANGE_TRACKING_MODE;
}

/** Replace the first paragraph occurrence of `anchor` with `replacement` in the seed body. */
function replaceFirst(seed: WordSeed, anchor: string, replacement: string): void {
  const needle = anchor.toLowerCase();
  for (const p of seed.paragraphs) {
    const idx = p.text.toLowerCase().indexOf(needle);
    if (idx >= 0) {
      p.text = p.text.slice(0, idx) + replacement + p.text.slice(idx + anchor.length);
      return;
    }
  }
}

/* ─────────────────────────── the simulator facade ──────────────────────── */

/** A read-back view of the Word document after a run. */
export interface WordSnapshot {
  bodyText: string;
  paragraphs: ReadonlyArray<WordParagraphSeed>;
  comments: ReadonlyArray<WordCommentSeed>;
  inserts: ReadonlyArray<{ anchor: string; text: string }>;
  addedComments: ReadonlyArray<{ anchor: string; text: string }>;
  changeTrackingMode?: string;
}

/** The installed Word simulator. */
export interface WordSimulator {
  readonly seed: WordSeed;
  readonly office: OfficeSeed;
  readonly officeHandlers: OfficeHandlerRegistry;
  snapshot(): WordSnapshot;
  restore(): void;
}

/**
 * Install an in-memory Word host: writes `globalThis.Word` + `globalThis.Office` so the REAL
 * {@link "@ge/bridge-word"!WordBridge} runs against `seed`. Defaults to {@link defaultWordSeed}
 * (a contract-ish document) + a modern requirement set.
 */
export function installFakeWord(
  seed: WordSeed = defaultWordSeed(),
  requirements: Record<string, number> = { WordApi: 6 },
): WordSimulator {
  const office = makeOfficeSeed(requirements);
  const { office: officeNs, handlers: officeHandlers } = makeFakeOffice(office);

  const word: FakeWordNamespace = {
    run: async <T>(callback: (ctx: FakeWordContext) => Promise<T>): Promise<T> =>
      callback(new FakeWordContext(seed, office)),
    InsertLocation: INSERT_LOCATION,
    ChangeTrackingMode: CHANGE_TRACKING_MODE,
  };

  const restore = composeRestores([installGlobal('Word', word), installGlobal('Office', officeNs)]);

  return {
    seed,
    office,
    officeHandlers,
    snapshot: () => ({
      bodyText: bodyText(seed),
      paragraphs: seed.paragraphs.map((p) => ({ ...p })),
      comments: seed.comments.map((c) => ({ ...c, replies: [...c.replies] })),
      inserts: seed.inserts.map((i) => ({ ...i })),
      addedComments: seed.addedComments.map((c) => ({ ...c })),
      ...(seed.changeTrackingMode !== undefined
        ? { changeTrackingMode: seed.changeTrackingMode }
        : {}),
    }),
    restore,
  };
}

/* ─────────────────────────── builders + default fixture ─────────────────── */

/** Build a {@link WordSeed} from paragraphs, defaulting empty selection/comments + write logs. */
export function wordSeed(init: {
  selectionText?: string;
  paragraphs: WordParagraphSeed[];
  comments?: WordCommentSeed[];
}): WordSeed {
  return {
    selectionText: init.selectionText ?? '',
    paragraphs: init.paragraphs,
    comments: init.comments ?? [],
    inserts: [],
    addedComments: [],
  };
}

/** A contract-ish Word document fixture: a heading + body paragraphs incl. an SLA claim to redline. */
export function defaultWordSeed(): WordSeed {
  return wordSeed({
    selectionText: 'The SLA is 99.5%.',
    paragraphs: [
      { text: 'Master Services Agreement', styleBuiltIn: 'Heading1' },
      { text: '1. Service Levels', styleBuiltIn: 'Heading2' },
      {
        text: 'The SLA is 99.5%. Downtime beyond this entitles the customer to service credits.',
        styleBuiltIn: 'Normal',
      },
      {
        text: 'Either party may terminate for convenience on 30 days written notice.',
        styleBuiltIn: 'Normal',
      },
    ],
    comments: [
      { id: 'c-existing-1', text: 'Confirm the SLA figure.', replies: [], resolved: false },
    ],
  });
}
