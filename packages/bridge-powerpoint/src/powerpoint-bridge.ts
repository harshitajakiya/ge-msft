import { unknownActuationResult } from '@ge/contracts';
import { createBridgeDispatch } from '@ge/runtime';
import type {
  ActuationKind,
  ActuationRequest,
  ActuationResult,
  CapabilityManifest,
  ContextRef,
  DocStateSnapshot,
  ResolvedContext,
} from '@ge/contracts';
import type { DocBridge } from '@ge/runtime';
import type { HostEvent, Unsubscribe } from '@ge/triggers';
import { buildDocStateSnapshot } from '@ge/content';
import { POWERPOINT_CAPABILITIES } from './capabilities.js';
import { isSet } from './capabilities-runtime.js';
import {
  mayNameSlide,
  searchSlides,
  selectedShapeToContext,
  selectedSlideToContext,
  shapeContextRef,
  slideContextRef,
  shapesToSlideText,
  slideElementsToDocStateBlocks,
  slideIndexForSelector,
  slideShapeListing,
  slidesToContext,
  tableValuesToText,
  type SlideElement,
} from './capture.js';
import { planInsertSlide } from './actuate-plan.js';
import { documentChanged, selectionChanged } from './events.js';

/**
 * Upper bound on slides materialized by a single read port (`captureDocState`/`searchDocument`/
 * `readRange`). A deck can be large; reading every shape's text across the Office.js bridge is
 * O(slides × shapes) syncs, so we cap the scan to keep a per-turn read bounded (ADR-0006 — every
 * read is bounded). A deck over this is read as its first {@link MAX_READ_SLIDES} slides.
 */
export const MAX_READ_SLIDES = 60;

/**
 * The PowerPoint `DocBridge`. The ONLY place Office.js (`PowerPoint.run`) is touched. Reads via
 * the native object model (selected slides → shapes' text + speaker notes); writes by composing
 * slides into the deck (`insertSlidesFromBase64` for a prebuilt deck, else `slides.add()` +
 * placeholder text) and setting speaker notes — each reversible and provenanced via the shared
 * `ActuationResult` shape. Pure mapping lives in `capture.ts` / `actuate-plan.ts` (unit-tested);
 * this file is the host wiring.
 *
 * Requirement-set versions used (confirmed against node_modules/@types/office-js/index.d.ts):
 *   - `Presentation.getSelectedSlides()` / `getSelectedShapes()` → PowerPointApi 1.5 (l.178785 / l.178776).
 *   - `Slide.shapes` (ShapeCollection) → PowerPointApi 1.3 (l.186098); `Shape.textFrame` → 1.4 (l.186557);
 *     `TextRange.text` (read/write) → PowerPointApi 1.4 (l.180262-180267).
 *   - `Slide.id` → 1.2 (l.186126), `Slide.index` → 1.8 (l.186133).
 *   - `SlideCollection.add()` → PowerPointApi 1.3 (l.187424); `insertSlidesFromBase64` → 1.2 (l.178812).
 *   - `ShapeCollection.addTextBox` / `addGeometricShape` / `addLine` → PowerPointApi 1.4
 *     (l.184213 / l.184146 / l.184178) with `ShapeAddOptions` geometry (l.181900).
 *   - `Shape.fill` → 1.4 (l.186512; `ShapeFill.foregroundColor` l.182363), `Shape.lineFormat` → 1.4
 *     (l.186527; `ShapeLineFormat.color` l.186392), `TextRange.font` (`ShapeFont`) → 1.4
 *     (l.180230, class l.179862); `Shape.setZOrder` → PowerPointApi 1.8 (l.186782).
 *   - `ShapeCollection.addTable(rowCount, columnCount, TableAddOptions)` → PowerPointApi 1.8
 *     (l.184202; options l.184011); `Shape.getTable()` → 1.8 (l.186746); `Table.getCellOrNullObject`
 *     → 1.8 (l.183654); `TableCell.text` (read/write) → 1.8 (l.182631).
 *   - There is NO PowerPoint image-insertion API in this typings version (`addImage` exists only on
 *     Excel's ShapeCollection, l.55471; PPT's `getImageAsBase64` is read-only), so `insert-image`
 *     stays un-advertised for this surface.
 *   - PowerPoint exposes NO object-model selection/change event in this typings, so `watch` uses the
 *     Office-level `Office.EventType.DocumentSelectionChanged` (l.645) + `ActiveViewChanged` (l.582)
 *     with add/removeHandlerAsync (l.3875 / l.3965). Neither carries a coauthor source → origin 'local'.
 */

export class PowerPointBridge implements DocBridge {
  private static readonly dispatcher = createBridgeDispatch<PowerPointBridge>(
    'powerpoint',
    {
      'insert-slide': (host, request) => host.applyInsertSlide(request),
      'set-shape-text': (host, request) => host.applySetShapeText(request),
      'add-shape': (host, request) => host.applyAddShape(request),
      'format-shape': (host, request) => host.applyFormatShape(request),
      'add-table-slide': (host, request) => host.applyAddTableSlide(request),
    },
    { provenance: 'unsupported' },
  );
  static readonly handledActuations = PowerPointBridge.dispatcher.handledActuations;

  readonly surface = 'powerpoint' as const;

  getCapabilities(): CapabilityManifest {
    return POWERPOINT_CAPABILITIES;
  }

  async listContext(): Promise<ContextRef[]> {
    // getSelectedSlides is PowerPointApi 1.5; on an older host we can still list the deck.
    if (!isSet('PowerPointApi', '1.5')) {
      return [{ id: 'pp:deck', kind: 'document', surface: 'powerpoint', title: 'Whole deck' }];
    }
    return PowerPoint.run(async (ctx) => {
      const selected = ctx.presentation.getSelectedSlides();
      selected.load('items/id,items/index');
      await ctx.sync();

      const refs: ContextRef[] = [];
      const first = selected.items[0];
      if (first) {
        const slide = await readSlide(ctx, first);
        refs.push(slideContextRef(slide));
        // Only shapes with text: an empty line or picture has nothing to attach (resolving it gives
        // no context), and listing it gave the model a "shape" to aim a text write at.
        for (const shape of slide.shapes ?? []) {
          if (shape.text.trim()) refs.push(shapeContextRef(slide, shape));
        }
      }
      refs.push({ id: 'pp:deck', kind: 'document', surface: 'powerpoint', title: 'Whole deck' });
      return refs;
    });
  }

  async resolveContext(ref: ContextRef): Promise<ResolvedContext[]> {
    if (ref.kind === 'shape') {
      const target = shapeRevealTarget(ref);
      if (!target) return [];
      return PowerPoint.run(async (ctx) => readShapeContext(ctx, target));
    }
    if (ref.kind === 'slide') {
      const target = slideRevealTarget(ref);
      if (!target) return [];
      try {
        return await PowerPoint.run(async (ctx) => {
          const slide = ctx.presentation.slides.getItem(target.slideId);
          const element = await readSlide(ctx, slide);
          return selectedSlideToContext(element);
        });
      } catch {
        return [];
      }
    }
    if (ref.kind === 'selection') {
      return PowerPoint.run(async (ctx) => {
        const selected = ctx.presentation.getSelectedSlides();
        selected.load('items/id,items/index');
        await ctx.sync();
        const slide = selected.items[0];
        if (!slide) return [];
        const element = await readSlide(ctx, slide);
        return selectedSlideToContext(element);
      });
    }
    // Whole deck → each slide's shape text → native blocks → chunks (bounded, batched read).
    return slidesToContext('pp:deck', 'Whole deck', await this.readAllSlides());
  }

  /** Monotonic `<doc_state>` version, bumped on each capture (ADR-0003 Layer B element 1). */
  private docStateVersion = 0;

  /**
   * ADR-0003 Layer B element 1 / ADR-0006 `outline` read: an ambient structural snapshot of the
   * deck — a slide inventory (per-slide title + body) read from the native object model and mapped
   * through the same `native.slide()` blocks grounding context uses. Bounded to
   * {@link MAX_READ_SLIDES} slides; the snapshot's `inventory` lists each slide. Reading the deck
   * needs `TextRange.text` (PowerPointApi 1.4) — on an older host we yield `undefined` (the runtime
   * just streams without the ambient part). Empty deck → `undefined`. Version increments per capture.
   */
  async captureDocState(): Promise<DocStateSnapshot | undefined> {
    if (!isSet('PowerPointApi', '1.4')) return undefined;
    const slides = await this.readAllSlides();
    if (slides.length === 0) return undefined;
    this.docStateVersion += 1;
    return buildDocStateSnapshot({
      surface: 'powerpoint',
      version: this.docStateVersion,
      blocks: slideElementsToDocStateBlocks(slides),
    });
  }

  /**
   * ADR-0006 `search` read: scan the deck's slide text for `query` and return matching slides as
   * `ResolvedContext` data (never instructions), bounded by `searchSlides`. Reads via the native
   * model (gated on PowerPointApi 1.4); empty query / older host / no match → `[]`.
   */
  async searchDocument(query: string): Promise<ResolvedContext[]> {
    const q = query.trim();
    if (!q || !isSet('PowerPointApi', '1.4')) return [];
    const slides = await this.readAllSlides();
    return searchSlides(slides, q);
  }

  /**
   * ADR-0006 addressable `read <slide:N>` verb: resolve a slide selector (`slide:N` / `slide N` /
   * bare 1-based `N`) to that single slide's text as `ResolvedContext` data. Unaddressable selectors
   * (a name, junk) / out-of-range index / older host / empty deck → `[]` — the bridge degrades rather
   * than guessing. Reads via the native model (gated on PowerPointApi 1.4); the read is one slide, so
   * it is inherently bounded.
   */
  async readRange(selector: string): Promise<ResolvedContext[]> {
    if (!mayNameSlide(selector) || !isSet('PowerPointApi', '1.4')) return [];
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id,items/index');
      await ctx.sync();
      const index = slideIndexForSelector(
        slides.items.map((item) => item.id),
        selector,
      );
      const slide = index === undefined ? undefined : slides.items[index];
      if (!slide) return [];
      const element = await readSlide(ctx, slide);
      // A single-slide read is how the model looks before a shape command: list each shape's id,
      // type and which one `shape=title` resolves to, so it can address the shape it means.
      const shapes = ctx.presentation.slides.getItem(slide.id).shapes;
      shapes.load('items/id,items/type');
      await ctx.sync();
      const title = await findTitleShape(ctx, shapes.items).catch(() => undefined);
      const types = new Map(shapes.items.map((shape) => [shape.id, shape.type as string]));
      const listed: SlideElement = {
        ...element,
        shapes: (element.shapes ?? []).map((shape) => ({
          ...shape,
          ...(types.get(shape.shapeId) ? { type: types.get(shape.shapeId) } : {}),
          ...(title?.id === shape.shapeId ? { isTitle: true } : {}),
        })),
      };
      return selectedSlideToContext({
        ...element,
        body: [...element.body, ...slideShapeListing(listed)],
      });
    });
  }

  canRevealContext(ref: ContextRef): boolean {
    return ref.surface === 'powerpoint' && powerpointRevealTarget(ref) !== undefined;
  }

  async revealContext(ref: ContextRef): Promise<void> {
    const target = powerpointRevealTarget(ref);
    if (!target) return;
    await PowerPoint.run(async (ctx) => {
      ctx.presentation.setSelectedSlides([target.slideId]);
      if (target.shapeId) {
        const slide = ctx.presentation.slides.getItem(target.slideId);
        slide.setSelectedShapes([target.shapeId]);
      }
      await ctx.sync();
    });
  }

  /**
   * Read up to {@link MAX_READ_SLIDES} slides of the deck into pure {@link SlideElement}s — the
   * shared host read behind the whole-deck context, `captureDocState` and `searchDocument`. Bounded
   * so a huge deck can't blow the per-turn budget; read-only (loads shape text, writes nothing).
   */
  private async readAllSlides(): Promise<SlideElement[]> {
    return PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id,items/index');
      await ctx.sync();
      return readSlides(ctx, slides.items.slice(0, MAX_READ_SLIDES));
    });
  }

  async actuate(req: ActuationRequest): Promise<ActuationResult> {
    return PowerPointBridge.dispatcher.dispatch(this, req);
  }

  private async applySetShapeText(req: ActuationRequest): Promise<ActuationResult> {
    const slideId = req.params.target?.slideId;
    const shapeId = req.params.target?.shapeId;
    const text = req.params.text;
    if (!slideId || !shapeId) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'no_target',
          message: 'set-shape-text needs target.slideId and target.shapeId',
        },
      };
    }
    if (text === undefined) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_text', message: 'set-shape-text needs params.text' },
      };
    }
    if (!isSet('PowerPointApi', '1.4')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'unsupported', message: 'PowerPointApi 1.4 is required for shape text.' },
      };
    }

    let mutationQueued = false;
    try {
      return await PowerPoint.run(async (ctx) => {
        const target = await resolveShapeTarget(ctx, slideId, shapeId);
        if ('error' in target) return slideNotFound(req, target.error);
        const slide = ctx.presentation.slides.getItem(target.slideId);
        const shape = slide.shapes.getItemOrNullObject(target.shapeId);
        shape.load('isNullObject,type');
        await ctx.sync();
        if (shape.isNullObject) {
          return {
            ok: false,
            changeId: req.changeId,
            kind: req.kind,
            degraded: true,
            error: {
              code: 'target_conflict',
              message: 'The selected PowerPoint shape no longer exists.',
            },
          };
        }
        if (!TEXT_SHAPE_TYPES.has(shape.type)) {
          return slideNotFound(
            req,
            `Shape id ${target.shapeId} is a ${shape.type} and has no text. For the slide title use ` +
              `shape pp:shape:${slideId}:title "…", or the id of a text box.`,
          );
        }
        const range = shape.textFrame.textRange;
        range.load('text');
        await ctx.sync();
        const priorText = range.text ?? '';
        mutationQueued = true;
        range.text = text;
        await ctx.sync();
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${target.slideId}:${target.shapeId}`,
          inverse: {
            op: 'restore-text',
            anchor: `pp:shape:${target.slideId}:${target.shapeId}`,
            priorText,
          },
        };
      });
    } catch (error) {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          `PowerPoint did not confirm the dispatched change${hostErrorSuffix(error)}. ` +
            'Inspect the slide before trying again.',
        );
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'target_conflict',
          message: 'The PowerPoint shape could not be re-read before writing.',
        },
      };
    }
  }

  private async applyInsertSlide(req: ActuationRequest): Promise<ActuationResult> {
    const plan = planInsertSlide(req);
    if (!plan.base64 && !plan.title && plan.bullets.length === 0) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'empty_slide', message: 'insert-slide needs params.slide or params.ooxml' },
      };
    }
    if (plan.base64) {
      const base64 = plan.base64;
      return PowerPoint.run(async (ctx) => {
        // Prebuilt deck path: the agent supplied a Base64 PPTX — let the host merge it (1.2).
        const options = await deckInsertOptions(ctx, plan);
        ctx.presentation.insertSlidesFromBase64(base64, options);
        await ctx.sync();
        const location =
          plan.slideCount === undefined ? 'inserted-deck' : `inserted-deck:${plan.slideCount}`;
        return { ok: true, changeId: req.changeId, kind: req.kind, location };
      });
    }

    // Native compose path: append a slide, then write the title/bullets into it.
    //
    // PowerPoint for the web gives a just-added slide a PROVISIONAL id: writes addressed through
    // that id — `getItem(id)` (5010 "InvalidParam passed to GetItem(id)") or the slide object from
    // the `items` list (GeneralException) — are rejected, while writes by position (`getItemAt`)
    // work; the slide gets its real id once the request completes. So: find the new slide's
    // POSITION by the one id that is new (never a stale index — a co-author's concurrent add must
    // not redirect the text), write to it by position, and read its real id back afterwards.
    let newIndex: number | undefined;
    const result = await PowerPoint.run(async (ctx): Promise<ActuationResult> => {
      try {
        const added = await appendSlide(ctx);
        if ('error' in added) return unknownActuationResult(req, added.error);
        const { index, provisionalId, shapes } = added;
        await writeComposedSlide(ctx, index, shapes, plan.title, plan.bullets);
        await ctx.sync();
        newIndex = index;
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `slide:${provisionalId}`,
        };
      } catch (error) {
        // The slide may already be in the deck; never report an unconfirmed write as success.
        return unknownActuationResult(
          req,
          `PowerPoint did not confirm the new slide and its text${hostErrorSuffix(error)}. ` +
            'Inspect the deck before trying again.',
        );
      }
    });
    if (!result.ok || newIndex === undefined) return result;
    // Best effort: report the slide's settled id (the one `outline` and later commands will see).
    const settledId = await readSlideIdAt(newIndex);
    return settledId ? { ...result, location: `slide:${settledId}` } : result;
  }

  private async applyAddShape(req: ActuationRequest): Promise<ActuationResult> {
    const resolution = resolveAddShape(req);
    if (!resolution.ok) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: resolution.code, message: resolution.message },
      };
    }
    if (!isSet('PowerPointApi', '1.4')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'unsupported', message: 'PowerPointApi 1.4 is required to add shapes.' },
      };
    }
    let mutationQueued = false;
    try {
      return await PowerPoint.run(async (ctx) => {
        const op = resolution.op;
        const target = await resolveSlideRef(ctx, op.slideId);
        if ('error' in target) return slideNotFound(req, target.error);
        const size = await knownSlideSize(ctx);
        const unplaced = missingGeometryMessage(size, op);
        if (unplaced) return slideNotFound(req, unplaced);
        const offSlide = offSlideMessage(size, op);
        if (offSlide) return slideNotFound(req, offSlide);
        const slide = ctx.presentation.slides.getItem(target.slideId);
        const options: PowerPoint.ShapeAddOptions = {};
        if (op.left !== undefined) options.left = op.left;
        if (op.top !== undefined) options.top = op.top;
        if (op.width !== undefined) options.width = op.width;
        if (op.height !== undefined) options.height = op.height;
        mutationQueued = true;
        const added =
          op.type === 'textBox'
            ? slide.shapes.addTextBox(op.text, options)
            : op.type === 'line'
              ? slide.shapes.addLine(op.connector, options)
              : slide.shapes.addGeometricShape(op.geometry, options);
        if (op.fill !== undefined) added.fill.foregroundColor = op.fill;
        added.load('id');
        await ctx.sync();
        const mintedId = added.id;
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${target.slideId}:${mintedId}`,
          inverse: { op: 'delete-object', objectType: 'shape', name: mintedId },
        };
      });
    } catch (error) {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          `PowerPoint did not confirm the dispatched change${hostErrorSuffix(error)}. ` +
            'Inspect the slide before trying again.',
        );
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'target_conflict',
          message: 'The PowerPoint slide could not be read before adding the shape.',
        },
      };
    }
  }

  private async applyFormatShape(req: ActuationRequest): Promise<ActuationResult> {
    const slideId = req.params.target?.slideId;
    const shapeId = req.params.target?.shapeId;
    const format = req.params.shapeFormat;
    if (!slideId || !shapeId) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'no_target',
          message: 'format-shape needs target.slideId and target.shapeId',
        },
      };
    }
    if (!format || Object.keys(format).length === 0) {
      // An empty format would "apply" nothing and still report success.
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'no_format',
          message:
            'format-shape needs at least one of fill=, line=, fontColor=, fontSize=, fontBold=, ' +
            "fontItalic=, fontName= or zOrder=. To change a shape's text use: shape " +
            'pp:shape:<slide>:<shape> "text".',
        },
      };
    }
    if (!isSet('PowerPointApi', '1.4')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'unsupported',
          message: 'PowerPointApi 1.4 is required for shape formatting.',
        },
      };
    }
    if (format.zOrder !== undefined && !isSet('PowerPointApi', '1.8')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'unsupported', message: 'PowerPointApi 1.8 is required for shape z-order.' },
      };
    }

    let mutationQueued = false;
    try {
      return await PowerPoint.run(async (ctx) => {
        const target = await resolveShapeTarget(ctx, slideId, shapeId);
        if ('error' in target) return slideNotFound(req, target.error);
        const slide = ctx.presentation.slides.getItem(target.slideId);
        const shape = slide.shapes.getItemOrNullObject(target.shapeId);
        shape.load('isNullObject');
        await ctx.sync();
        if (shape.isNullObject) {
          return {
            ok: false,
            changeId: req.changeId,
            kind: req.kind,
            degraded: true,
            error: {
              code: 'target_conflict',
              message: 'The selected PowerPoint shape no longer exists.',
            },
          };
        }
        // Capture each prior value just before overwriting it so the recorded inverse holds only
        // what THIS change mutated (restore-shape-format prior keys mirror the params fields).
        const prior: Record<string, string> = {};
        if (format.fill !== undefined) {
          const fill = shape.fill;
          fill.load('foregroundColor');
          await ctx.sync();
          prior['fill'] = String(fill.foregroundColor ?? '');
          mutationQueued = true;
          fill.foregroundColor = format.fill;
        }
        if (format.line !== undefined) {
          const line = shape.lineFormat;
          line.load('color');
          await ctx.sync();
          prior['line'] = String(line.color ?? '');
          mutationQueued = true;
          line.color = format.line;
        }
        if (format.font !== undefined) {
          const font = shape.textFrame.textRange.font;
          font.load('bold,italic,color,size,name,underline');
          await ctx.sync();
          const f = format.font;
          if (f.bold !== undefined) {
            prior['font.bold'] = String(font.bold ?? '');
            mutationQueued = true;
            font.bold = f.bold;
          }
          if (f.italic !== undefined) {
            prior['font.italic'] = String(font.italic ?? '');
            mutationQueued = true;
            font.italic = f.italic;
          }
          if (f.underline !== undefined) {
            prior['font.underline'] = String(font.underline ?? '');
            mutationQueued = true;
            font.underline = f.underline ? 'Single' : 'None';
          }
          if (f.color !== undefined) {
            prior['font.color'] = String(font.color ?? '');
            mutationQueued = true;
            font.color = f.color;
          }
          if (f.size !== undefined) {
            prior['font.size'] = String(font.size ?? '');
            mutationQueued = true;
            font.size = f.size;
          }
          if (f.name !== undefined) {
            prior['font.name'] = String(font.name ?? '');
            mutationQueued = true;
            font.name = f.name;
          }
        }
        if (format.zOrder !== undefined) {
          mutationQueued = true;
          shape.setZOrder(Z_ORDER[format.zOrder]);
        }
        await ctx.sync();
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${target.slideId}:${target.shapeId}`,
          inverse: { op: 'restore-shape-format', shapeId: target.shapeId, prior },
        };
      });
    } catch (error) {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          `PowerPoint did not confirm the dispatched change${hostErrorSuffix(error)}. ` +
            'Inspect the slide before trying again.',
        );
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'target_conflict',
          message: 'The PowerPoint shape could not be formatted.',
        },
      };
    }
  }

  private async applyAddTableSlide(req: ActuationRequest): Promise<ActuationResult> {
    const slideId = req.params.target?.slideId;
    const grid = req.params.tableGrid;
    if (!slideId) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: {
          code: 'no_target',
          message:
            'add-table-slide needs a slide: slide=new title="…" creates a slide for the table; ' +
            `or add it to an existing slide. ${SLIDE_TARGET_HINT}`,
        },
      };
    }
    const columnCount = grid?.rows.reduce((max, row) => Math.max(max, row.length), 0) ?? 0;
    if (!grid || grid.rows.length === 0 || columnCount === 0) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'no_table', message: 'add-table-slide needs a non-empty tableGrid.rows' },
      };
    }
    if (!isSet('PowerPointApi', '1.8')) {
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        error: { code: 'unsupported', message: 'PowerPointApi 1.8 is required for slide tables.' },
      };
    }

    let mutationQueued = false;
    let newSlideIndex: number | undefined;
    try {
      const result = await PowerPoint.run(async (ctx): Promise<ActuationResult> => {
        const offSlide = offSlideMessage(await knownSlideSize(ctx), grid);
        if (offSlide) return slideNotFound(req, offSlide);
        let shapes: PowerPoint.ShapeCollection;
        let slideLocation: string;
        /** Where a table goes when the command gave no position: below a title this command wrote. */
        let belowTitle: { left: number; top: number; width: number } | undefined;
        if (/^new$/i.test(slideId.trim())) {
          // `slide=new`: this command's name promises a slide, so create one for the table. Write
          // through a fresh position proxy (see appendSlide — PowerPoint web rejects writes addressed
          // by a just-added slide's provisional id), and report the settled id afterwards.
          mutationQueued = true;
          const added = await appendSlide(ctx);
          if ('error' in added) return unknownActuationResult(req, added.error);
          newSlideIndex = added.index;
          slideLocation = added.provisionalId;
          const title = req.params.slide?.title?.trim();
          // The title goes into the layout's title placeholder when there is one, else a text box
          // (same rules as insert-slide).
          if (title) {
            await writeComposedSlide(ctx, added.index, added.shapes, title, []);
            // Without this the host drops the table near the top, over the title (seen live).
            const size = await slideSize(ctx);
            const margin = Math.round(size.width * 0.067);
            belowTitle = {
              left: margin,
              top: Math.round(size.height * 0.27),
              width: size.width - margin * 2,
            };
          }
          shapes = ctx.presentation.slides.getItemAt(added.index).shapes;
        } else {
          if (req.params.slide?.title?.trim()) {
            return slideNotFound(
              req,
              'title= is only used with slide=new (it names the slide that command creates). To ' +
                'retitle an existing slide use: shape pp:shape:<slide>:title "New title".',
            );
          }
          const target = await resolveSlideRef(ctx, slideId);
          if ('error' in target) return slideNotFound(req, target.error);
          slideLocation = target.slideId;
          shapes = ctx.presentation.slides.getItem(target.slideId).shapes;
        }
        const options: PowerPoint.TableAddOptions = {};
        if (grid.left !== undefined) options.left = grid.left;
        else if (belowTitle) options.left = belowTitle.left;
        if (grid.top !== undefined) options.top = grid.top;
        else if (belowTitle) options.top = belowTitle.top;
        if (grid.width !== undefined) options.width = grid.width;
        else if (belowTitle) options.width = belowTitle.width;
        if (grid.height !== undefined) options.height = grid.height;
        mutationQueued = true;
        const added = shapes.addTable(grid.rows.length, columnCount, options);
        const table = added.getTable();
        grid.rows.forEach((row, r) => {
          row.forEach((value, c) => {
            table.getCellOrNullObject(r, c).text = value;
          });
        });
        added.load('id');
        await ctx.sync();
        const mintedId = added.id;
        return {
          ok: true,
          changeId: req.changeId,
          kind: req.kind,
          location: `shape:${slideLocation}:${mintedId}`,
          inverse: { op: 'delete-object', objectType: 'shape', name: mintedId },
        };
      });
      if (!result.ok || newSlideIndex === undefined) return result;
      const settledId = await readSlideIdAt(newSlideIndex);
      if (!settledId) return result;
      // The command created the slide: undoing it means deleting the slide (and the table with it).
      return {
        ...result,
        location: result.location?.replace(/^shape:[^:]+(?=:)/, `shape:${settledId}`),
        inverse: { op: 'delete-object', objectType: 'slide', name: settledId },
      };
    } catch (error) {
      if (mutationQueued)
        return unknownActuationResult(
          req,
          `PowerPoint did not confirm the dispatched change${hostErrorSuffix(error)}. ` +
            'Inspect the slide before trying again.',
        );
      return {
        ok: false,
        changeId: req.changeId,
        kind: req.kind,
        degraded: true,
        error: {
          code: 'target_conflict',
          message: 'The PowerPoint slide could not be read before adding the table.',
        },
      };
    }
  }

  // NOTE: a `set-speaker-notes` actuation was handled here but ALWAYS degraded — this Office.js
  // typings version exposes no `Slide.notes`/notesSlide write path. Per ADR-0006 we removed the
  // phantom from the manifest AND its `actuate()` case (advertised==handled), rather than keep a
  // case that can never succeed. Re-add (manifest + case + CLI verb) once the host typings ship a
  // notes writer. The pure `planSpeakerNotes` plan stays in `actuate-plan.ts` for that day.

  /**
   * Stream PowerPoint host events into the trigger engine via the Office-level event bus
   * (PowerPoint has no object-model selection/change event in this typings). Each registration
   * is defensive: a failed/absent registration simply means we never emit that event — it never
   * throws. Returns an `Unsubscribe` that removes every handler we added.
   */
  watch(emit: (event: HostEvent) => void): Unsubscribe {
    let onSelection: (() => void) | undefined;
    let onView: (() => void) | undefined;

    try {
      const handler = (): void => emit(selectionChanged());
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, handler);
      onSelection = handler;
    } catch {
      // Selection observation unavailable on this host — simply don't emit it.
    }

    try {
      const handler = (): void => emit(documentChanged());
      Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, handler);
      onView = handler;
    } catch {
      // Slide-navigation observation unavailable — skip.
    }

    let unsubscribed = false;
    return () => {
      if (unsubscribed) return;
      unsubscribed = true;
      if (onSelection) {
        try {
          Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
            handler: onSelection,
          });
        } catch {
          // best-effort teardown
        } finally {
          onSelection = undefined;
        }
      }
      if (onView) {
        try {
          Office.context.document.removeHandlerAsync(Office.EventType.ActiveViewChanged, {
            handler: onView,
          });
        } catch {
          // best-effort teardown
        } finally {
          onView = undefined;
        }
      }
    };
  }
}

/**
 * Geometric shapes {@link PowerPointBridge.actuate} `add-shape` accepts — typed against the
 * literal-union overload of `ShapeCollection.addGeometricShape` (PowerPointApi 1.4, l.184157) so
 * the whitelist doubles as the host parameter type (no casts, nothing outside this set is sent).
 */
const GEOMETRIC_SHAPES = [
  'Rectangle',
  'RoundRectangle',
  'Ellipse',
  'Triangle',
  'Diamond',
  'Pentagon',
  'Hexagon',
  'Octagon',
  'Star5',
  'Chevron',
  'RightArrow',
  'LeftRightArrow',
  'Cloud',
] as const;
type GeometricShapeName = (typeof GEOMETRIC_SHAPES)[number];

/** Contract connector names → the host's `ConnectorType` literals (`addLine`, PowerPointApi 1.4). */
const CONNECTOR_TYPES: Record<'straight' | 'elbow' | 'curve', 'Straight' | 'Elbow' | 'Curve'> = {
  straight: 'Straight',
  elbow: 'Elbow',
  curve: 'Curve',
};

/** Contract z-order names → the host's `ShapeZOrder` literals (`setZOrder`, PowerPointApi 1.8). */
const Z_ORDER: Record<
  'front' | 'back' | 'forward' | 'backward',
  'BringToFront' | 'SendToBack' | 'BringForward' | 'SendBackward'
> = {
  front: 'BringToFront',
  back: 'SendToBack',
  forward: 'BringForward',
  backward: 'SendBackward',
};

interface AddShapeGeometry {
  slideId: string;
  fill?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}

/**
 * The fully-resolved host plan for an `add-shape` actuation. The discriminated `type` carries the
 * whitelisted geometry so every branch of the bridge's add call is exhaustively narrowed.
 */
type AddShapeOp =
  | (AddShapeGeometry & { type: 'textBox'; text: string })
  | (AddShapeGeometry & { type: 'line'; connector?: 'Straight' | 'Elbow' | 'Curve' })
  | (AddShapeGeometry & { type: 'geometric'; geometry: GeometricShapeName });

type AddShapeResolution =
  | { ok: true; op: AddShapeOp }
  | {
      ok: false;
      code: 'no_target' | 'no_shape' | 'no_text' | 'unsupported';
      message: string;
    };

/**
 * Pure validation for `add-shape`: resolve `params.shape` + `params.target.slideId` into a typed
 * host op, or a precise error code — before any host object is touched.
 */
function resolveAddShape(req: ActuationRequest): AddShapeResolution {
  const slideId = req.params.target?.slideId;
  if (!slideId) {
    return {
      ok: false,
      code: 'no_target',
      message: `add-shape needs target.slideId. ${SLIDE_TARGET_HINT}`,
    };
  }
  const shape = req.params.shape;
  if (!shape) {
    return { ok: false, code: 'no_shape', message: 'add-shape needs params.shape' };
  }
  const base: AddShapeGeometry = {
    slideId,
    fill: shape.fill,
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  };
  if (shape.shapeType === 'textBox') {
    const text = shape.text?.trim() ?? '';
    if (!text) {
      // Live: a malformed `/add-shape slide:2 {"text": …}` arrived with no text and was "applied"
      // as an empty box in the corner.
      return {
        ok: false,
        code: 'no_text',
        message:
          'add-shape shapeType=textBox needs text="…". Example: /add-shape slide=2 shapeType=textBox ' +
          'text="Draft" left=72 top=330 width=400 height=40',
      };
    }
    return { ok: true, op: { ...base, type: 'textBox', text: shape.text ?? '' } };
  }
  if (shape.shapeType === 'line') {
    return {
      ok: true,
      op: {
        ...base,
        type: 'line',
        connector: shape.connectorType ? CONNECTOR_TYPES[shape.connectorType] : undefined,
      },
    };
  }
  const geometry = GEOMETRIC_SHAPES.find((candidate) => candidate === shape.geometryType);
  if (!geometry) {
    return {
      ok: false,
      code: 'unsupported',
      message: `add-shape geometry "${shape.geometryType ?? ''}" is not supported.`,
    };
  }
  return { ok: true, op: { ...base, type: 'geometric', geometry } };
}

async function deckInsertOptions(
  ctx: PowerPoint.RequestContext,
  plan: ReturnType<typeof planInsertSlide>,
): Promise<PowerPoint.InsertSlideOptions | undefined> {
  const options: PowerPoint.InsertSlideOptions = {};
  if (plan.formatting !== undefined) options.formatting = plan.formatting;
  const targetSlideId = plan.targetSlideId ?? (await appendTargetSlideId(ctx, plan.targetIndex));
  if (targetSlideId !== undefined) options.targetSlideId = targetSlideId;
  return Object.keys(options).length === 0 ? undefined : options;
}

async function appendTargetSlideId(
  ctx: PowerPoint.RequestContext,
  targetIndex: number | undefined,
): Promise<string | undefined> {
  const slides = ctx.presentation.slides;
  const count = slides.getCount();
  await ctx.sync();
  if (count.value <= 0) return undefined;
  const index = targetIndex ?? count.value - 1;
  const slide = slides.getItemAt(index);
  slide.load('id');
  await ctx.sync();
  return slide.id;
}

interface PowerPointRevealTarget {
  slideId: string;
  shapeId?: string;
}

function powerpointRevealTarget(ref: ContextRef): PowerPointRevealTarget | undefined {
  if (ref.surface !== 'powerpoint' || !isSet('PowerPointApi', '1.5')) return undefined;
  return shapeRevealTarget(ref) ?? slideRevealTarget(ref);
}

function slideRevealTarget(ref: ContextRef): PowerPointRevealTarget | undefined {
  const id =
    prefixedValue(ref.id, 'pp:slide:', 'slide:') ?? prefixedValue(ref.anchor?.locator, 'slide:');
  return id ? { slideId: id } : undefined;
}

function shapeRevealTarget(ref: ContextRef): PowerPointRevealTarget | undefined {
  const fromId = prefixedValue(ref.id, 'pp:shape:', 'shape:');
  const fromLocator = prefixedValue(ref.anchor?.locator, 'pp:shape:', 'shape:');
  const raw = fromId ?? fromLocator;
  if (!raw) return undefined;
  const [slideId, shapeId] = raw.split(':');
  if (!slideId || !shapeId) return undefined;
  return { slideId, shapeId };
}

function prefixedValue(value: string | undefined, ...prefixes: string[]): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) return undefined;
  for (const prefix of prefixes) {
    if (trimmed.startsWith(prefix)) {
      const rest = trimmed.slice(prefix.length).trim();
      return rest || undefined;
    }
  }
  return undefined;
}

/**
 * Shape types whose `textFrame` is readable. Per the typings, `Shape.textFrame` throws
 * `InvalidArgument` for a shape without one (image, table, chart, group, media, ...), and a throw
 * fails the whole batched sync — so only these types are asked for text.
 */
const TEXT_SHAPE_TYPES: ReadonlySet<string> = new Set([
  'GeometricShape',
  'TextBox',
  'Placeholder',
  'Callout',
  'Freeform',
]);

/**
 * Shapes whose content is read: text-frame shapes, plus native tables (via `Table.values`,
 * PowerPointApi 1.8) where the host supports them. Pictures and charts have no text API.
 */
function isReadableShape(shape: PowerPoint.Shape): boolean {
  if (shape.type === 'Table') return isSet('PowerPointApi', '1.8');
  return TEXT_SHAPE_TYPES.has(shape.type);
}

/**
 * A placeholder reports `type: 'Placeholder'` whatever it holds; one filled with a picture, table
 * or chart has no text frame. `placeholderFormat.containedType` (PowerPointApi 1.8) is `null` for an
 * empty or text placeholder, else the contained shape's type.
 */
function placeholderHoldsText(containedType: string | null | undefined): boolean {
  return (
    containedType === null ||
    containedType === undefined ||
    (containedType !== 'Placeholder' && TEXT_SHAPE_TYPES.has(containedType))
  );
}

/** Read one slide's shapes' text (+ id/index) into a pure {@link SlideElement}. */
async function readSlide(
  ctx: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
): Promise<SlideElement> {
  const [element] = await readSlides(ctx, [slide]);
  if (!element) throw new Error('PowerPoint slide read returned no slide.');
  return element;
}

/**
 * Read slides' shape text in a constant number of host round-trips (at most three syncs, whatever
 * the slide count). Each sync is a full round-trip on PowerPoint for the web (~0.85s measured), so a
 * per-slide sync loop made a 9-slide whole-deck attach take ~16s; batched, the same read takes ~1.5s.
 */
async function readSlides(
  ctx: PowerPoint.RequestContext,
  slides: PowerPoint.Slide[],
): Promise<SlideElement[]> {
  if (slides.length === 0) return [];
  const shapeCollections = slides.map((slide) => {
    slide.load('id,index');
    const shapes = slide.shapes;
    shapes.load('items/id,items/type');
    return shapes;
  });
  await ctx.sync();

  const candidates = shapeCollections.flatMap((shapes) => shapes.items.filter(isReadableShape));
  const readable = await withoutNonTextPlaceholders(ctx, candidates);
  const texts = await readShapeContent(ctx, readable);

  return slides.map((slide, slideIndex) => {
    const shapes = shapeCollections[slideIndex]?.items ?? [];
    const shapeTexts = shapes.map((shape) => texts.get(shape) ?? '');
    const { title, body } = shapesToSlideText(shapeTexts);
    return {
      index: slide.index,
      slideId: slide.id,
      title,
      body,
      shapes: shapes.map((shape, index) => ({
        shapeId: shape.id,
        text: shapeTexts[index] ?? '',
      })),
    };
  });
}

/**
 * Drop placeholders holding a picture/table/chart (one sync, only when the batch has placeholders).
 * Without PowerPointApi 1.8, or if the host refuses the lookup, placeholders are kept and
 * {@link readShapeContent}'s per-shape fallback skips any that throw.
 */
async function withoutNonTextPlaceholders(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<PowerPoint.Shape[]> {
  const placeholders = shapes.filter((shape) => shape.type === 'Placeholder');
  if (placeholders.length === 0 || !isSet('PowerPointApi', '1.8')) return shapes;
  try {
    const formats = new Map(
      placeholders.map((shape) => {
        const format = shape.placeholderFormat;
        format.load('containedType');
        return [shape, format] as const;
      }),
    );
    await ctx.sync();
    return shapes.filter((shape) => {
      const format = formats.get(shape);
      return !format || placeholderHoldsText(format.containedType);
    });
  } catch {
    return shapes;
  }
}

/**
 * Queue a read of one shape's content: a table's cell grid, else its text frame. Returns a getter to
 * call after the next `ctx.sync()`.
 */
function queueShapeRead(shape: PowerPoint.Shape): () => string {
  if (shape.type === 'Table') {
    const table = shape.getTable();
    table.load('values');
    return () => tableValuesToText(table.values ?? []);
  }
  const range = shape.textFrame.textRange;
  range.load('text');
  return () => range.text ?? '';
}

/**
 * Read `shapes`' content (text frames and tables) in one sync. A single shape the host refuses
 * fails the whole batch, so on failure fall back to one sync per shape and skip the refusals —
 * slower, but one unexpected shape can't block the read of the rest of the deck.
 */
async function readShapeContent(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<Map<PowerPoint.Shape, string>> {
  const texts = new Map<PowerPoint.Shape, string>();
  if (shapes.length === 0) return texts;
  try {
    const reads = shapes.map((shape) => [shape, queueShapeRead(shape)] as const);
    await ctx.sync();
    for (const [shape, read] of reads) texts.set(shape, read());
    return texts;
  } catch {
    for (const shape of shapes) {
      try {
        const read = queueShapeRead(shape);
        await ctx.sync();
        texts.set(shape, read());
      } catch {
        // Nothing readable on this shape: leave it empty.
      }
    }
    return texts;
  }
}

async function readShapeContext(
  ctx: PowerPoint.RequestContext,
  target: PowerPointRevealTarget,
): Promise<ResolvedContext[]> {
  const slide = ctx.presentation.slides.getItem(target.slideId);
  slide.load('id,index');
  const shape = slide.shapes.getItemOrNullObject(target.shapeId ?? '');
  shape.load('id,isNullObject,type');
  await ctx.sync();
  if (shape.isNullObject || !target.shapeId || !isReadableShape(shape)) return [];
  const [readable] = await withoutNonTextPlaceholders(ctx, [shape]);
  if (!readable) return [];
  const texts = await readShapeContent(ctx, [readable]);
  return selectedShapeToContext(
    { index: slide.index, slideId: slide.id },
    { shapeId: shape.id, text: texts.get(readable) ?? '' },
  );
}

/** Placeholder types that take a slide title / its body text, in preference order. */
const TITLE_PLACEHOLDERS: readonly string[] = ['Title', 'CenterTitle', 'VerticalTitle'];
const BODY_PLACEHOLDERS: readonly string[] = ['Body', 'Content', 'Subtitle', 'VerticalBody'];

/**
 * Slide size used for fallback text boxes when the host can't report it (`pageSetup` is
 * PowerPointApi 1.10). 720 × 405 pt is the smallest standard slide (10" × 5.625"), so boxes laid
 * out for it stay on-slide on 16:9 widescreen (960 × 540) and 4:3 (720 × 540) decks too.
 */
const FALLBACK_SLIDE_SIZE = { width: 720, height: 405 };

/**
 * Write a composed title + bullets into a freshly added slide.
 *
 * `slides.add()` uses the deck's default layout, and a layout may have no title/body placeholders at
 * all (a "Blank" layout, or decks generated by tools such as pptxgenjs). So: write into the title and
 * body placeholders when the slide has them (matched by placeholder type on PowerPointApi 1.8, else
 * the first two placeholders in order); for any part with no placeholder, add a sized text box.
 * Every requested part is written somewhere — never a silent no-op.
 *
 * `shapes` (from a loaded probe) is only READ, for types and positions. Every WRITE goes through a
 * fresh, never-loaded `slides.getItemAt(slideIndex)` (and `shapes.getItemAt(i)`) proxy: PowerPoint
 * for the web rejects writes addressed through a just-added slide's provisional id (5010
 * "InvalidParam passed to GetItem(id)"), and a loaded proxy is addressed by id.
 */
async function writeComposedSlide(
  ctx: PowerPoint.RequestContext,
  slideIndex: number,
  shapes: PowerPoint.Shape[],
  title: string,
  bullets: string[],
): Promise<void> {
  const { titleShape, bodyShape } = await findTextPlaceholders(ctx, shapes);
  const needsBox = (title && !titleShape) || (bullets.length > 0 && !bodyShape);
  const size = needsBox ? await slideSize(ctx) : FALLBACK_SLIDE_SIZE;
  const margin = Math.round(size.width * 0.067);
  const width = size.width - margin * 2;
  const freshShapes = (): PowerPoint.ShapeCollection =>
    ctx.presentation.slides.getItemAt(slideIndex).shapes;
  const writeInto = (shape: PowerPoint.Shape, text: string): void => {
    freshShapes().getItemAt(shapes.indexOf(shape)).textFrame.textRange.text = text;
  };

  if (title) {
    if (titleShape) {
      writeInto(titleShape, title);
    } else {
      const box = freshShapes().addTextBox(title, {
        left: margin,
        top: Math.round(size.height * 0.07),
        width,
        height: Math.round(size.height * 0.18),
      });
      box.textFrame.textRange.font.size = 32;
      box.textFrame.textRange.font.bold = true;
    }
  }
  if (bullets.length > 0) {
    if (bodyShape) {
      // A body placeholder brings the layout's own bullet style.
      writeInto(bodyShape, bullets.join('\n'));
    } else {
      const box = freshShapes().addTextBox(bullets.map((bullet) => `• ${bullet}`).join('\n'), {
        left: margin,
        top: Math.round(size.height * 0.28),
        width,
        height: Math.round(size.height * 0.62),
      });
      box.textFrame.textRange.font.size = 18;
    }
  }
}

/** The new slide's title and body placeholders, if its layout has them. */
async function findTextPlaceholders(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<{ titleShape?: PowerPoint.Shape; bodyShape?: PowerPoint.Shape }> {
  const placeholders = shapes.filter((shape) => shape.type === 'Placeholder');
  if (placeholders.length === 0) return {};
  if (!isSet('PowerPointApi', '1.8')) {
    // No placeholder types to match on: title first, body second (the layout convention).
    return { titleShape: placeholders[0], bodyShape: placeholders[1] };
  }
  const formats = placeholders.map((shape) => {
    const format = shape.placeholderFormat;
    format.load('type');
    return [shape, format] as const;
  });
  await ctx.sync();
  const firstOf = (types: readonly string[]): PowerPoint.Shape | undefined => {
    for (const type of types) {
      const hit = formats.find(([, format]) => format.type === type);
      if (hit) return hit[0];
    }
    return undefined;
  };
  const titleShape = firstOf(TITLE_PLACEHOLDERS);
  const bodyShape = firstOf(BODY_PLACEHOLDERS);
  return { ...(titleShape ? { titleShape } : {}), ...(bodyShape ? { bodyShape } : {}) };
}

/** The deck's slide size in points, or {@link FALLBACK_SLIDE_SIZE} if the host can't report it. */
async function slideSize(
  ctx: PowerPoint.RequestContext,
): Promise<{ width: number; height: number }> {
  if (!isSet('PowerPointApi', '1.10')) return FALLBACK_SLIDE_SIZE;
  try {
    const setup = ctx.presentation.pageSetup;
    setup.load('slideWidth,slideHeight');
    await ctx.sync();
    if (setup.slideWidth > 0 && setup.slideHeight > 0) {
      return { width: setup.slideWidth, height: setup.slideHeight };
    }
  } catch {
    // Fall through to the size that fits every standard slide.
  }
  return FALLBACK_SLIDE_SIZE;
}

/**
 * Resolve the slide a write command names. Commands come from the model, which may give an exact
 * host id (`256#0` on PowerPoint web), `slide:<id>` / `pp:slide:<id>`, a 1-based slide number, or `last` for the
 * deck's last slide (e.g. the one a preceding `insert-slide` just added — the planner makes a
 * `slide=last` write depend on that insert, so it never runs if the insert failed). Resolved with one
 * read-only sync BEFORE any mutation is queued, so a bad reference fails cleanly with the valid
 * choices instead of surfacing as an "outcome uncertain" write.
 */
async function resolveSlideRef(
  ctx: PowerPoint.RequestContext,
  raw: string,
): Promise<{ slideId: string } | { error: string }> {
  const slides = ctx.presentation.slides;
  slides.load('items/id');
  await ctx.sync();
  const ids = slides.items.map((slide) => slide.id);
  const key = raw.trim().replace(/^(?:pp:)?slide:/i, '');
  if (/^last$/i.test(key)) {
    const last = ids.at(-1);
    return last ? { slideId: last } : { error: 'The deck has no slides yet.' };
  }
  if (ids.includes(key)) return { slideId: key };
  if (/^\d+$/.test(key)) {
    const id = ids[Number(key) - 1];
    if (id) return { slideId: id };
  }
  const shown = ids.slice(0, 12).join(', ') + (ids.length > 12 ? ', …' : '');
  return {
    error:
      `Slide ${JSON.stringify(raw.slice(0, 64))} was not found. Use a slide id from the outline (${shown}), ` +
      `a slide number 1–${ids.length}, or "last". To create a new slide with a title and bullets, ` +
      'use the slide command: slide "Title" "bullet" "bullet".',
  };
}

/**
 * Appended to a missing-slide error so the model can recover in one turn: models reach for the
 * `/add-*` commands to CREATE a slide with text; `slide "Title" "bullet"` is the command for that.
 */
const SLIDE_TARGET_HINT =
  'Pass slide= a slide id from the outline, a slide number, or last. To create a new slide with a title and ' +
  'bullets, use the slide command instead: slide "Title" "bullet" "bullet".';

/**
 * Append a slide and identify it. Finds the new slide's POSITION by the one id that is new (never a
 * pre-computed index — a co-author's concurrent add must not redirect later writes), then reads its
 * shapes through a probe proxy. The caller must WRITE through fresh, never-loaded
 * `slides.getItemAt(index)` proxies: PowerPoint for the web gives a just-added slide a provisional
 * id that rejects writes (5010 "InvalidParam passed to GetItem(id)"), and loading a proxy
 * re-addresses it by id. Queues the add — callers treat any failure after this as outcome-unknown.
 */
async function appendSlide(
  ctx: PowerPoint.RequestContext,
): Promise<
  { index: number; provisionalId: string; shapes: PowerPoint.Shape[] } | { error: string }
> {
  const slides = ctx.presentation.slides;
  slides.load('items/id');
  await ctx.sync();
  const existingIds = new Set(slides.items.map((item) => item.id));
  slides.add();
  await ctx.sync();
  const after = ctx.presentation.slides;
  after.load('items/id');
  await ctx.sync();
  const newIndexes = after.items.flatMap((item, index) =>
    existingIds.has(item.id) ? [] : [index],
  );
  const index = newIndexes.length === 1 ? newIndexes[0] : undefined;
  const provisionalId = index === undefined ? undefined : after.items[index]?.id;
  if (index === undefined || provisionalId === undefined) {
    return {
      error:
        'PowerPoint added a slide but it could not be identified. Inspect the deck before trying again.',
    };
  }
  const probe = ctx.presentation.slides.getItemAt(index);
  probe.load('id');
  const shapes = probe.shapes;
  shapes.load('items/id,items/type');
  await ctx.sync();
  if (probe.id !== provisionalId) {
    return {
      error:
        'The deck changed while the slide was being added. Inspect the deck before trying again.',
    };
  }
  return { index, provisionalId, shapes: shapes.items };
}

/**
 * Resolve the slide + shape a shape command names. Slides as in {@link resolveSlideRef}; shapes by
 * exact host id or `title` (the Title/CenterTitle placeholder, else the first shape with text).
 * Deliberately no shape ordinals (see below). Read-only, before any mutation; an unknown reference
 * returns the valid choices.
 */
async function resolveShapeTarget(
  ctx: PowerPoint.RequestContext,
  slideRef: string,
  shapeRef: string,
): Promise<{ slideId: string; shapeId: string } | { error: string }> {
  const slide = await resolveSlideRef(ctx, slideRef);
  if ('error' in slide) return slide;
  const collection = ctx.presentation.slides.getItem(slide.slideId).shapes;
  collection.load('items/id,items/type');
  await ctx.sync();
  const shapes = collection.items;
  const key = shapeRef.trim().replace(/^(?:pp:)?shape:/i, '');
  const exact = shapes.find((shape) => shape.id === key);
  if (exact) return { slideId: slide.slideId, shapeId: exact.id };
  if (/^title$/i.test(key)) {
    const title = await findTitleShape(ctx, shapes);
    if (title) return { slideId: slide.slideId, shapeId: title.id };
  }
  // No shape ordinals: PowerPoint shape ids are small numbers themselves (2, 3, 4…), so "2" would be
  // ambiguous between "id 2" and "the 2nd shape". Only exact ids and `title` resolve.
  const shown = await describeShapes(ctx, shapes.slice(0, 12));
  // Spell out both ref forms: the `shape` command takes `pp:shape:<slide>:<shape>`, the `/…-shape`
  // kinds take `shape=`. A model told only "shape=title" retries the `shape` command blind.
  const slideKey = slideRef
    .trim()
    .replace(/^(?:pp:)?slide:/i, '')
    .slice(0, 64);
  return {
    error:
      `Shape ${JSON.stringify(key.slice(0, 64))} was not found on that slide (there are no shape ` +
      `numbers). For the slide title use title as the shape: shape pp:shape:${slideKey}:title "…" ` +
      `(or shape=title on /format-shape). Otherwise use one of its shape ids: ` +
      `${shown || 'the slide has no shapes'}.`,
  };
}

/**
 * `id 86 (TextBox, title: "What is Tokenomics?"), id 85 (GeometricShape, no text)` — for a
 * not-found shape error. Ids and types alone left the model guessing: after "Shape 1" failed on a
 * deck whose first shape was a decorative line, it wrote the title into the line. Shape text is
 * document content: quoted, single-line and clipped, like the read listing. Read-only.
 */
async function describeShapes(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<string> {
  let texts = new Map<PowerPoint.Shape, string>();
  let title: PowerPoint.Shape | undefined;
  try {
    const readable = await withoutNonTextPlaceholders(ctx, shapes.filter(isReadableShape));
    texts = await readShapeContent(ctx, readable);
    title = await findTitleShape(ctx, shapes);
  } catch {
    // Fall back to ids and types.
  }
  return shapes
    .map((shape) => {
      const text = (texts.get(shape) ?? '').replace(/\s+/g, ' ').trim();
      const clipped = text.length > 40 ? `${text.slice(0, 40)}…` : text;
      const label = shape === title ? 'title: ' : '';
      return `id ${shape.id} (${shape.type}, ${text ? `${label}${JSON.stringify(clipped)}` : 'no text'})`;
    })
    .join(', ');
}

/**
 * A slide's title shape: its Title/CenterTitle placeholder, else the text shape set in the largest
 * type (see {@link largestTextShape}), else the first shape with text.
 */
async function findTitleShape(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
): Promise<PowerPoint.Shape | undefined> {
  const { titleShape } = await findTextPlaceholders(ctx, shapes);
  if (titleShape && isSet('PowerPointApi', '1.8')) return titleShape;
  const readable = await withoutNonTextPlaceholders(ctx, shapes.filter(isReadableShape));
  const texts = await readShapeContent(ctx, readable);
  const withText = readable.filter(
    (shape) => shape.type !== 'Table' && (texts.get(shape) ?? '').trim(),
  );
  if (withText.length <= 1) return withText[0];
  return (await largestTextShape(ctx, withText, texts)) ?? withText[0];
}

/**
 * The shape whose first visible character is set in the largest font (ties: the topmost). Decks
 * without title placeholders (Google Slides exports, hand-built slides) often put a small label
 * ("SECTION / 01") above the title, so the first shape with text was the label, not the title.
 * The first character, because a whole range with mixed runs reports `font.size` as null. One
 * read-only sync; undefined when the host reports no sizes.
 */
async function largestTextShape(
  ctx: PowerPoint.RequestContext,
  shapes: PowerPoint.Shape[],
  texts: Map<PowerPoint.Shape, string>,
): Promise<PowerPoint.Shape | undefined> {
  try {
    const probes = shapes.map((shape) => {
      const text = texts.get(shape) ?? '';
      const first = shape.textFrame.textRange.getSubstring(Math.max(0, text.search(/\S/)), 1);
      first.font.load('size');
      shape.load('top');
      return [shape, first] as const;
    });
    await ctx.sync();
    let best: { shape: PowerPoint.Shape; size: number; top: number } | undefined;
    for (const [shape, first] of probes) {
      const size = first.font.size ?? 0;
      const top = shape.top ?? Number.POSITIVE_INFINITY;
      if (size > 0 && (!best || size > best.size || (size === best.size && top < best.top))) {
        best = { shape, size, top };
      }
    }
    return best?.shape;
  } catch {
    return undefined;
  }
}

/** The deck's slide size when the host reports it (PowerPointApi 1.10), else undefined. */
async function knownSlideSize(
  ctx: PowerPoint.RequestContext,
): Promise<{ width: number; height: number } | undefined> {
  if (!isSet('PowerPointApi', '1.10')) return undefined;
  const size = await slideSize(ctx);
  return size === FALLBACK_SLIDE_SIZE ? undefined : size;
}

/**
 * `add-shape` needs explicit placement (the capability's contract: "never infer placement from
 * prose"). Without it the host drops the shape at the top-left corner — so "near the bottom" came
 * out at the top. The message gives the slide size so the next turn can place it.
 */
function missingGeometryMessage(
  size: { width: number; height: number } | undefined,
  geometry: { left?: number; top?: number; width?: number; height?: number },
): string | undefined {
  const missing = (['left', 'top', 'width', 'height'] as const).filter(
    (k) => geometry[k] === undefined,
  );
  if (missing.length === 0) return undefined;
  const slide = size
    ? `The slide is ${size.width} × ${size.height} pt (the bottom edge is top=${size.height}). `
    : '';
  return (
    `add-shape needs an explicit position: missing ${missing.join(', ')}. ${slide}` +
    'Give left= top= width= height= in points, e.g. near the bottom: ' +
    (size
      ? `left=${Math.round(size.width * 0.1)} top=${Math.round(size.height * 0.82)} width=${Math.round(size.width * 0.6)} height=${Math.round(size.height * 0.1)}.`
      : 'left=72 top=330 width=400 height=40.')
  );
}

/**
 * A corrective message when explicit geometry would place the shape outside the slide (models
 * guess a 960 × 540 slide on a 720 × 405 deck), so the next turn can fix it; undefined when it fits
 * or the size is unknown.
 */
function offSlideMessage(
  size: { width: number; height: number } | undefined,
  geometry: { left?: number; top?: number; width?: number; height?: number },
): string | undefined {
  if (!size) return undefined;
  const { left = 0, top = 0, width = 0, height = 0 } = geometry;
  if (left >= 0 && top >= 0 && left + width <= size.width && top + height <= size.height) {
    return undefined;
  }
  return (
    `That position is outside the slide, which is ${size.width} × ${size.height} pt: keep ` +
    `left + width ≤ ${size.width} and top + height ≤ ${size.height} ` +
    `(got left ${left}, top ${top}, width ${width}, height ${height}).`
  );
}

/** `" (PowerPoint: <code>)"` for an Office.js error, so an "uncertain" result says why. */
function hostErrorSuffix(error: unknown): string {
  const code = (error as { code?: unknown } | undefined)?.code;
  return typeof code === 'string' && /^[\w.-]{1,64}$/.test(code) ? ` (PowerPoint: ${code})` : '';
}

/** The id of the slide at `index`, read in a fresh request (undefined if it can't be read). */
async function readSlideIdAt(index: number): Promise<string | undefined> {
  try {
    return await PowerPoint.run(async (ctx) => {
      const slide = ctx.presentation.slides.getItemAt(index);
      slide.load('id');
      await ctx.sync();
      return slide.id;
    });
  } catch {
    return undefined;
  }
}

function slideNotFound(req: ActuationRequest, message: string): ActuationResult {
  return {
    ok: false,
    changeId: req.changeId,
    kind: req.kind,
    degraded: true,
    error: { code: 'target_conflict', message },
  };
}

/** Actual dispatch keys; conformance checks these against the advertised capabilities. */
export const HANDLED_ACTUATIONS: readonly ActuationKind[] = PowerPointBridge.handledActuations;
