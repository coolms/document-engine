/**
 * `@coolms/document-engine` — the public surface.
 *
 * The chain, in the order a document travels it:
 *
 *   bytes -> OpcPackage -> XmlDocument -> StyleSheet + FontCatalogue
 *         -> readWordDocument -> breakIntoLines -> layoutPages -> pages
 *
 * Every step is usable on its own. The editor needs the whole chain; a server
 * that only rewrites a template needs the first two and nothing else.
 */

// Fonts and measurement.
export { TrueTypeFont } from './font/truetype-font.js';
export type { MeasuredText, VerticalMetrics } from './font/truetype-font.js';
export { GposKerning } from './font/gpos-kerning.js';

// Line breaking and pagination.
export { breakIntoLines } from './layout/line-breaker.js';
export type { BreakOptions, FieldKind, Strike, Underline, UnderlineStyle, Line, LinePiece, LineWidth, StyledRun, TabStops } from './layout/line-breaker.js';
export { fieldText, hasField, resolveFields } from './layout/fields.js';
export { formatNumeral } from './text/numerals.js';
export type { NumeralStyle } from './text/numerals.js';
export type { FieldContext } from './layout/fields.js';
export { OBJECT_REPLACEMENT } from './layout/image.js';
export type { ImageContent, InlineImage } from './layout/image.js';
export { alignLine } from './layout/alignment.js';
export { resolveCellBorders } from './layout/borders.js';
export type { BorderSide, BorderStyle, BoxBorders, CellPosition } from './layout/borders.js';
export { baselineOffsetPx, EXACT_BASELINE_RATIO } from './layout/baseline.js';
export type { BaselineMetrics } from './layout/baseline.js';
export type { Alignment } from './layout/alignment.js';
export { DEFAULT_PARAGRAPH_STYLE, DEFAULT_TAB_PX, isTable, layoutPages, layoutSections, stackBlocks } from './layout/page-layout.js';
export type {
    Block,
    ContentBox,
    FurnitureSource,
    LayoutOptions,
    ListMarker,
    Page,
    PageColumn,
    PageGeometry,
    Paragraph,
    ParagraphStyle,
    PlacedCell,
    PlacedFurniture,
    PlacedLine,
    PlacedRow,
    Section,
    CellVerticalAlign,
    RowHeightRule,
    Table,
    TableCell,
    TableRow,
    VerticalMerge,
} from './layout/page-layout.js';

// XML surgery and the units OOXML measures in.
export { XmlDocument, XmlElement, XmlText, XmlVerbatim, escapeAttribute, escapeText, unescapeXml } from './ooxml/xml.js';
export type { XmlAttribute, XmlNode } from './ooxml/xml.js';
export {
    EMU_PER_INCH,
    PX_PER_INCH,
    POINTS_PER_INCH,
    TWIPS_PER_INCH,
    eighthPointsToPoints,
    eighthPointsToPx,
    emuToPoints,
    emuToPx,
    halfPointsToPoints,
    halfPointsToPx,
    pointsToEmu,
    pointsToHalfPoints,
    pointsToPx,
    pointsToTwips,
    pxToEmu,
    pxToHalfPoints,
    pxToPoints,
    pxToTwips,
    twipsToPoints,
    twipsToPx,
} from './ooxml/units.js';

// The OPC container.
export { OpcPackage, webStreamsCodec } from './opc/opc-package.js';
export type { ZipCodec } from './opc/opc-package.js';
export { crc32 } from './opc/crc32.js';
export { findById, findByType, readRelationships, relationshipPartFor, resolveTarget } from './opc/relationships.js';
export type { Relationship } from './opc/relationships.js';

// Rich text — an editor's document shape rather than a .docx's.
export { DEFAULT_FLOW_STYLE, isFlowTable, paginateFlow } from './richtext/flow.js';
export type {
    BlockStyle,
    FlowBlock,
    FlowItem,
    FlowOptions,
    FlowPageStart,
    FlowPagination,
    FlowStyle,
    FlowTable,
    FlowTableCell,
    FlowTableRow,
    TextSpan,
} from './richtext/flow.js';

// WordprocessingML.
export { FONT_MANIFEST_ASSET, FontCatalogue, fontStyleKey, offeredFontFamilies } from './word/font-catalogue.js';
export type { FontManifest, FontStyleKey, ResolvedFont } from './word/font-catalogue.js';
export { StyleSheet, readParagraphProperties, readRunProperties } from './word/style-sheet.js';
export type { LineRule, ParagraphProperties, RunProperties, StyleNumbering } from './word/style-sheet.js';
export { Numbering, NumberingCounters } from './word/numbering.js';
export type { NumberFormat, NumberingLevel } from './word/numbering.js';
export { readWordDocument } from './word/document-reader.js';
export type {
    Diagnostic,
    DiagnosticKind,
    DocumentSection,
    FurnitureParts,
    FurnitureVariant,
    PageFurniture,
    ReadOptions,
    WordDocument,
} from './word/document-reader.js';
export { openWordFile } from './word/word-package.js';
export type { OpenedWordFile } from './word/word-package.js';

// Drawing.
export { renderPage } from './render/display-list.js';
export type { DrawOp, ImageOp, LineOp, RectOp, RenderedPage, TextOp } from './render/display-list.js';
export { renderPageToSvg, serialise } from './render/svg.js';
export type { SvgOptions } from './render/svg.js';
