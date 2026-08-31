/** Every GDI+ object payload in these files repeats this version stamp. */
export const GDIPLUS_VERSION = 0xdbc01002;

// ── record codes replayed ──

export const PLUS_END_OF_FILE = 0x4002;
export const PLUS_OBJECT = 0x4008;
export const PLUS_FILL_RECTS = 0x400a;
export const PLUS_DRAW_LINES = 0x400d;
export const PLUS_FILL_PATH = 0x4014;
export const PLUS_DRAW_PATH = 0x4015;
export const PLUS_DRAW_IMAGE_POINTS = 0x401b;
export const PLUS_SAVE = 0x4025;
export const PLUS_RESTORE = 0x4026;
export const PLUS_SET_WORLD_TRANSFORM = 0x402a;
export const PLUS_RESET_WORLD_TRANSFORM = 0x402b;
export const PLUS_MULTIPLY_WORLD_TRANSFORM = 0x402c;

export const OBJ_BRUSH = 1;
export const OBJ_PEN = 2;
export const OBJ_PATH = 3;
export const OBJ_IMAGE = 5;

// Carrier EMR codes consumed while replaying the text chain ([MS-EMF]
// RecordType values).
export const EMR_SAVEDC = 33;
export const EMR_RESTOREDC = 34;
export const EMR_SET_WORLD_TRANSFORM = 35;
export const EMR_MODIFY_WORLD_TRANSFORM = 36;
export const EMR_SELECT_OBJECT = 37;
export const EMR_CREATE_PEN = 38;
export const EMR_CREATE_BRUSH_INDIRECT = 39;
export const EMR_DELETE_OBJECT = 40;
export const EMR_SETTEXTCOLOR = 24;
export const EMR_SETTEXTALIGN = 22;
export const EMR_BITBLT = 76;
export const EMR_STRETCHDIBITS = 81;
export const EMR_EXT_CREATE_FONT = 82;
export const EMR_EXT_TEXT_OUT_A = 83;
export const EMR_EXT_TEXT_OUT_W = 84;
export const EMR_POLYLINE16 = 87;
export const EMR_EXT_CREATE_PEN = 95;
// Path construction and consumption: a figure opened with BeginPath grows
// through MoveTo/LineTo/PolylineTo16/PolyBezierTo16 (+ CloseFigure), is
// frozen by EndPath, and paints via FillPath/StrokePath/StrokeAndFillPath.
export const EMR_MOVE_TO_EX = 27;
export const EMR_BEGIN_PATH = 58;
export const EMR_END_PATH = 60;
export const EMR_CLOSE_FIGURE = 61;
export const EMR_FILL_PATH = 62;
export const EMR_STROKE_AND_FILL_PATH = 63;
export const EMR_STROKE_PATH = 64;
export const EMR_ABORT_PATH = 68;
export const EMR_LINE_TO = 54;
export const EMR_POLYGON16 = 86;
export const EMR_POLYBEZIERTO16 = 88;
export const EMR_POLYLINETO16 = 89;

/** EmrText offsets inside EMR_EXTTEXTOUTW (relative to the record start):
 *  bounds[16], graphicsMode, ex/eyScale, then Reference xy, Chars, offString,
 *  Options (+ an optional rect). The string lives at offString, relative to
 *  the record start. */
export const EXT_REF_X = 36;
export const EXT_REF_Y = 40;
export const EXT_CHARS = 44;
export const EXT_OFF_STRING = 48;
