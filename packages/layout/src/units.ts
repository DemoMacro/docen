// Unit conversions shared by the adapters and the engine. The LayoutDoc
// projection is all-px: an adapter converts its document units (OOXML
// twips/half-points/EMU/eighths-of-a-point) through these helpers exactly
// once, so the engine never sees a second unit system.

// 1in = 25.4mm = 72pt = 96px → 1pt = 4/3 px. An OOXML twip is 1/20 pt.
export const PT_TO_PX = 4 / 3;
export const TWIP_TO_PX = 4 / 3 / 20;

// 1 px = 9525 EMU (914400 EMU/inch ÷ 96 px/inch). Drawing offsets/margins are EMU.
export const EMU_PER_PX = 9525;

export const ptToPx = (pt: number): number => pt * PT_TO_PX;
export const twipToPx = (twip: number): number => twip * TWIP_TO_PX;
export const emuToPx = (emu: number): number => emu / EMU_PER_PX;
