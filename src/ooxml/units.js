// Office Open XML's units in CSS pixels (96 per inch): drawings measure in EMUs (914,400 an inch), most of
// WordprocessingML in twips (twentieths of a point), font sizes in half-points.

/** EMUs in a pixel. */
export const EMU_PER_PX = 9525;
/** Twips in a pixel. */
export const TWIPS_PER_PX = 15;

/** EMUs (a number or an attribute's string) to pixels, unrounded. */
export const emuToPx = (v) => Number(v) / EMU_PER_PX;
/** Pixels to whole EMUs. */
export const pxToEmu = (px) => Math.round(px * EMU_PER_PX);
/** Twips (a number or an attribute's string) to pixels, unrounded. */
export const twipsToPx = (v) => Number(v) / TWIPS_PER_PX;
