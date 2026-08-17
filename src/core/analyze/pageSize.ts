import { twipsToInches } from '../ooxml/ns.js';

interface KnownSize {
  label: string;
  widthIn: number;
  heightIn: number;
}

/** Trim sizes a book interior is likely to use, KDP's list plus office paper. */
const KNOWN_SIZES: KnownSize[] = [
  { label: 'US Letter', widthIn: 8.5, heightIn: 11 },
  { label: 'A4', widthIn: 8.27, heightIn: 11.69 },
  { label: 'A5', widthIn: 5.83, heightIn: 8.27 },
  { label: 'US Legal', widthIn: 8.5, heightIn: 14 },
  { label: 'Mass market', widthIn: 4.25, heightIn: 6.87 },
  { label: 'Pocket', widthIn: 5, heightIn: 8 },
  { label: 'Digest', widthIn: 5.5, heightIn: 8.5 },
  { label: 'Trade paperback', widthIn: 6, heightIn: 9 },
  { label: 'US Trade', widthIn: 6.14, heightIn: 9.21 },
  { label: 'Royal', widthIn: 6.69, heightIn: 9.61 },
  { label: 'Executive', widthIn: 7, heightIn: 10 },
  { label: 'Crown quarto', widthIn: 7.44, heightIn: 9.69 },
  { label: 'Small square', widthIn: 8.5, heightIn: 8.5 },
  { label: 'Novella', widthIn: 5, heightIn: 7 },
  { label: 'Comic', widthIn: 6.63, heightIn: 10.25 },
  { label: 'KDP 5.06 x 7.81', widthIn: 5.06, heightIn: 7.81 },
  { label: 'KDP 5.25 x 8', widthIn: 5.25, heightIn: 8 },
  { label: 'KDP 7.5 x 9.25', widthIn: 7.5, heightIn: 9.25 },
  { label: 'KDP 8 x 10', widthIn: 8, heightIn: 10 },
];

const TOLERANCE_IN = 0.03;

/** Format twips as inches with the fewest digits that stay accurate. */
export function inchLabel(twips: number): string {
  const inches = twipsToInches(twips);
  const rounded = Math.round(inches * 100) / 100;
  return String(rounded).replace(/\.?0+$/, '') || '0';
}

/** e.g. `5.5" x 8.5" (Digest)`; falls back to plain dimensions. */
export function describePageSize(widthTwips: number, heightTwips: number): string {
  const w = twipsToInches(widthTwips);
  const h = twipsToInches(heightTwips);
  const dims = `${inchLabel(widthTwips)}" x ${inchLabel(heightTwips)}"`;
  const known = KNOWN_SIZES.find(
    (s) => Math.abs(s.widthIn - w) <= TOLERANCE_IN && Math.abs(s.heightIn - h) <= TOLERANCE_IN,
  );
  return known ? `${dims} (${known.label})` : dims;
}
