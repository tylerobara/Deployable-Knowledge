import type { ExtractedChunk, ExtractionResult, Source } from './parse-shared';
import { rowsToTablePages } from './tabular-extract';

const MAX_SPREADSHEET_CELLS = 200_000;

export async function extractSpreadsheet(
	source: Source,
	onProgress?: (ratio: number, message: string) => void
): Promise<ExtractionResult> {
	const xlsx = (await import('node-xlsx')).default;
	const sheets = xlsx.parse(source.path);
	const chunks: ExtractedChunk[] = [];
	let cellCount = 0;

	sheets.forEach((sheet, sheetIndex) => {
		onProgress?.(
			(sheetIndex + 1) / Math.max(sheets.length, 1),
			`Reading sheet ${sheetIndex + 1} of ${sheets.length}`
		);

		const rows: string[][] = sheet.data.map((row) =>
			row.map((cell) => (cell == null ? '' : String(cell)))
		);

		for (const cells of rows) {
			cellCount += cells.length;
			if (cellCount > MAX_SPREADSHEET_CELLS) {
				throw new Error(
					`This spreadsheet has more than ${MAX_SPREADSHEET_CELLS.toLocaleString('en-US')} cells and cannot be ingested.`
				);
			}
		}

		chunks.push(...rowsToTablePages(source, rows, sheetIndex, sheet.name));
	});

	return { chunks, pageCount: sheets.length };
}
