import { describe, it, expect } from 'vitest';
import { ExcelBridge, coordinateToIndex, indexToCoordinate } from '../src';

describe('ExcelBridge', () => {
	describe('Coordinate Utilities', () => {
		it('should convert coordinate to index', () => {
			const result = coordinateToIndex('A1');
			expect(result).toEqual({ row: 0, col: 0 });
		});

		it('should convert index to coordinate', () => {
			const result = indexToCoordinate(0, 0);
			expect(result).toBe('A1');
		});

		it('should handle larger coordinates', () => {
			expect(coordinateToIndex('Z10')).toEqual({ row: 9, col: 25 });
			expect(indexToCoordinate(9, 25)).toBe('Z10');

			expect(coordinateToIndex('AA1')).toEqual({ row: 0, col: 26 });
			expect(indexToCoordinate(0, 26)).toBe('AA1');
		});
	});

	describe('Excel Writing', () => {
		it('should create Excel blob from simple data', () => {
			const data = [
				['Name', 'Age'],
				['John', 25],
				['Jane', 30],
			];

			const blob = ExcelBridge.write(data);

			expect(blob).toBeInstanceOf(Blob);
			expect(blob.type).toBe(
				'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			);
		});

		it('should create Excel buffer from simple data', () => {
			const data = [
				['Header 1', 'Header 2'],
				['Data 1', 'Data 2'],
			];

			const buffer = ExcelBridge.writeBuffer(data);

			expect(buffer).toBeInstanceOf(Uint8Array);
			expect(buffer.length).toBeGreaterThan(0);
		});
	});

	describe('Error Handling', () => {
		it('should throw error for invalid coordinate', () => {
			expect(() => coordinateToIndex('INVALID')).toThrow(
				'Invalid coordinate format',
			);
		});

		it('should handle empty data gracefully', () => {
			const blob = ExcelBridge.write([]);
			expect(blob).toBeInstanceOf(Blob);
		});
	});
});
