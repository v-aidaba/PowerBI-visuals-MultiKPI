/**
 *  Power BI Visualizations
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */
import { IDataRepresentationPoint } from "../converter/data/dataRepresentation";

export interface IDataGapResult {
    hasGaps: boolean;
    totalMissingDays: number;
    gaps: { startDate: Date; endDate: Date; missingDays: number }[];
}


export class DataGapDetector {
    
    public static detectGaps(points: IDataRepresentationPoint[]): IDataGapResult {
        const result: IDataGapResult = { hasGaps: false, totalMissingDays: 0, gaps: [] };
        if (!points || points.length < 2) return result;

        const sorted = [...points].sort((a, b) => a.x.getTime() - b.x.getTime());
        const valid = sorted.filter(p => this.isValid(p));
        if (valid.length < 2) return result;

        let currentGapStart: Date | null = null;
        let gapDays = 0;

        for (let i = 1; i < sorted.length; i++) {
            const prev = sorted[i - 1];
            const curr = sorted[i];
            const isPrevValid = this.isValid(prev);
            const isCurrValid = this.isValid(curr);

            const days = this.daysBetween(prev.x, curr.x);

            if (!isPrevValid || !isCurrValid || days > 1) {
                if (!currentGapStart) currentGapStart = prev.x;
                gapDays += days > 1 ? days - 1 : 1;
            } else if (currentGapStart) {
                result.gaps.push({
                    startDate: new Date(currentGapStart.getTime() + 86400000),
                    endDate: new Date(curr.x.getTime() - 86400000),
                    missingDays: gapDays
                });
                result.totalMissingDays += gapDays;
                currentGapStart = null;
                gapDays = 0;
            }
        }

        if (currentGapStart && gapDays > 0) {
            const last = sorted[sorted.length - 1];
            result.gaps.push({ startDate: currentGapStart, endDate: last.x, missingDays: gapDays });
            result.totalMissingDays += gapDays;
        }

        result.hasGaps = result.gaps.length > 0;
        return result;
    }

    private static isValid(p: IDataRepresentationPoint): boolean {
        const y = p?.y;
        return y != null && isFinite(y) && !isNaN(y);
    }

    private static daysBetween(a: Date, b: Date): number {
        return Math.ceil((b.getTime() - a.getTime()) / 86400000);
    }

    public static formatGapMessage(template: string, totalMissingDays: number): string {
        return (template || `⚠️ ${totalMissingDays} missing days detected`)
            .replace("${1}", totalMissingDays.toString());
    }
}