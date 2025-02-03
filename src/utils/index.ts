/* eslint-disable @typescript-eslint/no-explicit-any */

import { Worksheet } from "exceljs";

// 유틸리티 함수들
export const getDistinctYearlyPrefixes = (data: any[]): string[] => {
  const yearSet = new Set<string>();
  data.forEach((row) => {
    if (row.J) {
      const jValue = row.J.toString();
      const prefix = jValue.slice(0, 2);
      yearSet.add(`20${prefix}`);
    }
  });
  return Array.from(yearSet);
};

export const getDistinctMonthlyPrefixes = (data: any[]): string[] => {
  const monthlySet = new Set<string>();
  data.forEach((row) => {
    if (row.J) {
      const jValue = row.J.toString();
      const prefix = jValue.slice(0, 2);
      const suffix = jValue.slice(2, 4);
      monthlySet.add(`${prefix}년_${suffix}월`);
    }
  });
  return Array.from(monthlySet).sort((a, b) => {
    const [yearA, monthA] = a.split('년_').map(Number);
    const [yearB, monthB] = b.split('년_').map(Number);
    
    return yearA - yearB || monthA - monthB;
  });;
};


export const numberToColumnLetter = (num: number) => {
  let letter = "";
  while (num > 0) {
    const mod = (num - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    num = Math.floor((num - mod) / 26);
  }
  return letter;
};

export const autoAdjustColumnWidths = (worksheet: Worksheet) => {
  worksheet.columns.forEach((column: any) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell: any) => {
      let cellValue = cell.value;
      if (cellValue === null || cellValue === undefined) {
        cellValue = "";
      } else if (typeof cellValue === "object" && cellValue.richText) {
        cellValue = cellValue.richText.map((part: any) => part.text).join("");
      } else {
        cellValue = cellValue.toString();
      }
      maxLength = Math.max(maxLength, cellValue.length);
    });
    column.width = maxLength < 10 ? 10 : maxLength + 2;
  });
};
