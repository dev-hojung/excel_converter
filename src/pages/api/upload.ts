/* eslint-disable @typescript-eslint/no-explicit-any */
// pages/api/upload.ts
import { NextApiRequest, NextApiResponse } from 'next';
import { formidable } from 'formidable'; // <-- named import
import fs from 'fs';
import { Workbook, Worksheet } from 'exceljs';

export const config = {
  api: {
    bodyParser: false,
  },
};

interface MyRowData {
  [key: string]: any;
  A?: string;
  B?: string;
  C?: string;
  D?: string;
  E?: string;
  F?: Date;
  G?: string;
  H?: number;
  I?: number;
  J?: string;
  K?: string;
}

export default function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== 'POST') {
    return res.status(405).json({ message: 'Method not allowed' });
  }

  // v3 이상: 함수 방식
  const form = formidable(); 

  form.parse(req, async (err, fields, files) => {
    if (err) {
      console.error('Form parse error:', err);
      return res
        .status(500)
        .json({ message: '파일 업로드 중 오류가 발생했습니다.' });
    }

    const file = files.excel;
    if (!file || !Array.isArray(file)) {
      return res
        .status(400)
        .json({ message: '엑셀 파일이 제공되지 않았습니다.' });
    }

    try {
      const filePath = file[0].filepath;

      // exceljs로 Workbook 읽기
      const workbook = new Workbook();
      await workbook.xlsx.readFile(filePath);

      const worksheet = workbook.getWorksheet(1);
      if (!worksheet) {
        return res
          .status(400)
          .json({ message: '워크시트를 찾을 수 없습니다.' });
      }

      // 데이터 수집/가공
      const originalData: MyRowData[] = [];
      worksheet.eachRow((row) => {
        const rowObj: MyRowData = {};
        row.eachCell((cell, colNumber) => {
          const colName = worksheet.getColumn(colNumber).letter;
          rowObj[colName] = cell.value;
        });
        originalData.push(rowObj);
      });

      const dataWithoutHeader = originalData.slice(1);
      const distinctYears = getDistinctYearPrefixes(dataWithoutHeader);

      const code = dataWithoutHeader.map(item => item.C);

      const uniqueCode = [...new Set(code)];

     // uniqueModel 배열을 순회하여, 각 모델에 해당하는 데이터를 그룹화 및 집계
      const aggregatedResults = uniqueCode.map(code => {
        // 1. 해당 모델(C값)과 일치하는 데이터만 필터링
        const filtered = dataWithoutHeader.filter(item => item.C === code);
        const A = filtered[0].D;
        const B = code;
        
        const startChar = 'C';
        const startCode = startChar.charCodeAt(0);
        const endCode = String.fromCharCode(startCode + distinctYears.length);
        const result: {[key in string]: any} = {};

        distinctYears.forEach((value, index) => {
          // 시작 문자 'C'의 코드에 인덱스를 더하여 다음 알파벳을 구함
          const key = String.fromCharCode(startCode + index);
          result[key] = filtered.filter(item => `20${item.J?.substring(0, 2)}` === value).reduce((acc, arr) => acc + (arr.I ? arr.I : 0), 0) ;

        });

        result[endCode] = Object.values(result).reduce((acc, curr) => acc + curr, 0);

        return {
          A,
          B,
          ...result, 
        }
      }).sort((a, b) => String(a.A).localeCompare(String(b.A)));

      const newWorkbook = new Workbook();
      const newWorksheet = newWorkbook.addWorksheet();
      const header = ['제품군', '모델명', ...distinctYears, '총 합계'];
      newWorksheet.addRow(header);
      
      // 데이터 행과 그룹별 요약 행을 추가하는 로직
      // (A의 값이 바뀌는 시점에 이전 그룹의 합계를 구하는 요약 행을 삽입)

      let currentGroup: string | null | undefined = null;       // 현재 그룹(제품군)
      let groupStartRow = newWorksheet.rowCount + 1; // 그룹 시작 데이터 행 번호 (헤더 이후, 즉 2부터)

      // aggregatedResults 배열 순회 (이미 A 기준으로 정렬되었다고 가정)
      aggregatedResults.forEach((dataRow) => {
        // 만약 처음이거나 현재 그룹과 dataRow.A가 동일하다면 그냥 데이터 행 추가
        if (currentGroup === null) {
          currentGroup = dataRow.A;
        } else if (dataRow.A !== currentGroup) {
          // 제품군(A)이 바뀌었으므로, 이전 그룹의 데이터를 합산하는 요약 행 추가

          // 현재 워크시트의 마지막 행 번호 + 1 (새로 추가할 요약 행의 행 번호)
          const currentExcelRow = newWorksheet.rowCount + 1;
          // 요약 행의 첫 두 셀은 제품군과 '합계'라는 표시
          const summaryRow = newWorksheet.addRow([ `${currentGroup} 계` ]);
          
          // 총 열 수는 header 배열의 길이
          const totalColumns = header.length;
          // 데이터가 들어가는 숫자 셀은 3번째 열부터 마지막 열까지 (예, C, D, E)
          for (let col = 3; col <= totalColumns; col++) {
            const colLetter = numberToColumnLetter(col);
            // 그룹의 시작 행(groupStartRow)부터 (요약 행 바로 위)까지 합계를 구하는 수식
            const formula = `SUM(${colLetter}${groupStartRow}:${colLetter}${currentExcelRow - 1})`;
            summaryRow.getCell(col).value = { formula };
          }

          newWorksheet.mergeCells(`A${summaryRow.number}:B${summaryRow.number}`);
          summaryRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
          summaryRow.eachCell((cell) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFFFE599' },
            };
            cell.border = {
              top: { style: 'thin', color: { argb: 'FF000000' } },
              left: { style: 'thin', color: { argb: 'FF000000' } },
              bottom: { style: 'thin', color: { argb: 'FF000000' } },
              right: { style: 'thin', color: { argb: 'FF000000' } },
            };
            cell.font = { bold: true };
          });
          
          // 그룹이 바뀌었으므로, 새로운 그룹 설정 및 그룹 시작 행 갱신
          currentGroup = dataRow.A;
          groupStartRow = newWorksheet.rowCount + 1; // 요약 행이 추가되었으므로 다음 행부터 새 그룹 시작
        }
        
        // 현재 dataRow를 데이터 행으로 추가  
        // 여기서는 aggregatedResults의 각 행이 [제품군, 모델명, C, D, E] 순서라고 가정함
        newWorksheet.addRow(Object.values(dataRow));
      });

      // 마지막 그룹에 대해서도 요약 행 추가 (반복문 종료 후)
      if (currentGroup !== null) {
        const currentExcelRow = newWorksheet.rowCount + 1;
        const summaryRow = newWorksheet.addRow([ `${currentGroup} 계` ]);
        const totalColumns = header.length;
        for (let col = 3; col <= totalColumns; col++) {
          const colLetter = numberToColumnLetter(col);
          const formula = `SUM(${colLetter}${groupStartRow}:${colLetter}${currentExcelRow - 1})`;
          summaryRow.getCell(col).value = { formula };
        }
        newWorksheet.mergeCells(`A${summaryRow.number}:B${summaryRow.number}`);
        summaryRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
        summaryRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFE599' },
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } },
          };
          cell.font = { bold: true };
        });
      }

      autoAdjustColumnWidths(newWorksheet);

      const buffer = await newWorkbook.xlsx.writeBuffer();
      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      res.setHeader('Content-Disposition', 'attachment; filename=custom.xlsx');

      return res.status(200).send(buffer);
    } catch (error) {
      console.error('Excel processing error:', error);
      return res
        .status(500)
        .json({ message: '엑셀 처리 중 오류가 발생했습니다.' });
    } finally {
      // 임시 파일 삭제
      if (file[0].filepath && fs.existsSync(file[0].filepath)) {
        fs.unlinkSync(file[0].filepath);
      }
    }
  });
}

function getDistinctYearPrefixes(data: any[]): string[] {
  // 중복 제거를 위해 Set 사용
  const yearSet = new Set<string>();

  data.forEach((row) => {
    if (row.J) {
      // J가 문자열이든 숫자든, 일단 문자열로 처리
      const jValue = row.J.toString();
      // 앞 2글자 추출
      const prefix = jValue.slice(0, 2);
      // Set에 추가 (자동으로 중복 제외)
      yearSet.add(`20${prefix}`);
    }
  });

  // Set을 배열로 변환
  return Array.from(yearSet);
}

function numberToColumnLetter(num: number) {
  let letter = '';
  while (num > 0) {
    const mod = (num - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    num = Math.floor((num - mod) / 26);
  }
  return letter;
}

function autoAdjustColumnWidths(worksheet: Worksheet) {
  worksheet.columns.forEach((column: any) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell: any) => {
      let cellValue = cell.value;
      if (cellValue === null || cellValue === undefined) {
        cellValue = '';
      } else if (typeof cellValue === 'object' && cellValue.richText) {
        // richText인 경우 텍스트 부분만 추출
        cellValue = cellValue.richText.map((part: any) => part.text).join('');
      } else {
        cellValue = cellValue.toString();
      }
      // 최대 길이 갱신 (필요시 trim() 처리 가능)
      maxLength = Math.max(maxLength, cellValue.length);
    });
    // 최소 너비를 10으로, 최대 길이에 약간의 여유를 주어 설정 (여유값은 조절 가능)
    column.width = maxLength < 10 ? 10 : maxLength + 2;
  });
}