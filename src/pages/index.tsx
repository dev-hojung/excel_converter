/* eslint-disable @typescript-eslint/no-explicit-any */
import React, { useState, DragEvent } from 'react';

import ExcelJS, { Worksheet } from 'exceljs';

export default function UploadPage() {
  const [file, setFile] = useState<File | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string>('');
  const [isLoading, setIsLoading] = useState(false); // 로딩 상태

  // 파일 직접 선택 시
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0]);
    }
    e.target.value = '';
  };

  // 드래그된 파일을 영역에 놓을 때
  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      setFile(e.dataTransfer.files[0]);
    }
  };

  // 드래그가 영역 위를 지날 때 (기본 이벤트 막아야 drop 가능)
  const handleDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  // 파일 제거
  const handleRemoveFile = () => {
    setFile(null);
    setDownloadUrl('');
  };


const processFile = async (file: File) => {
  if (!file) {
    alert('파일을 먼저 선택(또는 드래그)하세요.');
    return;
  }

  setIsLoading(true);
  try {
    // FileReader 대신 file.arrayBuffer() 사용 (최신 브라우저 지원)
    const arrayBuffer = await file.arrayBuffer();

    // 기존 엑셀 파일 읽기
    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.load(arrayBuffer);

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      alert('워크시트를 찾을 수 없습니다.');
      return;
    }

    // 데이터 수집/가공 (헤더 제외)
    const originalData: any[] = [];
    worksheet.eachRow((row) => {
      const rowObj: any = {};
      row.eachCell((cell, colNumber) => {
        const colName = worksheet.getColumn(colNumber).letter;
        rowObj[colName] = cell.value;
      });
      originalData.push(rowObj);
    });
    const dataWithoutHeader = originalData.slice(1);
    const distinctYears = getDistinctYearPrefixes(dataWithoutHeader);
    const codeArr = dataWithoutHeader.map((item) => item.C);
    const uniqueCode = [...new Set(codeArr)];

    // uniqueModel 배열을 순회하여 그룹별 집계 (정렬된 aggregatedResults 생성)
    const aggregatedResults = uniqueCode
      .map((code) => {
        const filtered = dataWithoutHeader.filter((item) => item.C === code);
        const A = filtered[0].D;
        const B = code;
        const startChar = 'C';
        const startCode = startChar.charCodeAt(0);
        const endCode = String.fromCharCode(startCode + distinctYears.length);
        const result: { [key: string]: any } = {};

        distinctYears.forEach((value, index) => {
          const key = String.fromCharCode(startCode + index);
          result[key] = filtered
            .filter((item) => `20${item.J?.substring(0, 2)}` === value)
            .reduce((acc, curr) => acc + (curr.I ? curr.I : 0), 0);
        });

        result[endCode] = Object.values(result).reduce(
          (acc, curr) => acc + curr,
          0
        );

        return { A, B, ...result };
      })
      .sort((a, b) => String(a.A).localeCompare(String(b.A)));

    // 새 워크북 생성 (변환 결과)
    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet('Converted');
    const header = ['제품군', '모델명', ...distinctYears, '총 합계'];
    newWorksheet.addRow(header).eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'd8eef2' },
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } },
      };
      cell.font = { bold: true };
    })

    // 그룹별 데이터 행 및 요약 행 추가 (요약 행은 그룹이 바뀌는 시점에 추가)
    let currentGroup: string | null = null;
    let groupStartRow = newWorksheet.rowCount + 1; // 데이터 시작 행 번호 (헤더 바로 아래)

    aggregatedResults.forEach((dataRow) => {
      if (currentGroup === null) {
        currentGroup = dataRow.A;
      } else if (dataRow.A !== currentGroup) {
        // 그룹이 바뀌었으므로 이전 그룹의 요약 행 추가
        const currentExcelRow = newWorksheet.rowCount + 1;
        const summaryRow = newWorksheet.addRow([`${currentGroup} 계`]);
        const totalColumns = header.length;
        for (let col = 3; col <= totalColumns; col++) {
          const colLetter = numberToColumnLetter(col);
          const formula = `SUM(${colLetter}${groupStartRow}:${colLetter}${
            currentExcelRow - 1
          })`;
          summaryRow.getCell(col).value = { formula };
        }
        // 병합 및 스타일 적용
        newWorksheet.mergeCells(`A${summaryRow.number}:B${summaryRow.number}`);
        summaryRow.getCell(1).alignment = {
          vertical: 'middle',
          horizontal: 'center',
        };
        summaryRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'ffe9da' },
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } },
          };
          cell.font = { bold: true };
        });
        currentGroup = dataRow.A;
        groupStartRow = newWorksheet.rowCount + 1;
      }

      // 데이터 행 추가 (객체의 순서를 직접 구성)
      newWorksheet.addRow(Object.values(dataRow));
    });

    // 마지막 그룹 요약 행 추가
    if (currentGroup !== null) {
      const currentExcelRow = newWorksheet.rowCount + 1;
      const summaryRow = newWorksheet.addRow([`${currentGroup} 계`]);
      const totalColumns = header.length;
      for (let col = 3; col <= totalColumns; col++) {
        const colLetter = numberToColumnLetter(col);
        const formula = `SUM(${colLetter}${groupStartRow}:${colLetter}${
          currentExcelRow - 1
        })`;
        summaryRow.getCell(col).value = { formula };
      }
      newWorksheet.mergeCells(`A${summaryRow.number}:B${summaryRow.number}`);
      summaryRow.getCell(1).alignment = {
        vertical: 'middle',
        horizontal: 'center',
      };
      summaryRow.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'ffe9da' },
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

    // 엑셀 파일을 버퍼로 생성 후 Blob URL 생성
    const outputBuffer = await newWorkbook.xlsx.writeBuffer();
    const blob = new Blob([outputBuffer], {
      type:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    setDownloadUrl(url);
    alert('변환 완료되었습니다. 다운로드 버튼을 클릭하세요!');
  } catch (error) {
    console.error(error);
    alert('변환 중 오류가 발생했습니다.');
  } finally {
    setIsLoading(false);
  }
};



  // 다운로드
  const handleDownload = () => {
    if (!downloadUrl) return;
    const now = new Date();

    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0'); // 월은 0부터 시작하므로 +1
    const day = String(now.getDate()).padStart(2, '0');

    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const seconds = String(now.getSeconds()).padStart(2, '0');

    const formattedDateTime = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;

    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = `${formattedDateTime}.xlsx`; 
    link.click();
  };

  return (
    <div className="relative max-w-xl mx-auto py-8">
      {/* 전체 화면 로딩 오버레이 */}
      {isLoading && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50">
          <div
            className="flex flex-col items-center p-8 rounded-md shadow-xl"
            style={{
              background: 'linear-gradient(135deg, #ffffff 0%, #f0f4f8 100%)',
            }}
          >
            <span className="inline-block h-16 w-16 animate-spin rounded-full border-8 border-purple-500 border-t-transparent mb-6"></span>
            <p className="text-lg text-gray-800 font-semibold">처리 중...</p>
          </div>
        </div>
      )}

      <h1 className="text-2xl font-bold mb-6 text-center">엑셀 변환</h1>

      {/* 드래그앤드롭 영역 */}
      <div
        className="w-full p-6 mb-4 text-center border-2 border-dashed border-gray-300 rounded-lg hover:border-blue-300 transition-colors cursor-pointer"
        onDragOver={handleDragOver}
        onDrop={handleDrop}
      >
        {file ? (
          <p className="text-gray-700 font-medium">
            {file.name} (선택됨)
          </p>
        ) : (
          <p className="text-gray-500">
            이 영역에 파일을 드래그&드롭 하거나,<br />
            아래 버튼으로 파일을 선택하세요.
          </p>
        )}
      </div>

      {/* 파일 선택 & 제거 버튼들 */}
      <div className="flex items-center gap-2 mb-4">
        <label
          htmlFor="excel-file"
          className="inline-block px-4 py-2 text-white bg-blue-500 rounded-md cursor-pointer hover:bg-blue-600 transition-colors"
        >
          파일 선택
        </label>
        <input
          id="excel-file"
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileChange}
          className="hidden"
        />

        {file && (
          <button
            onClick={handleRemoveFile}
            className="px-4 py-2 bg-red-500 text-white rounded-md hover:bg-red-600 transition-colors"
          >
            파일 제거
          </button>
        )}
      </div>

      {/* 업로드 & 변환 버튼 */}
      <button
        onClick={() => processFile(file as File)}
        className="px-4 py-2 mr-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors"
      >
        변환
      </button>

      {/* 다운로드 버튼 */}
      {downloadUrl && (
        <button
          onClick={handleDownload}
          className="px-4 py-2 bg-indigo-500 text-white rounded-md hover:bg-indigo-600 transition-colors"
        >
          다운로드
        </button>
      )}
    </div>
  );
}

// 유틸리티 함수들
const getDistinctYearPrefixes = (data: any[]): string[] => {
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

const numberToColumnLetter = (num: number) => {
  let letter = '';
  while (num > 0) {
    const mod = (num - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    num = Math.floor((num - mod) / 26);
  }
  return letter;
};

const autoAdjustColumnWidths = (worksheet: Worksheet) => {
  worksheet.columns.forEach((column: any) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell: any) => {
      let cellValue = cell.value;
      if (cellValue === null || cellValue === undefined) {
        cellValue = '';
      } else if (typeof cellValue === 'object' && cellValue.richText) {
        cellValue = cellValue.richText.map((part: any) => part.text).join('');
      } else {
        cellValue = cellValue.toString();
      }
      maxLength = Math.max(maxLength, cellValue.length);
    });
    column.width = maxLength < 10 ? 10 : maxLength + 2;
  });
};
