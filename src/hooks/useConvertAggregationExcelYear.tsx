/* eslint-disable @typescript-eslint/no-explicit-any */
import React, { DragEvent, useState } from "react";
import ExcelJS from "exceljs";
import {
  autoAdjustColumnWidths,
  getDistinctYearlyPrefixes,
  getDistinctMonthlyPrefixes,
  numberToColumnLetter,
} from "@/utils";

export const useConvertAggregationExcel = ({ tab }: { tab: 'yearly' | 'monthly' }) => {
  const [file, setFile] = useState<File | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string>("");
  const [isLoading, setIsLoading] = useState(false);

   /** 파일 선택, 드래그/드롭, 제거 핸들러 */
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.length) {
      setFile(e.target.files[0]);
    }
    e.target.value = "";
  };

  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (e.dataTransfer.files?.length) {
      setFile(e.dataTransfer.files[0]);
    }
  };

  const handleDragOver = (e: DragEvent<HTMLDivElement>) => e.preventDefault();

  const handleRemoveFile = () => {
    setFile(null);
    setDownloadUrl("");
  };

  /** 엑셀 변환 프로세스 */
  const processFile = async (file: File) => {
    if (!file) {
      alert("파일을 먼저 선택(또는 드래그)하세요.");
      return;
    }

    setIsLoading(true);
    try {
      const url = await convertExcelFile(file);
      setDownloadUrl(url);
      alert("변환 완료되었습니다. 다운로드 버튼을 클릭하세요!");
    } catch (error) {
      console.error(error);
      alert("변환 중 오류가 발생했습니다.");
    } finally {
      setIsLoading(false);
    }
  };

  /** 엑셀 파일 읽기 → 데이터 추출 → 그룹별 집계 → 새로운 워크북 생성 후 Blob URL 반환 */
  const convertExcelFile = async (file: File): Promise<string> => {
    // 1. 엑셀 파일 로드
    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      throw new Error("워크시트를 찾을 수 없습니다.");
    }

    // 2. 데이터 추출 (헤더 포함)
    const originalData = extractDataFromWorksheet(worksheet);
    const dataWithoutHeader = originalData.slice(1);

    // // 3. 집계에 필요한 값 준비
    const distinct = tab === 'yearly' ?  getDistinctYearlyPrefixes(dataWithoutHeader) : getDistinctMonthlyPrefixes(dataWithoutHeader);
    const aggregatedResults = aggregateData(dataWithoutHeader, distinct);

    // // 4. 새 워크북 생성
    const newWorkbook = createConvertedWorkbook(aggregatedResults, distinct);
    autoAdjustColumnWidths(newWorkbook.getWorksheet("Converted")!);

    // 5. Blob URL 생성
    const outputBuffer = await newWorkbook.xlsx.writeBuffer();
    const blob = new Blob([outputBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    return URL.createObjectURL(blob);
  };

  /** 워크시트에서 모든 데이터를 추출 */
  const extractDataFromWorksheet = (worksheet: ExcelJS.Worksheet): any[] => {
    const originalData: any[] = [];
    worksheet.eachRow((row) => {
      const rowObj: any = {};
      row.eachCell((cell, colNumber) => {
        const colLetter = worksheet.getColumn(colNumber).letter;
        rowObj[colLetter] = cell.value;
      });
      originalData.push(rowObj);
    });
    return originalData;
  };

  /** 코드별 그룹 집계 (필요한 값: A, B, 각 연도별 합계, 총 합계) */
  const aggregateData = (data: any[], distinct: string[]): any[] => {
    const codeArr = data.map((item) => item.C);
    const uniqueCodes = Array.from(new Set(codeArr));

    return uniqueCodes
      .map((code) => {
        const filtered = data.filter((item) => item.C === code);
        const A = filtered[0].D;
        const B = code;
        const startChar = "C";
        const result: { [key: string]: any } = {};

        distinct.forEach((distinctItem, index) => {
          const key = String.fromCharCode(startChar.charCodeAt(0) + index);
          result[key] = filtered
            .filter(
              (item) => tab === 'yearly' ? `20${item.J?.substring(0, 2)}` : `${item.J?.substring(0, 2)}년_${item.J?.substring(2, 4)}월` === distinctItem
            )
            .reduce((acc, curr) => acc + (curr.H || 0), 0);
        });

        const endCode = String.fromCharCode(
          startChar.charCodeAt(0) + distinct.length
        );
        result[endCode] = Object.values(result).reduce(
          (acc, curr) => acc + curr,
          0
        );

        return { A, B, ...result };
      })
      .sort((a, b) => String(a.A).localeCompare(String(b.A)));
  };

  /** 새 워크북 생성: 헤더, 데이터 행, 그룹별 요약 행 추가 */
  const createConvertedWorkbook = (
    aggregatedResults: any[],
    distinct: string[]
  ): ExcelJS.Workbook => {
    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet("Converted");

    const header = ["제품군", "모델명", ...distinct, "총 합계"];
    // 헤더 행 추가 및 스타일 적용
    newWorksheet.addRow(header).eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "d8eef2" },
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cell.font = { bold: true };
    });

    let currentGroup: string | null = null;
    let groupStartRow = newWorksheet.rowCount + 1;

    aggregatedResults.forEach((dataRow) => {
      // 그룹이 바뀌면 이전 그룹의 요약 행 추가
      if (currentGroup && dataRow.A !== currentGroup) {
        addSummaryRow(newWorksheet, header, currentGroup, groupStartRow);
        currentGroup = dataRow.A;
        groupStartRow = newWorksheet.rowCount + 1;
      } else if (!currentGroup) {
        currentGroup = dataRow.A;
      }
      newWorksheet.addRow(Object.values(dataRow));
    });

    // 마지막 그룹 요약 행 추가
    if (currentGroup) {
      addSummaryRow(newWorksheet, header, currentGroup, groupStartRow);
    }

    return newWorkbook;
  };

  /** 그룹별 요약 행 추가 (합계 계산 및 스타일 적용) */
  const addSummaryRow = (
    worksheet: ExcelJS.Worksheet,
    header: string[],
    groupName: string,
    groupStartRow: number
  ) => {
    const currentRow = worksheet.rowCount + 1;
    const summaryRow = worksheet.addRow([`${groupName} 계`]);

    for (let col = 3; col <= header.length; col++) {
      const colLetter = numberToColumnLetter(col);
      const formula = `SUM(${colLetter}${groupStartRow}:${colLetter}${
        currentRow - 1
      })`;
      summaryRow.getCell(col).value = { formula };
    }

    worksheet.mergeCells(`A${summaryRow.number}:B${summaryRow.number}`);
    summaryRow.getCell(1).alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    summaryRow.eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ffe9da" },
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } },
      };
      cell.font = { bold: true };
    });
  };

  /** 다운로드 버튼 클릭 핸들러 */
  const handleDownload = () => {
    if (!downloadUrl) return;

    const now = new Date();
    const formattedDateTime = `${now.getFullYear()}-${String(
      now.getMonth() + 1
    ).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")} ${String(
      now.getHours()
    ).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}:${String(
      now.getSeconds()
    ).padStart(2, "0")}`;

    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = `${tab === 'monthly' ? '월별_' : '년도별_'}${formattedDateTime}.xlsx`;
    link.click();
  };

  return {
    file,
    downloadUrl,
    isLoading,
    handleDragOver,
    handleDrop,
    handleFileChange,
    handleRemoveFile,
    handleDownload,
    processFile,
    setDownloadUrl
  }
}