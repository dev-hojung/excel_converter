/* eslint-disable @typescript-eslint/no-explicit-any */
import React, { DragEvent, useState } from "react";
import ExcelJS from "exceljs";
import { autoAdjustColumnWidths } from "@/utils";

/**
 * 집계된 수량 데이터의 단일 항목 타입
 * @property {string} model - 모델명
 * @property {string} deliveryType - 배송타입
 * @property {number} quantity - 합산된 출하수량
 */
type AggregatedItem = {
  model: string;
  deliveryType: string;
  quantity: number;
};

/**
 * 헤더 컬럼 인덱스 탐지 결과 타입
 * @property {number} modelCol - 모델명 컬럼 인덱스 (1-based)
 * @property {number} deliveryTypeCol - 배송타입 컬럼 인덱스 (1-based)
 * @property {number} quantityCol - 출하수량 컬럼 인덱스 (1-based)
 */
type ColumnIndices = {
  modelCol: number;
  deliveryTypeCol: number;
  quantityCol: number;
};

/**
 * 수량 집계 엑셀 변환 커스텀 훅
 * @description 동일 모델 + 동일 배송타입 기준으로 출하수량을 합산하고,
 *              집계 전/후 총합 검증까지 수행하여 새로운 엑셀 파일로 생성합니다.
 */
export const useConvertQuantity = () => {
  const [file, setFile] = useState<File | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string>("");
  const [isLoading, setIsLoading] = useState(false);

  /**
   * 파일 변경 핸들러
   * @param {React.ChangeEvent<HTMLInputElement>} e - 파일 입력 이벤트
   */
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.length) {
      setFile(e.target.files[0]);
    }
    e.target.value = "";
  };

  /**
   * 드롭 핸들러
   * @param {DragEvent<HTMLDivElement>} e - 드래그앤드롭 이벤트
   */
  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (e.dataTransfer.files?.length) {
      setFile(e.dataTransfer.files[0]);
    }
  };

  /**
   * 드래그 오버 핸들러
   * @param {DragEvent<HTMLDivElement>} e - 드래그 오버 이벤트
   */
  const handleDragOver = (e: DragEvent<HTMLDivElement>) => e.preventDefault();

  /**
   * 파일 제거 핸들러
   * 선택된 파일과 다운로드 URL을 초기화합니다.
   */
  const handleRemoveFile = () => {
    setFile(null);
    setDownloadUrl("");
  };

  /**
   * 엑셀 파일 처리 프로세스 시작
   * @param {File} file - 변환할 엑셀 파일
   */
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
      alert(`변환 중 오류가 발생했습니다.\n${(error as Error).message}`);
    } finally {
      setIsLoading(false);
    }
  };

  /**
   * ExcelJS 셀 값에서 순수 텍스트를 추출하는 헬퍼 함수
   * richText 객체(스타일이 적용된 셀 값)와 일반 텍스트 모두 처리합니다.
   * @param {ExcelJS.CellValue} cellValue - ExcelJS 셀의 raw 값
   * @returns {string} 추출된 텍스트 문자열 (공백 제거됨)
   */
  const extractCellText = (cellValue: ExcelJS.CellValue): string => {
    if (cellValue === null || cellValue === undefined) return "";
    if (typeof cellValue === "object" && "richText" in (cellValue as any)) {
      return (cellValue as any).richText
        .map((part: any) => part.text)
        .join("")
        .trim();
    }
    return String(cellValue).trim();
  };

  /**
   * 워크시트의 1행(헤더행)을 분석하여 필수 컬럼의 인덱스를 자동 탐지합니다.
   * "모델명", "배송타입", "출하수량" 키워드를 포함한 셀을 찾습니다.
   * @param {ExcelJS.Worksheet} worksheet - 분석할 워크시트
   * @returns {ColumnIndices} 각 컬럼의 1-based 인덱스
   * @throws {Error} 필수 컬럼 중 하나라도 찾지 못한 경우
   */
  const detectColumnIndices = (worksheet: ExcelJS.Worksheet): ColumnIndices => {
    const headerRow = worksheet.getRow(1);
    let modelCol = -1;
    let deliveryTypeCol = -1;
    let quantityCol = -1;

    headerRow.eachCell((cell, colNumber) => {
      const header = extractCellText(cell.value).toLowerCase();
      if (header.includes("모델")) modelCol = colNumber;
      if (header.includes("배송")) deliveryTypeCol = colNumber;
      if (header.includes("출하수량") || header.includes("수량"))
        quantityCol = colNumber;
    });

    const missing: string[] = [];
    if (modelCol === -1) missing.push("모델명");
    if (deliveryTypeCol === -1) missing.push("배송타입");
    if (quantityCol === -1) missing.push("출하수량");

    if (missing.length > 0) {
      throw new Error(
        `헤더에서 다음 컬럼을 찾을 수 없습니다: [${missing.join(", ")}]\n` +
          `1행에 "모델명", "배송타입", "출하수량" 헤더가 있는지 확인하세요.`,
      );
    }

    console.log(
      `[수량집계] 컬럼 탐지 완료 — 모델명: ${modelCol}열, 배송타입: ${deliveryTypeCol}열, 출하수량: ${quantityCol}열`,
    );
    return { modelCol, deliveryTypeCol, quantityCol };
  };

  /**
   * 집계 전/후 출하수량 총합을 비교하여 데이터 정합성을 검증합니다.
   * 원본 총합과 집계 후 총합이 다르면 에러를 발생시킵니다.
   * @param {number} originalTotal - 원본 데이터의 출하수량 총합
   * @param {AggregatedItem[]} aggregatedResult - 집계 후 데이터 배열
   * @throws {Error} 총합이 일치하지 않는 경우
   */
  const validateAggregation = (
    originalTotal: number,
    aggregatedResult: AggregatedItem[],
  ): void => {
    const aggregatedTotal = aggregatedResult.reduce(
      (sum, item) => sum + item.quantity,
      0,
    );

    console.log(
      `[수량집계 검증] 원본 총 출하수량: ${originalTotal.toLocaleString()}`,
    );
    console.log(
      `[수량집계 검증] 집계 후 총 출하수량: ${aggregatedTotal.toLocaleString()}`,
    );

    if (originalTotal !== aggregatedTotal) {
      throw new Error(
        `[수량집계 검증 실패] 원본 총합(${originalTotal.toLocaleString()})과 ` +
          `집계 후 총합(${aggregatedTotal.toLocaleString()})이 일치하지 않습니다.`,
      );
    }

    console.log("[수량집계 검증] 검증 통과 ✅");
  };

  /**
   * 워크시트를 읽어 동일 모델 + 동일 배송타입 기준으로 출하수량을 집계합니다.
   * 헤더행(1행)에서 컬럼 인덱스를 자동 탐지하고, 집계 후 검증까지 수행합니다.
   * @param {ExcelJS.Worksheet} worksheet - 원본 데이터 워크시트
   * @returns {{ result: AggregatedItem[]; originalTotal: number }} 집계 결과 및 원본 총합
   */
  const aggregateByModelAndDeliveryType = (
    worksheet: ExcelJS.Worksheet,
  ): { result: AggregatedItem[]; originalTotal: number } => {
    const { modelCol, deliveryTypeCol, quantityCol } =
      detectColumnIndices(worksheet);

    // 레코드 키: "모델명_배송타입" 복합키로 집계 맵 구성
    const aggregatedMap: Record<string, AggregatedItem> = {};
    let originalTotal = 0;

    worksheet.eachRow((row, rowNumber) => {
      // 1행(헤더)은 스킵
      if (rowNumber === 1) return;

      const model = extractCellText(row.getCell(modelCol).value);
      const deliveryType = extractCellText(row.getCell(deliveryTypeCol).value);
      const quantity = Number(row.getCell(quantityCol).value) || 0;

      // 모델명이 비어있는 행은 빈 행으로 간주하여 스킵
      if (!model) return;

      originalTotal += quantity;

      const key = `${model}__${deliveryType}`;
      if (!aggregatedMap[key]) {
        aggregatedMap[key] = { model, deliveryType, quantity: 0 };
      }
      aggregatedMap[key].quantity += quantity;
    });

    // 모델명 오름차순, 동일 모델일 경우 배송타입 오름차순 정렬
    const result = Object.values(aggregatedMap).sort((a, b) => {
      if (a.model === b.model) {
        return a.deliveryType.localeCompare(b.deliveryType);
      }
      return a.model.localeCompare(b.model);
    });

    // 집계 전/후 총합 정합성 검증
    validateAggregation(originalTotal, result);

    return { result, originalTotal };
  };

  /**
   * 집계 결과를 새로운 ExcelJS 워크시트에 포맷팅하여 작성합니다.
   * 헤더 스타일, 데이터 행, 숫자 포맷(#,##0)을 적용합니다.
   * @param {AggregatedItem[]} aggregatedResult - 집계된 데이터 배열
   * @returns {ExcelJS.Workbook} 포맷팅이 완료된 새 워크북
   */
  const buildFormattedWorkbook = (
    aggregatedResult: AggregatedItem[],
  ): ExcelJS.Workbook => {
    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet("Converted");

    // 헤더 행 추가 및 스타일 적용
    newWorksheet.addRow(["모델명", "배송타입", "출하수량"]).eachCell((cell) => {
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

    // 데이터 행 추가: 출하수량 셀에 천단위 콤마 포맷 적용
    aggregatedResult.forEach(({ model, deliveryType, quantity }) => {
      const row = newWorksheet.addRow([model, deliveryType, quantity]);

      // C열(출하수량): 숫자 포맷 "#,##0" 적용 (예: 1,234)
      const quantityCell = row.getCell(3);
      quantityCell.numFmt = "#,##0";
    });

    // 열 너비 자동 조정
    autoAdjustColumnWidths(newWorksheet);

    return newWorkbook;
  };

  /**
   * 엑셀 파일을 읽어 집계 → 포맷팅 → Blob URL 생성까지 수행합니다.
   * @param {File} file - 변환할 엑셀 파일
   * @returns {Promise<string>} 변환된 파일의 다운로드 Blob URL
   */
  const convertExcelFile = async (file: File): Promise<string> => {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      throw new Error("워크시트를 찾을 수 없습니다.");
    }

    // 1. 동일 모델 + 동일 배송타입 기준으로 출하수량 집계 및 검증
    const { result: aggregatedResult } =
      aggregateByModelAndDeliveryType(worksheet);

    // 2. 집계 결과를 포맷팅하여 새 워크북 생성
    const newWorkbook = buildFormattedWorkbook(aggregatedResult);

    // 3. Blob URL 생성하여 반환
    const outputBuffer = await newWorkbook.xlsx.writeBuffer();
    const blob = new Blob([outputBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    return URL.createObjectURL(blob);
  };

  /**
   * 다운로드 버튼 클릭 핸들러
   * 현재 날짜/시간을 파일명에 포함하여 다운로드합니다.
   */
  const handleDownload = () => {
    if (!downloadUrl) return;

    const now = new Date();
    const formattedDateTime = `${now.getFullYear()}-${String(
      now.getMonth() + 1,
    ).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")} ${String(
      now.getHours(),
    ).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}:${String(
      now.getSeconds(),
    ).padStart(2, "0")}`;

    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = `수량집계_${formattedDateTime}.xlsx`;
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
    setDownloadUrl,
  };
};
