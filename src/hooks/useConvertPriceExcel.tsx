import React, { DragEvent, useState } from "react";
import ExcelJS from "exceljs";
import { autoAdjustColumnWidths } from "@/utils";

/**
 * 다양한 수직선(파이프) 문자 변형을 모두 인식하여 분리
 * @description 엑셀에서 입력 가능한 다양한 유니코드 수직선 문자를 모두 처리
 * - | (U+007C) - 일반 파이프
 * - ¦ (U+00A6) - Broken Bar
 * - │ (U+2502) - Box Drawings Light Vertical
 * - ┃ (U+2503) - Box Drawings Heavy Vertical
 * - ∣ (U+2223) - Divides
 * - ｜ (U+FF5C) - Fullwidth Vertical Line
 * @param text - 분리할 텍스트
 * @returns 분리된 문자열 배열
 */
const splitByPipeVariants = (text: string): string[] => {
  // 모든 수직선 문자 변형을 정규식으로 매칭
  const pipeRegex = /[│┃]/;
  return text.split(pipeRegex).map((s) => s.trim());
};

/**
 * ExcelJS 셀 값을 문자열로 변환
 * @description RichText, 객체, 숫자 등 다양한 형식을 문자열로 안전하게 변환
 * @param cellValue - ExcelJS 셀 값 (CellValue 타입)
 * @returns 문자열 값
 */
const getCellTextValue = (cellValue: ExcelJS.CellValue): string => {
  if (cellValue === null || cellValue === undefined) {
    return "";
  }

  // RichText 형식인 경우 (객체에 richText 배열이 있음)
  if (typeof cellValue === "object" && "richText" in cellValue) {
    const richTextValue = cellValue as ExcelJS.CellRichTextValue;
    return richTextValue.richText.map((rt) => rt.text).join("");
  }

  // 숫자, 문자열 등 기본 타입
  return String(cellValue);
};
/**
 * 가격 변환 엑셀 훅
 * @description 모델명, 물품대(변경후), 판매가(변경후) 컬럼의 | 구분자 데이터를 행 단위로 분리 변환
 * @returns 파일 처리 및 다운로드 관련 상태와 핸들러
 */
export const useConvertPriceExcel = () => {
  const [file, setFile] = useState<File | null>(null);
  const [downloadUrl, setDownloadUrl] = useState<string>("");
  const [isLoading, setIsLoading] = useState(false);

  /**
   * 파일 선택 핸들러
   * @param e - 파일 입력 이벤트
   */
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.length) {
      setFile(e.target.files[0]);
    }
    e.target.value = "";
  };

  /**
   * 드래그앤드롭 핸들러
   * @param e - 드래그 이벤트
   */
  const handleDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (e.dataTransfer.files?.length) {
      setFile(e.dataTransfer.files[0]);
    }
  };

  /**
   * 드래그 오버 핸들러
   * @param e - 드래그 이벤트
   */
  const handleDragOver = (e: DragEvent<HTMLDivElement>) => e.preventDefault();

  /**
   * 파일 제거 핸들러
   */
  const handleRemoveFile = () => {
    setFile(null);
    setDownloadUrl("");
  };

  /**
   * 엑셀 변환 프로세스 실행
   * @param file - 변환할 엑셀 파일
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
      alert("변환 중 오류가 발생했습니다.");
    } finally {
      setIsLoading(false);
    }
  };

  /**
   * 엑셀 파일 변환 메인 로직
   * @description
   * 1. 원본 엑셀 파일 로드
   * 2. 모델명, 물품대(변경후), 판매가(변경후) 컬럼 찾기
   * 3. | 구분자로 분리된 데이터를 각 행으로 분리
   * 4. 새로운 워크북 생성 및 Blob URL 반환
   * @param file - 변환할 엑셀 파일
   * @returns Blob URL
   */
  const convertExcelFile = async (file: File): Promise<string> => {
    // 1. 엑셀 파일 로드
    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      throw new Error("워크시트를 찾을 수 없습니다.");
    }

    // 2. 헤더 행에서 컬럼 인덱스 찾기
    const headerRow = worksheet.getRow(1);
    const columnIndexes = findColumnIndexes(headerRow);

    if (
      !columnIndexes.modelName ||
      !columnIndexes.price ||
      !columnIndexes.salePrice
    ) {
      throw new Error(
        "필수 컬럼(모델명, 물품대(변경후), 판매가(변경후))을 찾을 수 없습니다.",
      );
    }

    // 3. 데이터 추출 및 변환
    const convertedData = extractAndConvertData(worksheet, columnIndexes);

    // 4. 새 워크북 생성
    const newWorkbook = createConvertedWorkbook(convertedData);
    autoAdjustColumnWidths(newWorkbook.getWorksheet("Converted")!);

    // 5. Blob URL 생성
    const outputBuffer = await newWorkbook.xlsx.writeBuffer();
    const blob = new Blob([outputBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    return URL.createObjectURL(blob);
  };

  /**
   * 헤더 행에서 필수 컬럼 인덱스 찾기
   * @param headerRow - 헤더 행
   * @returns 각 컬럼의 인덱스 객체
   */
  const findColumnIndexes = (
    headerRow: ExcelJS.Row,
  ): {
    modelName: number | null;
    price: number | null;
    salePrice: number | null;
  } => {
    const result = {
      modelName: null as number | null,
      price: null as number | null,
      salePrice: null as number | null,
    };

    headerRow.eachCell((cell, colNumber) => {
      const cellValue = getCellTextValue(cell.value).trim();
      if (cellValue === "모델명") {
        result.modelName = colNumber;
      } else if (cellValue === "물품대(변경후)") {
        result.price = colNumber;
      } else if (cellValue === "판매가(변경후)") {
        result.salePrice = colNumber;
      }
    });

    return result;
  };

  /**
   * 워크시트에서 데이터 추출 및 | 구분자로 분리 변환
   * @description 각 행의 모델명, 물품대, 판매가 값을 | 기준으로 분리하여 1:1 매칭
   * @param worksheet - 원본 워크시트
   * @param columnIndexes - 컬럼 인덱스 정보
   * @returns 변환된 데이터 배열 [{modelName, price, salePrice}, ...]
   */
  const extractAndConvertData = (
    worksheet: ExcelJS.Worksheet,
    columnIndexes: {
      modelName: number | null;
      price: number | null;
      salePrice: number | null;
    },
  ): { modelName: string; price: string; salePrice: string }[] => {
    const result: { modelName: string; price: string; salePrice: string }[] =
      [];

    // 헤더 행(1) 제외하고 데이터 행 순회
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // 헤더 스킵

      // 각 컬럼 값 추출 (getCellTextValue로 RichText 등 처리)
      const modelNameValue = getCellTextValue(
        row.getCell(columnIndexes.modelName!).value,
      );
      const priceValue = getCellTextValue(
        row.getCell(columnIndexes.price!).value,
      );
      const salePriceValue = getCellTextValue(
        row.getCell(columnIndexes.salePrice!).value,
      );

      // 다양한 수직선 문자 변형(|, │, ┃ 등)으로 분리
      const modelNames = splitByPipeVariants(modelNameValue);
      const prices = splitByPipeVariants(priceValue);
      const salePrices = splitByPipeVariants(salePriceValue);

      // 가장 긴 배열 기준으로 1:1 매칭
      const maxLength = Math.max(
        modelNames.length,
        prices.length,
        salePrices.length,
      );

      for (let i = 0; i < maxLength; i++) {
        const modelName = modelNames[i] || "";
        const price = prices[i] || "";
        const salePrice = salePrices[i] || "";

        // 모든 값이 비어있으면 스킵
        if (!modelName && !price && !salePrice) continue;

        result.push({ modelName, price, salePrice });
      }
    });

    return result;
  };

  /**
   * 변환된 데이터로 새 워크북 생성
   * @param data - 변환된 데이터 배열
   * @returns 새로운 ExcelJS 워크북
   */
  const createConvertedWorkbook = (
    data: { modelName: string; price: string; salePrice: string }[],
  ): ExcelJS.Workbook => {
    const newWorkbook = new ExcelJS.Workbook();
    const newWorksheet = newWorkbook.addWorksheet("Converted");

    // 헤더 행 추가 및 스타일 적용
    const header = ["모델명", "물품대(변경후)", "판매가(변경후)"];
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

    // 데이터 행 추가
    data.forEach((item) => {
      newWorksheet.addRow([item.modelName, item.price, item.salePrice]);
    });

    return newWorkbook;
  };

  /**
   * 변환된 파일 다운로드 핸들러
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
    link.download = `가격변환_${formattedDateTime}.xlsx`;
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
