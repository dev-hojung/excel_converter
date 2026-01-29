import { useConvertPriceExcel } from "@/hooks/useConvertPriceExcel";

/**
 * 가격 변환 컴포넌트
 * @description 모델명, 물품대(변경후), 판매가(변경후) 컬럼의 | 구분자 데이터를 행 단위로 분리 변환하는 UI
 */
export const PriceConverter = () => {
  const {
    file,
    downloadUrl,
    isLoading,
    processFile,
    handleDownload,
    handleDragOver,
    handleDrop,
    handleFileChange,
    handleRemoveFile,
  } = useConvertPriceExcel();

  return (
    <div>
      {/* 전체 화면 로딩 오버레이 */}
      {isLoading && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50">
          <div
            className="flex flex-col items-center p-8 rounded-md shadow-xl"
            style={{
              background: "linear-gradient(135deg, #ffffff 0%, #f0f4f8 100%)",
            }}
          >
            <span className="inline-block h-16 w-16 animate-spin rounded-full border-8 border-purple-500 border-t-transparent mb-6"></span>
            <p className="text-lg text-gray-800 font-semibold">처리 중...</p>
          </div>
        </div>
      )}

      {/* 안내 문구 */}
      <div className="mb-4 p-4 bg-blue-50 rounded-lg border border-blue-200">
        <p className="text-sm text-blue-800">
          <strong>변환 대상:</strong> 모델명, 물품대(변경후), 판매가(변경후)
          컬럼
          <br />
          <strong>변환 방식:</strong> | 구분자로 분리된 데이터를 각 행으로 분리
        </p>
      </div>

      {/* 드래그앤드롭 영역 */}
      <div
        className="w-full p-6 mb-4 text-center border-2 border-dashed border-gray-300 rounded-lg hover:border-blue-300 transition-colors cursor-pointer"
        onDragOver={handleDragOver}
        onDrop={handleDrop}
      >
        {file ? (
          <p className="text-gray-700 font-medium">{file.name} (선택됨)</p>
        ) : (
          <p className="text-gray-500">
            이 영역에 파일을 드래그&드롭 하거나,
            <br />
            아래 버튼으로 파일을 선택하세요.
          </p>
        )}
      </div>

      {/* 파일 선택 및 제거 버튼 */}
      <div className="flex items-center gap-2 mb-4">
        <label
          htmlFor="excel-file-price"
          className="inline-block px-4 py-2 text-white bg-blue-500 rounded-md cursor-pointer hover:bg-blue-600 transition-colors"
        >
          파일 선택
        </label>
        <input
          id="excel-file-price"
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

      {/* 변환 & 다운로드 버튼 */}
      <button
        onClick={() => file && processFile(file)}
        className="px-4 py-2 mr-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors"
      >
        변환
      </button>

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
};
