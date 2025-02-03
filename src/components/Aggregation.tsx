import { useConvertAggregationExcel } from "@/hooks/useConvertAggregationExcelYear";
import { useEffect, useState } from "react";

export const Aggregation = () => {
  const [selectedOption, setSelectedOption] = useState<'yearly' | 'monthly'>('yearly')
  

  const {
    file: aggregationFile,
    downloadUrl: aggregationDownloadUrl,
    isLoading: aggregationIsLoading,
    processFile: processAggregationFile,
    handleDownload: handleAggregationDownload,
    handleDragOver: handleAggregationDragOver,
    handleDrop: handleAggregationDrop,
    handleFileChange: handleAggregationFileChange,
    handleRemoveFile: handleAggregationRemoveFile,
    setDownloadUrl,
  } = useConvertAggregationExcel({ tab: selectedOption });
  

  useEffect(() => {
    setDownloadUrl('')
  }, [selectedOption])

  return (
    <div>
      {/* 전체 화면 로딩 오버레이 */}
      {aggregationIsLoading && (
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


      {/* 라디오 버튼 (월별, 연도별) */}
      <div className="flex gap-6 mb-6 justify-center">
        {/* 월별 */}
        <label className="flex items-center gap-2 cursor-pointer">
          <input
            type="radio"
            name="aggregationType"
            value="monthly"
            checked={selectedOption === "monthly"}
            onChange={() => setSelectedOption("monthly")}
            className="hidden"
          />
          <div
            className={`w-6 h-6 rounded-full border-2 flex items-center justify-center transition-all ${
              selectedOption === "monthly"
                ? "border-blue-500 bg-blue-500"
                : "border-gray-400"
            }`}
          >
            {selectedOption === "monthly" && <div className="w-3 h-3 bg-white rounded-full"></div>}
          </div>
          <span className="text-gray-700 font-medium">월별</span>
        </label>

        {/* 연도별 */}
        <label className="flex items-center gap-2 cursor-pointer">
          <input
            type="radio"
            name="aggregationType"
            value="yearly"
            checked={selectedOption === "yearly"}
            onChange={() => setSelectedOption("yearly")}
            className="hidden"
          />
          <div
            className={`w-6 h-6 rounded-full border-2 flex items-center justify-center transition-all ${
              selectedOption === "yearly"
                ? "border-blue-500 bg-blue-500"
                : "border-gray-400"
            }`}
          >
            {selectedOption === "yearly" && <div className="w-3 h-3 bg-white rounded-full"></div>}
          </div>
          <span className="text-gray-700 font-medium">연도별</span>
        </label>
      </div>
      {/* 드래그앤드롭 영역 */}
      <div
        className="w-full p-6 mb-4 text-center border-2 border-dashed border-gray-300 rounded-lg hover:border-blue-300 transition-colors cursor-pointer"
        onDragOver={handleAggregationDragOver}
        onDrop={handleAggregationDrop}
      >
        {aggregationFile ? (
          <p className="text-gray-700 font-medium">{aggregationFile.name} (선택됨)</p>
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
          htmlFor="excel-file-aggregation"
          className="inline-block px-4 py-2 text-white bg-blue-500 rounded-md cursor-pointer hover:bg-blue-600 transition-colors"
        >
          파일 선택
        </label>
        <input
          id="excel-file-aggregation"
          type="file"
          accept=".xlsx, .xls"
          onChange={handleAggregationFileChange}
          className="hidden"
        />
        {aggregationFile && (
          <button
            onClick={handleAggregationRemoveFile}
            className="px-4 py-2 bg-red-500 text-white rounded-md hover:bg-red-600 transition-colors"
          >
            파일 제거
          </button>
        )}
      </div>

      {/* 변환 & 다운로드 버튼 */}
      <button
        onClick={() => aggregationFile && processAggregationFile(aggregationFile)}
        className="px-4 py-2 mr-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors"
      >
        변환
      </button>

      {aggregationDownloadUrl && (
        <button
          onClick={handleAggregationDownload}
          className="px-4 py-2 bg-indigo-500 text-white rounded-md hover:bg-indigo-600 transition-colors"
        >
          다운로드
        </button>
      )}
    </div>
  );
}