import { Aggregation } from "@/components/Aggregation";
import { PriceConverter } from "@/components/PriceConverter";
import { QuantityAggregation } from "@/components/QuantityAggregation";
import React, { useState } from "react";

export default function UploadPage() {
  // 현재 활성화된 탭 상태 (집계 / 수량집계 / 가격변환)
  const [activeTab, setActiveTab] = useState<
    "집계" | "동일모델 수량집계" | "가격변환"
  >("집계");

  return (
    <div className="relative max-w-xl mx-auto py-8">
      <h1 className="text-2xl font-bold mb-6 text-center">엑셀 변환</h1>
      {/* 탭 헤더 */}
      <div className="flex mb-4 bg-white rounded-lg shadow overflow-hidden">
        <button
          onClick={() => setActiveTab("집계")}
          className={`flex-1 py-3 text-center font-semibold transition-colors duration-300 ${
            activeTab === "집계"
              ? "bg-blue-500 text-white"
              : "bg-gray-200 text-gray-800 hover:bg-gray-300"
          }`}
        >
          집계
        </button>
        <button
          onClick={() => setActiveTab("동일모델 수량집계")}
          className={`flex-1 py-3 text-center font-semibold transition-colors duration-300 ${
            activeTab === "동일모델 수량집계"
              ? "bg-blue-500 text-white"
              : "bg-gray-200 text-gray-800 hover:bg-gray-300"
          }`}
        >
          수량집계
        </button>
        <button
          onClick={() => setActiveTab("가격변환")}
          className={`flex-1 py-3 text-center font-semibold transition-colors duration-300 ${
            activeTab === "가격변환"
              ? "bg-blue-500 text-white"
              : "bg-gray-200 text-gray-800 hover:bg-gray-300"
          }`}
        >
          가격변환
        </button>
      </div>
      {activeTab === "집계" && <Aggregation />}
      {/* 수량집계 탭 컨텐츠 */}
      {activeTab === "동일모델 수량집계" && <QuantityAggregation />}
      {/* 가격변환 탭 컨텐츠 */}
      {activeTab === "가격변환" && <PriceConverter />}
    </div>
  );
}
